"""Run with: python -m unittest discover -s tests -v.

Dependency-free checks exercise the original functions through small adapters.
The PPTX tests additionally require python/requirements.txt.
"""
import argparse
import ast
from contextlib import redirect_stdout, redirect_stderr
from copy import deepcopy
from hashlib import md5
import importlib.util
import io
import json
import os
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'python'))


def load_functions(filename, names, namespace):
    """Load unchanged function ASTs without importing optional PPTX packages."""
    tree = ast.parse((ROOT / 'python' / filename).read_text(encoding='utf-8'))
    selected = [node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in names]
    exec(compile(ast.Module(body=selected, type_ignores=[]), filename, 'exec'), namespace)
    return namespace


class TemplateTextTests(unittest.TestCase):
    def setUp(self):
        self.ns = load_functions('template_engine.py', {
            '_normalize_bullets', '_pick_title_block', '_body_blocks', '_write_entries', 'rewrite_page',
        }, {
            '_text_blocks': lambda slide: slide.blocks,
            '_iter_shapes': lambda shapes: iter(shapes),
            '_set_block_text': lambda shape, text: setattr(shape, 'text', text),
            '_clear_block': lambda shape: setattr(shape, 'text', ''),
            '_shrink_to_fit': lambda *args: None,
        })

    def slide(self, body_count=1, number=False):
        texts = [('Old title', 36)] + [('Old body ' + str(i), 18) for i in range(body_count)]
        if number:
            texts.append(('01', 100))
        shapes = [SimpleNamespace(text=text, has_text_frame=True, has_table=False) for text, _ in texts]
        return SimpleNamespace(shapes=shapes, blocks=[{'shape': shape, 'text': text, 'size': size}
            for shape, (text, size) in zip(shapes, texts)])

    def test_all_bullets_survive_one_body_box(self):
        slide = self.slide()
        self.ns['rewrite_page'](slide, {'title': 'New', 'bullets': ['First', 'Second', 'Third']}, 'content')
        self.assertEqual(slide.shapes[1].text.splitlines(), ['First', 'Second', 'Third'])

    def test_bullets_keep_order_across_boxes(self):
        slide = self.slide(2)
        self.ns['rewrite_page'](slide, {'bullets': ['First', 'Second', 'Third']}, 'content')
        self.assertEqual([line for shape in slide.shapes[1:] for line in shape.text.splitlines()], ['First', 'Second', 'Third'])

    def test_section_replaces_number_and_subtitle(self):
        slide = self.slide(number=True)
        self.ns['rewrite_page'](slide, {'title': 'New section', 'subtitle': 'New subtitle', 'number': '02'}, 'section')
        self.assertEqual([shape.text for shape in slide.shapes], ['New section', 'New subtitle', '02'])

    def test_unused_cover_and_toc_text_is_cleared(self):
        for kind in ('cover', 'toc', 'end', 'section'):
            with self.subTest(kind=kind):
                slide = self.slide(3)
                self.ns['rewrite_page'](slide, {'title': 'New'}, kind)
                self.assertEqual([shape.text for shape in slide.shapes], ['New', '', '', ''])

    def test_toc_overflow_survives(self):
        slide = self.slide()
        self.ns['rewrite_page'](slide, {'items': ['A', 'B', 'C']}, 'toc')
        self.assertEqual(slide.shapes[1].text.splitlines(), ['A', 'B', 'C'])

    def test_missing_body_box_is_an_explicit_error(self):
        with self.assertRaisesRegex(ValueError, 'no text box'):
            self.ns['rewrite_page'](self.slide(0), {'bullets': ['Must survive']}, 'content')


class BuildArgumentsTests(unittest.TestCase):
    def setUp(self):
        self.outline = None
        def build(outline, output):
            self.outline = outline
            return {'output': output}
        self.ns = load_functions('build.py', {'main'}, {
            'argparse': argparse, 'json': json, 'sys': sys,
            'THEMES': {'navy': {}, 'teal': {}}, 'build': build, '__doc__': 'build',
        })

    def test_theme_flag_with_stdin(self):
        with patch.object(sys, 'stdin', io.StringIO('{"slides": []}')), redirect_stdout(io.StringIO()):
            code = self.ns['main'](['build.py', '-', 'out.pptx', '--theme', 'teal'])
        self.assertEqual(code, 0)
        self.assertEqual(self.outline['theme'], 'teal')

    def test_missing_theme_value_fails(self):
        with redirect_stderr(io.StringIO()):
            self.assertEqual(self.ns['main'](['build.py', '-', 'out.pptx', '--theme']), 2)

    def test_invalid_outline_is_a_structured_error(self):
        output = io.StringIO()
        with patch.object(sys, 'stdin', io.StringIO('[]')), redirect_stdout(output):
            code = self.ns['main'](['build.py', '-', 'out.pptx'])
        self.assertEqual(code, 1)
        self.assertIn('JSON object', json.loads(output.getvalue())['error'])

    def test_missing_template_is_not_free_layout(self):
        ns = load_functions('build.py', {'_build'}, {'os': os})
        with tempfile.TemporaryDirectory() as temp:
            with self.assertRaisesRegex(FileNotFoundError, 'template not found'):
                ns['_build']({'template': str(Path(temp) / 'missing.pptx')}, str(Path(temp) / 'out.pptx'))

    def test_invalid_template_value_is_rejected(self):
        ns = load_functions('build.py', {'_build'}, {'os': os})
        for value in ('', '   ', 123):
            with self.subTest(value=value), self.assertRaises(ValueError):
                ns['_build']({'template': value}, 'out.pptx')


@unittest.skipUnless(importlib.util.find_spec('lxml'), 'lxml is not installed')
class TextFrameXmlTests(unittest.TestCase):
    def test_multiline_replacement_removes_old_fields_and_keeps_font_style(self):
        from lxml import etree
        uri = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        def qn(name):
            return '{' + uri + '}' + name.split(':')[1]

        class Run:
            def __init__(self, element):
                self._r = element
            @property
            def text(self):
                return self._r.find(qn('a:t')).text
            @text.setter
            def text(self, value):
                element = self._r.find(qn('a:t'))
                if element is None:
                    element = etree.SubElement(self._r, qn('a:t'))
                element.text = value

        class Paragraph:
            def __init__(self, element):
                self._p = element
            @property
            def runs(self):
                return [Run(element) for element in self._p.findall(qn('a:r'))]
            def add_run(self):
                return Run(etree.SubElement(self._p, qn('a:r')))

        class Frame:
            def __init__(self, element):
                self.element = element
            @property
            def paragraphs(self):
                return [Paragraph(element) for element in self.element.findall(qn('a:p'))]
            def clear(self):
                first = self.paragraphs[0]._p
                for element in list(first):
                    if element.tag != qn('a:pPr'):
                        first.remove(element)
                for paragraph in self.paragraphs[1:]:
                    self.element.remove(paragraph._p)
            def add_paragraph(self):
                return Paragraph(etree.SubElement(self.element, qn('a:p')))

        xml = etree.fromstring(f'''<a:txBody xmlns:a="{uri}">
          <a:p><a:pPr algn="l"/><a:r><a:rPr sz="1800"/><a:t>Old run</a:t></a:r>
          <a:fld id="old"><a:t>Old field</a:t></a:fld></a:p>
          <a:p><a:r><a:t>Old paragraph</a:t></a:r></a:p></a:txBody>''')
        ns = load_functions('template_engine.py', {'_set_block_text'}, {'deepcopy': deepcopy, 'qn': qn})
        ns['_set_block_text'](SimpleNamespace(text_frame=Frame(xml)), 'One\nTwo\nThree')
        self.assertEqual([element.text for element in xml.iter(qn('a:t'))], ['One', 'Two', 'Three'])
        self.assertEqual([element.get('sz') for element in xml.iter(qn('a:rPr'))], ['1800'] * 3)
        self.assertFalse(list(xml.iter(qn('a:fld'))))


@unittest.skipUnless(importlib.util.find_spec('PIL'), 'Pillow is not installed')
class ImageCacheTests(unittest.TestCase):
    def test_replaced_image_invalidates_cache_even_with_same_mtime_and_size(self):
        from PIL import Image, ImageDraw
        with tempfile.TemporaryDirectory() as temp:
            source = Path(temp) / 'image.bmp'
            Image.new('RGB', (20, 20), 'red').save(source)
            original_stat = source.stat()
            ns = load_functions('build.py', {'_prepare_image'}, {
                'os': os, 'md5': md5, 'Image': Image, 'ImageDraw': ImageDraw, 'IMAGE_CACHE_DIR': temp,
            })
            first = ns['_prepare_image'](str(source), 1, 0)
            self.assertNotEqual(first, str(source))
            Image.new('RGB', (20, 20), 'blue').save(source)
            self.assertEqual(source.stat().st_size, original_stat.st_size)
            os.utime(source, ns=(original_stat.st_atime_ns, original_stat.st_mtime_ns))
            second = ns['_prepare_image'](str(source), 1, 0)
            self.assertNotEqual(first, second)
            with Image.open(second) as result:
                self.assertEqual(result.getpixel((10, 10)), (0, 0, 255))
            self.assertEqual(second, ns['_prepare_image'](str(source), 1, 0))


@unittest.skipUnless(importlib.util.find_spec('pptx'), 'install python/requirements.txt to run PPTX integration tests')
class PptxIntegrationTests(unittest.TestCase):
    def setUp(self):
        from pptx import Presentation
        from pptx.util import Inches, Pt
        self.prs = Presentation()
        self.slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        for text, size, top in [('Old title', 36, 0.5), ('Old body', 18, 2)]:
            shape = self.slide.shapes.add_textbox(Inches(1), Inches(top), Inches(7), Inches(2))
            shape.text = text
            shape.text_frame.paragraphs[0].runs[0].font.size = Pt(size)

    def test_generated_file_preserves_bullets_background_and_new_notes(self):
        from pptx import Presentation
        from pptx.dml.color import RGBColor
        from template_engine import build_from_template
        self.slide.background.fill.solid()
        self.slide.background.fill.fore_color.rgb = RGBColor.from_string('123456')
        self.slide.notes_slide.notes_text_frame.text = 'Old private notes'
        with tempfile.TemporaryDirectory() as temp:
            template, output = Path(temp) / 'template.pptx', Path(temp) / 'output.pptx'
            self.prs.save(template)
            build_from_template({'samplePages': {'content': 1}, 'slides': [
                {'type': 'content', 'title': 'New title', 'bullets': ['One', 'Two', 'Three'], 'notes': 'New notes'},
                {'type': 'content', 'title': 'Another title', 'bullets': ['Four']},
            ]}, str(output), str(template))
            result = Presentation(output)
            self.assertEqual(len(result.slides), 2)
            self.assertEqual(result.slides[0].shapes[1].text_frame.text, 'One\nTwo\nThree')
            self.assertEqual(str(result.slides[0].background.fill.fore_color.rgb), '123456')
            self.assertEqual(result.slides[0].notes_slide.notes_text_frame.text, 'New notes')
            self.assertFalse(result.slides[1].has_notes_slide)

    def test_unsupported_type_and_empty_template_fail_before_saving(self):
        from pptx import Presentation
        from template_engine import build_from_template
        with tempfile.TemporaryDirectory() as temp:
            template, output = Path(temp) / 'template.pptx', Path(temp) / 'output.pptx'
            self.prs.save(template)
            with self.assertRaisesRegex(ValueError, 'does not support'):
                build_from_template({'slides': [{'type': 'table', 'rows': [['Keep me']]}]}, str(output), str(template))
            self.assertFalse(output.exists())
            Presentation().save(template)
            with self.assertRaisesRegex(ValueError, 'at least one slide'):
                build_from_template({'slides': [{'type': 'content'}]}, str(output), str(template))


if __name__ == '__main__':
    unittest.main()
