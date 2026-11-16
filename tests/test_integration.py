"""Real PPTX integration tests. npm run test:integration requires dependencies."""
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'python'))


@unittest.skipUnless(importlib.util.find_spec('pptx'), 'python-pptx is required')
class RealDeckTests(unittest.TestCase):
    def setUp(self):
        from pptx import Presentation
        from pptx.util import Inches, Pt
        self.temp = tempfile.TemporaryDirectory(prefix='ppt-forge-test-', dir=ROOT)
        self.addCleanup(self.temp.cleanup)
        self.template = Path(self.temp.name) / 'template.pptx'
        self.output = Path(self.temp.name) / 'output.pptx'
        self.prs = Presentation()
        self.slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        for name, text, size, top in [('Headline', 'Old title', 36, 0.4), ('Body', 'Old body', 18, 1.5), ('Brand', 'Brand preserved', 9, 7)]:
            shape = self.slide.shapes.add_textbox(Inches(0.5), Inches(top), Inches(8), Inches(1))
            shape.name = name
            shape.text = text
            shape.text_frame.paragraphs[0].runs[0].font.size = Pt(size)
        self.prs.save(self.template)

    def test_explicit_mapping_inventory_and_report(self):
        from pptx import Presentation
        from build import build
        from inspection import inspect_deck
        inventory = inspect_deck(str(self.template))
        names = {item['name']: item['id'] for item in inventory['slides'][0]['objects']}
        self.assertIn('capacityCjkApprox', inventory['slides'][0]['objects'][0])
        outline = {'template': str(self.template), 'samplePages': {'content': 1},
            'fieldMap': {'content': {'title': {'id': names['Headline']}, 'bullets': 'Body'}},
            'slides': [{'type': 'content', 'title': 'New title', 'bullets': ['One', 'Two', 'Three']}]}
        result = build(outline, str(self.output))
        deck = Presentation(self.output)
        texts = {shape.name: shape.text for shape in deck.slides[0].shapes}
        self.assertEqual(texts['Headline'], 'New title')
        self.assertEqual(texts['Body'], 'One\nTwo\nThree')
        self.assertEqual(texts['Brand'], 'Brand preserved')
        report = json.loads(Path(result['qualityReport']).read_text(encoding='utf-8'))
        self.assertFalse(report['readyForDelivery'])
        self.assertEqual(report['deck'], str(self.output))
        self.assertEqual(report['visualReview']['status'], 'pending')

    def test_bad_binding_and_schema_leave_existing_file_unchanged(self):
        from build import build
        self.output.write_bytes(b'existing user deck')
        for outline in (
            {'slides': [{'type': 'content', 'bullets': [17]}]},
            {'template': str(self.template), 'fieldMap': {'content': {'title': 'Missing'}},
             'slides': [{'type': 'content', 'title': 'New'}]},
            {'template': str(self.template), 'samplePages': {'content': 99}, 'slides': [{'type': 'content', 'title': 'New'}]},
        ):
            with self.subTest(outline=outline), self.assertRaises(ValueError):
                build(outline, str(self.output))
            self.assertEqual(self.output.read_bytes(), b'existing user deck')

    def test_named_table_binding_writes_every_cell(self):
        from pptx import Presentation
        from pptx.util import Inches
        from build import build
        table = self.slide.shapes.add_table(3, 2, Inches(1), Inches(3), Inches(6), Inches(2))
        table.name = 'Data'
        self.prs.save(self.template)
        build({'template': str(self.template), 'samplePages': {'table': 1},
            'fieldMap': {'table': {'title': 'Headline', 'table': 'Data'}},
            'slides': [{'type': 'table', 'title': 'Metrics', 'headers': ['Name', 'Value'], 'rows': [['Errors', 0], ['Sales', 12]]}]}, str(self.output))
        result = next(shape.table for shape in Presentation(self.output).slides[0].shapes if shape.has_table)
        self.assertEqual(result.cell(1, 1).text, '0')
        self.assertEqual(result.cell(2, 1).text, '12')

    def test_image_binding_keeps_shape_identity_and_replaces_media(self):
        from PIL import Image
        from pptx import Presentation
        from pptx.util import Inches
        from build import build
        first, second = Path(self.temp.name) / 'first.png', Path(self.temp.name) / 'second.png'
        Image.new('RGB', (30, 20), 'red').save(first)
        Image.new('RGB', (30, 20), 'blue').save(second)
        image = self.slide.shapes.add_picture(str(first), Inches(1), Inches(3), Inches(3), Inches(2))
        image.name = 'Photo'
        identity = image.shape_id
        self.prs.save(self.template)
        build({'template': str(self.template), 'fieldMap': {'image': {'title': 'Headline', 'image': 'Photo'}},
            'slides': [{'type': 'image', 'title': 'New', 'image': str(second)}]}, str(self.output))
        result = next(shape for shape in Presentation(self.output).slides[0].shapes if shape.name == 'Photo')
        self.assertEqual(result.shape_id, identity)
        self.assertEqual(result.image.blob, second.read_bytes())

    def test_free_layout_types_and_zero_value(self):
        from pptx import Presentation
        from build import build
        slides = [{'type': 'cover', 'title': 'Demo'}, {'type': 'toc', 'items': ['One']},
            {'type': 'section', 'number': '01', 'title': 'Section'}, {'type': 'content', 'bullets': ['Bullet']},
            {'type': 'two-column', 'leftBullets': ['L'], 'rightBullets': ['R']},
            {'type': 'table', 'headers': ['H'], 'rows': [['V']]}, {'type': 'image'},
            {'type': 'image-text', 'bullets': ['B']}, {'type': 'gallery', 'captions': ['Placeholder']},
            {'type': 'stats', 'items': [{'value': 0, 'label': 'Errors'}]}, {'type': 'end'}]
        result = build({'slides': slides}, str(self.output))
        deck = Presentation(self.output)
        self.assertEqual(len(deck.slides), 11)
        self.assertIn('0', [shape.text for shape in deck.slides[9].shapes if shape.has_text_frame])
        self.assertFalse(result['quality']['readyForDelivery'])

    def test_cli_invalid_data_has_structured_errors(self):
        result = subprocess.run([sys.executable, str(ROOT / 'python' / 'build.py'), '-', str(self.output)],
            input=json.dumps({'slides': [{'type': 'content', 'bullets': [3]}]}), text=True, capture_output=True, encoding='utf-8')
        self.assertEqual(result.returncode, 1)
        self.assertEqual(json.loads(result.stdout)['code'], 'INVALID_OUTLINE')
        self.assertFalse(self.output.exists())


if __name__ == '__main__':
    unittest.main()
