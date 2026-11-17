import json
import os
from pathlib import Path
import sys
import tempfile
import subprocess
from types import SimpleNamespace
import unittest
from unittest.mock import patch
from zipfile import ZipFile

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'python'))
import bindings
import quality
from validation import validate_outline, ValidationError
from test_regressions import load_functions


class OutlineValidationTests(unittest.TestCase):
    def test_reports_multiple_paths_in_one_pass(self):
        with self.assertRaises(ValidationError) as caught:
            validate_outline({'slides': [{'type': 'content', 'bullets': [42], 'bulletz': []}, None], 'style': {'palette': {'ink': '#fff'}}})
        paths = {item['path'] for item in caught.exception.issues}
        self.assertTrue({'slides[0].bullets[0]', 'slides[0].bulletz', 'slides[1]', 'style.palette.ink'} <= paths)

    def test_rejects_free_layout_silent_truncation(self):
        for slide in ({'type': 'toc', 'items': ['A'] * 9}, {'type': 'stats', 'items': [{'value': 1}] * 7},
                      {'type': 'gallery', 'images': ['a'] * 5}):
            with self.subTest(slide=slide['type']), self.assertRaises(ValidationError):
                validate_outline({'slides': [slide]})

    def test_mapping_selectors_and_pages_are_validated(self):
        with self.assertRaises(ValidationError) as caught:
            validate_outline({'template': 'template.pptx', 'samplePages': {'content': 0},
                'fieldMap': {'content': {'title': {'id': True}, 'bullets': []}}, 'slides': [{'type': 'content'}]})
        self.assertGreaterEqual(len(caught.exception.issues), 3)

    def test_mapping_requires_template(self):
        with self.assertRaises(ValidationError):
            validate_outline({'fieldMap': {'content': {'title': 'Title'}}, 'slides': [{'title': 'Hello'}]})

    def test_valid_outline_keeps_numeric_zero_and_nested_bullets(self):
        outline = {'slides': [{'type': 'stats', 'items': [{'value': 0, 'label': 'Errors'}]},
                             {'bullets': [{'text': 'Nested', 'level': 1}]}]}
        self.assertIs(validate_outline(outline), outline)

    def test_invalid_numeric_style_is_rejected(self):
        for value in (True, '18', -1, float('nan'), float('inf')):
            with self.subTest(value=value), self.assertRaises(ValidationError):
                validate_outline({'style': {'bodySize': value}, 'slides': [{}]})


class FieldBindingTests(unittest.TestCase):
    def setUp(self):
        self.shapes = [SimpleNamespace(shape_id=2, name='Title', has_text_frame=True, has_table=False),
                       SimpleNamespace(shape_id=3, name='Body', has_text_frame=True, has_table=False)]
        self.slide = SimpleNamespace(shapes=self.shapes)
        self.adapter = patch.object(bindings, 'iter_shapes', lambda shapes: iter(shapes))
        self.adapter.start()
        self.addCleanup(self.adapter.stop)

    def test_names_and_ids_resolve_to_same_object(self):
        self.assertIs(bindings.select_shape(self.slide, 'Title'), bindings.select_shape(self.slide, {'id': 2}))

    def test_ambiguous_names_are_errors(self):
        self.shapes[1].name = 'Title'
        with self.assertRaisesRegex(ValueError, 'matched 2'):
            bindings.select_shape(self.slide, 'Title')

    def test_missing_content_binding_is_an_error(self):
        with self.assertRaisesRegex(ValueError, 'no fieldMap binding'):
            bindings.plan_fields(self.slide, {'type': 'content', 'title': 'Title', 'bullets': ['Keep']}, {'title': 'Title'})

    def test_reusing_one_shape_is_an_error(self):
        with self.assertRaisesRegex(ValueError, 'more than once'):
            bindings.plan_fields(self.slide, {'type': 'content'}, {'title': 'Title', 'bullets': {'id': 2}})

    def test_table_capacity_is_checked_before_rewrite(self):
        self.shapes[1].has_table = True
        self.shapes[1].table = SimpleNamespace(rows=[None], columns=[None])
        with self.assertRaisesRegex(ValueError, 'exceeds template capacity'):
            bindings.plan_fields(self.slide, {'type': 'table', 'headers': ['A'], 'rows': [['B']]}, {'table': 'Body'})

    def test_stats_keep_zero(self):
        self.assertEqual(bindings.field_value({'type': 'stats', 'items': [{'value': 0, 'label': 'Errors'}]}, 'items'), ['0\nErrors'])


class QualityEvidenceTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='ppt-forge-test-', dir=Path(__file__).resolve().parents[1])
        self.addCleanup(self.temp.cleanup)
        self.deck = Path(self.temp.name) / 'deck.pptx'
        self.deck.write_bytes(b'fixture deck bytes')
        self.image = Path(self.temp.name) / 'page-1.png'
        self.image.write_bytes(b'fixture preview bytes')
        self.report = {'formatVersion': 1, 'deck': str(self.deck), 'deckSha256': quality.file_hash(self.deck),
            'slideCount': 1, 'static': {'findings': []}, 'render': {'status': 'not-requested', 'missingFonts': []},
            'visualReview': {'status': 'pending', 'findings': []}}
        quality.write_json(quality.report_path(self.deck), quality.update_status(self.report))

    def render(self, **kwargs):
        return quality.record_render(self.deck, {'deckSha256': quality.file_hash(self.deck),
            'images': [{'page': 1, 'path': str(self.image)}], **kwargs})

    def test_static_or_render_success_alone_never_passes_delivery(self):
        self.assertFalse(self.report['readyForDelivery'])
        report = self.render(missingFonts=['Missing Font'])
        self.assertFalse(report['readyForDelivery'])
        self.assertEqual(report['render']['missingFonts'], ['Missing Font'])

    def test_review_pass_and_failure_are_persisted(self):
        rendered = self.render()
        reviewed = quality.record_review(self.deck, {'reviewToken': rendered['reviewToken'], 'findings': []})
        reviewed = quality.record_review(self.deck, {'reviewType': 'content', 'reviewToken': rendered['reviewToken'], 'findings': []})
        self.assertTrue(reviewed['readyForDelivery'])
        failed = quality.record_review(self.deck, {'reviewToken': rendered['reviewToken'],
            'findings': [{'page': 1, 'severity': 'error', 'message': 'Title clipped'}]})
        self.assertFalse(failed['readyForDelivery'])
        self.assertEqual(json.loads(Path(quality.report_path(self.deck)).read_text(encoding='utf-8'))['status'], 'blocked')

    def test_stale_token_and_preview_are_rejected(self):
        rendered = self.render()
        with self.assertRaisesRegex(ValueError, 'token'):
            quality.record_review(self.deck, {'reviewToken': 'old', 'findings': []})
        self.image.write_bytes(b'changed preview')
        with self.assertRaisesRegex(ValueError, 'Preview changed'):
            quality.record_review(self.deck, {'reviewToken': rendered['reviewToken'], 'findings': []})

    def test_deck_changed_during_render_is_rejected(self):
        with self.assertRaisesRegex(ValueError, 'changed during rendering'):
            self.render(deckSha256='old hash')

    def test_changed_deck_cannot_reuse_review(self):
        rendered = self.render()
        self.deck.write_bytes(b'changed deck bytes')
        with self.assertRaisesRegex(ValueError, 'stale'):
            quality.record_review(self.deck, {'reviewToken': rendered['reviewToken'], 'findings': []})

    def test_check_invalidates_changed_previews_after_a_pass(self):
        rendered = self.render()
        quality.record_review(self.deck, {'reviewToken': rendered['reviewToken'], 'findings': []})
        self.image.write_bytes(b'changed preview')
        result = subprocess.run([sys.executable, str(Path(quality.__file__)), 'check', str(self.deck)], capture_output=True, text=True, encoding='utf-8')
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertFalse(json.loads(result.stdout)['readyForDelivery'])

    def test_invalid_review_preserves_previous_report(self):
        rendered = self.render()
        before = Path(quality.report_path(self.deck)).read_bytes()
        with self.assertRaisesRegex(ValueError, 'valid page'):
            quality.record_review(self.deck, {'reviewToken': rendered['reviewToken'],
                'findings': [{'page': 2, 'severity': 'error', 'message': 'Wrong page'}]})
        self.assertEqual(Path(quality.report_path(self.deck)).read_bytes(), before)

    def test_partial_and_failed_render_cannot_be_reviewed(self):
        for evidence in ({'images': []}, {'images': [], 'error': 'Renderer unavailable'}):
            with self.subTest(evidence=evidence):
                rendered = self.render(**evidence)
                with self.assertRaisesRegex(ValueError, 'every page'):
                    quality.record_review(self.deck, {'reviewToken': rendered['reviewToken'], 'findings': []})

    def test_duplicate_pages_are_rejected(self):
        with self.assertRaisesRegex(ValueError, 'duplicate'):
            self.render(images=[{'page': 1, 'path': str(self.image)}] * 2)

    def test_internal_relationships_are_checked(self):
        deck = Path(self.temp.name) / 'broken.pptx'
        with ZipFile(deck, 'w') as archive:
            archive.writestr('[Content_Types].xml', '<Types/>')
            archive.writestr('ppt/presentation.xml', '<presentation/>')
            archive.writestr('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="ppt/missing.xml"/></Relationships>')
        self.assertIn('BROKEN_RELATIONSHIP', [item['code'] for item in quality.integrity_findings(deck)])


class AtomicBuildTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='ppt-forge-test-', dir=Path(__file__).resolve().parents[1])
        self.addCleanup(self.temp.cleanup)
        self.output = Path(self.temp.name) / 'deck.pptx'
        self.output.write_bytes(b'original deck')

    def namespace(self, findings=None, fail=False):
        def renderer(outline, path):
            Path(path).write_bytes(b'complete generated deck')
            if fail:
                raise ValueError('Generation failed')
            return {'slideCount': 1, 'theme': 'navy'}
        def analyzer(path):
            return quality.update_status({'deck': path, 'deckSha256': quality.file_hash(path), 'slideCount': 1,
                'static': {'findings': findings or []}, 'render': {'status': 'not-requested'},
                'visualReview': {'status': 'pending', 'findings': []}})
        return load_functions('build.py', {'build'}, {'os': os, 'tempfile': tempfile, 'validate_outline': validate_outline,
            'ValidationError': ValidationError, '_build': renderer, 'analyze_deck': analyzer,
            'report_path': quality.report_path, 'write_json': quality.write_json})

    def test_generation_failure_and_strict_warnings_preserve_original(self):
        for ns in (self.namespace(fail=True), self.namespace([{'severity': 'warning', 'page': 1, 'code': 'TEXT_CAPACITY', 'message': 'Too much text'}])):
            with self.assertRaises(ValueError):
                ns['build']({'qualityMode': 'strict', 'slides': [{}]}, str(self.output))
            self.assertEqual(self.output.read_bytes(), b'original deck')
            self.assertFalse(list(self.output.parent.glob('.ppt-forge-*.pptx')))

    def test_success_publishes_complete_file_and_matching_report(self):
        result = self.namespace()['build']({'slides': [{}]}, str(self.output))
        self.assertEqual(self.output.read_bytes(), b'complete generated deck')
        report = json.loads(Path(result['qualityReport']).read_text(encoding='utf-8'))
        self.assertEqual(report['deckSha256'], quality.file_hash(self.output))
        self.assertFalse(report['readyForDelivery'])


if __name__ == '__main__':
    unittest.main()
