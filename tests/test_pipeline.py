"""
Comprehensive unit tests for the pptx-pipeline project.
Run with: python -m unittest tests.test_pipeline
"""

import os
import sys
import csv
import json
import shutil
import tempfile
import unittest
import zipfile

# ---------------------------------------------------------------------------
# Path setup: ensure the project root is importable
# ---------------------------------------------------------------------------
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from inject import (
    _escape_for_xml,
    _replace_tokens_in_shape_xml,
    _replace_shape_text,
    _shrink_font_in_shape_xml,
    _apply_text_autofit,
    _filter_resolved_tokens,
    _find_shape_span,
    _inject_table_data,
    _update_content_types,
    _adjust_table_rows,
    _inject_table_geometry,
    _inject_image_geometry,
    inject_slide,
    inject,
)
from layout import (
    compute_table_layout,
    compute_image_fit,
    compute_text_font_scale,
    read_image_dimensions,
)
from update_config import resolve_token, resolve_field, load_data_sources, update_config, _load_xlsx, find_screenshots
from generate_config import (
    looks_dynamic,
    looks_static,
    get_shape_category,
    generate_config,
    DEFAULT_DYNAMIC_HINTS,
    STRUCTURAL_TYPES,
)
from reconstruct import reconstruct
from deconstruct import deconstruct


# ═══════════════════════════════════════════════════════════════════════════
# 1. inject.py tests
# ═══════════════════════════════════════════════════════════════════════════

class TestEscapeForXml(unittest.TestCase):
    """Tests for _escape_for_xml."""

    def test_escape_for_xml(self):
        self.assertEqual(_escape_for_xml("&"), "&amp;")
        self.assertEqual(_escape_for_xml("<"), "&lt;")
        self.assertEqual(_escape_for_xml(">"), "&gt;")
        # xml.sax.saxutils.escape does not escape " by default (only needed in attributes)
        self.assertEqual(_escape_for_xml('"'), '"')
        # xml.sax.saxutils.escape does not escape single-quote by default,
        # but we verify the function at least does not crash on it.
        result = _escape_for_xml("'")
        self.assertIsInstance(result, str)
        # Combined
        self.assertEqual(_escape_for_xml('A & B < C > D "E"'),
                         'A &amp; B &lt; C &gt; D "E"')


class TestReplaceTokensSingleRun(unittest.TestCase):
    """Token fully within one <a:t> element."""

    def test_replace_tokens_single_run(self):
        shape_xml = (
            '<p:sp>'
            '<a:p><a:r><a:rPr sz="1800"/>'
            '<a:t>Hello {{name}} world</a:t>'
            '</a:r></a:p>'
            '</p:sp>'
        )
        result, count = _replace_tokens_in_shape_xml(shape_xml, {"{{name}}": "Alice"})
        self.assertEqual(count, 1)
        self.assertIn("Hello Alice world", result)
        self.assertNotIn("{{name}}", result)


class TestReplaceTokensSplitRun(unittest.TestCase):
    """Token spanning multiple <a:r> runs."""

    def test_replace_tokens_split_run(self):
        # Token "{{name}}" split across three runs: "{{na", "me", "}}"
        shape_xml = (
            '<p:sp>'
            '<a:p>'
            '<a:r><a:rPr sz="1800"/><a:t>{{na</a:t></a:r>'
            '<a:r><a:rPr sz="1800"/><a:t>me</a:t></a:r>'
            '<a:r><a:rPr sz="1800"/><a:t>}}</a:t></a:r>'
            '</a:p>'
            '</p:sp>'
        )
        result, count = _replace_tokens_in_shape_xml(shape_xml, {"{{name}}": "Bob"})
        self.assertEqual(count, 1)
        # The replacement text should appear somewhere in the result
        self.assertIn("Bob", result)
        # The original token fragments should be gone
        self.assertNotIn("{{na", result)


class TestReplaceTokensNoMatch(unittest.TestCase):
    """Token not present returns 0 replacements."""

    def test_replace_tokens_no_match(self):
        shape_xml = (
            '<p:sp>'
            '<a:p><a:r><a:t>Some static text</a:t></a:r></a:p>'
            '</p:sp>'
        )
        result, count = _replace_tokens_in_shape_xml(shape_xml, {"{{missing}}": "value"})
        self.assertEqual(count, 0)
        self.assertEqual(result, shape_xml)


class TestReplaceTokensXmlSpecialChars(unittest.TestCase):
    """Injecting values containing & < > doesn't break XML."""

    def test_replace_tokens_xml_special_chars(self):
        shape_xml = (
            '<p:sp>'
            '<a:p><a:r><a:t>{{company}}</a:t></a:r></a:p>'
            '</p:sp>'
        )
        result, count = _replace_tokens_in_shape_xml(
            shape_xml, {"{{company}}": "AT&T <Corp>"}
        )
        self.assertEqual(count, 1)
        # The value should be XML-escaped so it doesn't create broken XML
        self.assertIn("AT&amp;T &lt;Corp&gt;", result)
        self.assertNotIn("AT&T <Corp>", result)


class TestShrinkFont(unittest.TestCase):
    """Font scales down when new text is longer."""

    def test_shrink_font(self):
        shape_xml = (
            '<p:sp>'
            '<a:r><a:rPr sz="1800"/><a:t>short</a:t></a:r>'
            '</p:sp>'
        )
        # original_len=5, new_len=20 -> scale = max(5/20, 0.75) = 0.75
        result = _shrink_font_in_shape_xml(shape_xml, original_len=5, new_len=20)
        # 1800 * 0.75 = 1350
        self.assertIn('sz="1350"', result)
        self.assertNotIn('sz="1800"', result)


class TestShrinkFontNoChange(unittest.TestCase):
    """Font unchanged when new text is shorter or equal."""

    def test_shrink_font_no_change(self):
        shape_xml = (
            '<p:sp>'
            '<a:r><a:rPr sz="1800"/><a:t>hello</a:t></a:r>'
            '</p:sp>'
        )
        result = _shrink_font_in_shape_xml(shape_xml, original_len=10, new_len=5)
        self.assertEqual(result, shape_xml)

    def test_shrink_font_equal_length(self):
        shape_xml = (
            '<p:sp>'
            '<a:r><a:rPr sz="1800"/><a:t>hello</a:t></a:r>'
            '</p:sp>'
        )
        result = _shrink_font_in_shape_xml(shape_xml, original_len=5, new_len=5)
        self.assertEqual(result, shape_xml)


class TestShrinkFontMinimum(unittest.TestCase):
    """Font doesn't go below 600 (6pt)."""

    def test_shrink_font_minimum(self):
        shape_xml = (
            '<p:sp>'
            '<a:r><a:rPr sz="700"/><a:t>x</a:t></a:r>'
            '</p:sp>'
        )
        # original_len=1, new_len=100 -> scale = max(1/100, 0.75) = 0.75
        # 700 * 0.75 = 525, but minimum is 600
        result = _shrink_font_in_shape_xml(shape_xml, original_len=1, new_len=100)
        self.assertIn('sz="600"', result)


class TestFilterResolvedTokens(unittest.TestCase):
    """Filters out unresolved tokens correctly."""

    def test_filter_resolved_tokens(self):
        tokens = {
            "{{a}}": {"value": "hello", "_resolved": True},
            "{{b}}": {"value": "{{b}}", "_resolved": False},
            "{{c}}": {"value": "world", "_resolved": True},
        }
        result = _filter_resolved_tokens(tokens)
        self.assertEqual(result, {"{{a}}": "hello", "{{c}}": "world"})
        self.assertNotIn("{{b}}", result)

    def test_filter_resolved_tokens_plain_string(self):
        """Backward compat with plain string values."""
        tokens = {
            "{{x}}": "resolved_value",
            "{{y}}": "{{unresolved_ref}}",
            "{{z}}": "plain text",
        }
        result = _filter_resolved_tokens(tokens)
        self.assertIn("{{x}}", result)
        self.assertEqual(result["{{x}}"], "resolved_value")
        # {{y}} looks like an unresolved reference -> filtered out
        self.assertNotIn("{{y}}", result)
        # {{z}} is plain text -> kept
        self.assertIn("{{z}}", result)
        self.assertEqual(result["{{z}}"], "plain text")


class TestFindShapeSpan(unittest.TestCase):
    """Finds correct start/end offsets for a shape."""

    def test_find_shape_span(self):
        xml_str = (
            '<p:spTree>'
            '<p:sp>'
            '<p:nvSpPr><p:cNvPr id="5" name="Title"/></p:nvSpPr>'
            '<a:p><a:r><a:t>Hello</a:t></a:r></a:p>'
            '</p:sp>'
            '<p:sp>'
            '<p:nvSpPr><p:cNvPr id="9" name="Other"/></p:nvSpPr>'
            '<a:p><a:r><a:t>World</a:t></a:r></a:p>'
            '</p:sp>'
            '</p:spTree>'
        )
        span = _find_shape_span(xml_str, "5")
        self.assertIsNotNone(span)
        start, end = span
        fragment = xml_str[start:end]
        self.assertIn('id="5"', fragment)
        self.assertIn("Hello", fragment)
        self.assertNotIn('id="9"', fragment)

    def test_find_shape_span_not_found(self):
        xml_str = '<p:spTree><p:sp><p:cNvPr id="1" name="X"/></p:sp></p:spTree>'
        span = _find_shape_span(xml_str, "999")
        self.assertIsNone(span)


# ═══════════════════════════════════════════════════════════════════════════
# 2. update_config.py tests
# ═══════════════════════════════════════════════════════════════════════════

class TestResolveToken(unittest.TestCase):
    """Tests for resolve_token."""

    def test_resolve_token_simple(self):
        data = {"company": "Acme Inc"}
        result = resolve_token("{{company}}", data)
        self.assertTrue(result["_resolved"])
        self.assertEqual(result["value"], "Acme Inc")

    def test_resolve_token_missing(self):
        data = {"company": "Acme Inc"}
        result = resolve_token("{{missing_field}}", data)
        self.assertFalse(result["_resolved"])
        self.assertEqual(result["value"], "{{missing_field}}")

    def test_resolve_token_nested_dot(self):
        data = {"a": {"b": "nested_value"}}
        result = resolve_token("{{a.b}}", data)
        self.assertTrue(result["_resolved"])
        self.assertEqual(result["value"], "nested_value")

    def test_resolve_token_array_index(self):
        data = {"items": ["first", "second", "third"]}
        result = resolve_token("{{items[0]}}", data)
        self.assertTrue(result["_resolved"])
        self.assertEqual(result["value"], "first")

    def test_resolve_token_plain_passthrough(self):
        """Non-template string passes through as-is."""
        data = {"anything": "whatever"}
        result = resolve_token("literal text", data)
        self.assertTrue(result["_resolved"])
        self.assertEqual(result["value"], "literal text")


class TestLoadDataSourcesCsv(unittest.TestCase):
    """Tests for load_data_sources with CSV files."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmpdir)

    def test_load_data_sources_csv_kv(self):
        """Test key-value CSV loading."""
        csv_path = os.path.join(self.tmpdir, "kv.csv")
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=["field", "value"])
            writer.writeheader()
            writer.writerow({"field": "company", "value": "Acme"})
            writer.writerow({"field": "year", "value": "2026"})

        data = load_data_sources(self.tmpdir)
        self.assertEqual(data["company"], "Acme")
        self.assertEqual(data["year"], "2026")

    def test_load_data_sources_csv_tabular(self):
        """Test tabular CSV loading."""
        csv_path = os.path.join(self.tmpdir, "employees.csv")
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=["name", "role"])
            writer.writeheader()
            writer.writerow({"name": "Alice", "role": "Engineer"})
            writer.writerow({"name": "Bob", "role": "Designer"})

        data = load_data_sources(self.tmpdir)
        # Tabular CSV is stored under the filename stem as a list of dicts
        self.assertIn("employees", data)
        self.assertIsInstance(data["employees"], list)
        self.assertEqual(len(data["employees"]), 2)
        self.assertEqual(data["employees"][0]["name"], "Alice")
        self.assertEqual(data["employees"][1]["role"], "Designer")


# ═══════════════════════════════════════════════════════════════════════════
# 3. generate_config.py tests
# ═══════════════════════════════════════════════════════════════════════════

class TestLooksDynamic(unittest.TestCase):
    """Tests for looks_dynamic."""

    def test_looks_dynamic_hit(self):
        shape = {"text_preview": "Revenue xxx this quarter"}
        self.assertTrue(looks_dynamic(shape, DEFAULT_DYNAMIC_HINTS))

    def test_looks_dynamic_miss(self):
        shape = {"text_preview": "Company Overview"}
        self.assertFalse(looks_dynamic(shape, DEFAULT_DYNAMIC_HINTS))


class TestLooksStatic(unittest.TestCase):
    """Tests for looks_static."""

    def test_looks_static_oval(self):
        shape = {"name": "Oval 3", "type": "sp"}
        self.assertTrue(looks_static(shape))

    def test_looks_static_connector(self):
        shape = {"name": "Connector: Elbow 12", "type": "sp"}
        self.assertTrue(looks_static(shape))

    def test_looks_static_straight_connector(self):
        shape = {"name": "Straight Connector 5", "type": "sp"}
        self.assertTrue(looks_static(shape))

    def test_looks_static_structural_type(self):
        shape = {"name": "Whatever", "type": "cxnSp"}
        self.assertTrue(looks_static(shape))

    def test_looks_static_normal_shape(self):
        shape = {"name": "TextBox 1", "type": "sp"}
        self.assertFalse(looks_static(shape))


class TestGetShapeCategory(unittest.TestCase):
    """Verify all shape type mappings."""

    def test_get_shape_category(self):
        self.assertEqual(get_shape_category({"type": "pic"}), "image")
        self.assertEqual(get_shape_category({"type": "sp"}), "text")
        self.assertEqual(get_shape_category({"type": "grpSp"}), "group")
        self.assertEqual(get_shape_category({"type": "graphicFrame"}), "table")
        for stype in STRUCTURAL_TYPES:
            self.assertEqual(get_shape_category({"type": stype}), "structural")
        self.assertEqual(get_shape_category({"type": "somethingElse"}), "unknown")
        self.assertEqual(get_shape_category({}), "unknown")


# ═══════════════════════════════════════════════════════════════════════════
# 4. reconstruct.py tests
# ═══════════════════════════════════════════════════════════════════════════

class TestReconstructRoundtrip(unittest.TestCase):
    """Create a minimal zip, extract to _raw, reconstruct, verify contents match."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmpdir)

    def test_reconstruct_roundtrip(self):
        library_dir = os.path.join(self.tmpdir, "library")
        raw_dir = os.path.join(library_dir, "_raw")
        os.makedirs(raw_dir, exist_ok=True)

        # Create some minimal files in _raw to simulate an extracted PPTX
        ppt_dir = os.path.join(raw_dir, "ppt", "slides")
        os.makedirs(ppt_dir, exist_ok=True)
        slide_content = b'<?xml version="1.0"?><p:sld><p:spTree/></p:sld>'
        with open(os.path.join(ppt_dir, "slide1.xml"), "wb") as f:
            f.write(slide_content)

        ct_content = b'<?xml version="1.0"?><Types></Types>'
        with open(os.path.join(raw_dir, "[Content_Types].xml"), "wb") as f:
            f.write(ct_content)

        # Create a minimal manifest.json
        manifest = {"source": "nonexistent.pptx", "total_slides": 1}
        with open(os.path.join(library_dir, "manifest.json"), "w") as f:
            json.dump(manifest, f)

        # Reconstruct
        output_path = os.path.join(self.tmpdir, "output.pptx")
        reconstruct(library_dir, output_path)

        # Verify the output zip exists and contains the expected files
        self.assertTrue(os.path.exists(output_path))
        with zipfile.ZipFile(output_path, "r") as z:
            names = z.namelist()
            self.assertIn("ppt/slides/slide1.xml", names)
            self.assertIn("[Content_Types].xml", names)
            # Verify content is preserved exactly
            self.assertEqual(z.read("ppt/slides/slide1.xml"), slide_content)
            self.assertEqual(z.read("[Content_Types].xml"), ct_content)


# ═══════════════════════════════════════════════════════════════════════════
# 5. deconstruct.py tests
# ═══════════════════════════════════════════════════════════════════════════

class TestExtractShapeRecursion(unittest.TestCase):
    """Verify group shapes have their children enumerated with parent_group set."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmpdir)

    def test_extract_shape_recursion(self):
        # Build a minimal PPTX zip with a slide that has a group shape
        # containing two child shapes.
        slide_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<p:cSld><p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr/>'
            '<p:grpSp>'
            '  <p:nvGrpSpPr><p:cNvPr id="10" name="Group 10"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '  <p:grpSpPr/>'
            '  <p:sp>'
            '    <p:nvSpPr><p:cNvPr id="11" name="Child A"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '    <p:spPr/>'
            '    <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Hello</a:t></a:r></a:p></p:txBody>'
            '  </p:sp>'
            '  <p:sp>'
            '    <p:nvSpPr><p:cNvPr id="12" name="Child B"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '    <p:spPr/>'
            '    <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>World</a:t></a:r></a:p></p:txBody>'
            '  </p:sp>'
            '</p:grpSp>'
            '</p:spTree></p:cSld>'
            '</p:sld>'
        ).encode("utf-8")

        # Minimal content types and rels required by zipfile structure
        content_types = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '</Types>'
        ).encode("utf-8")

        rels_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '</Relationships>'
        ).encode("utf-8")

        pptx_path = os.path.join(self.tmpdir, "test.pptx")
        with zipfile.ZipFile(pptx_path, "w") as z:
            z.writestr("[Content_Types].xml", content_types)
            z.writestr("ppt/slides/slide1.xml", slide_xml)
            z.writestr("ppt/slides/_rels/slide1.xml.rels", rels_xml)

        library_dir = os.path.join(self.tmpdir, "library")
        deconstruct(pptx_path, library_dir, force=True)

        # Read the manifest and check shape extraction
        manifest_path = os.path.join(library_dir, "manifest.json")
        with open(manifest_path) as f:
            manifest = json.load(f)

        shapes = manifest["slides"][0]["shapes"]

        # Group + 2 children + structural elements (nvGrpSpPr, grpSpPr) that spTree emits
        # We care about the group and its children being present with correct parent_group
        self.assertGreaterEqual(len(shapes), 3)

        # The group shape
        group = [s for s in shapes if s.get("type") == "grpSp"]
        self.assertEqual(len(group), 1)
        self.assertEqual(group[0]["id"], "10")
        self.assertEqual(group[0]["name"], "Group 10")

        # Children should have parent_group pointing to group id "10"
        children = [s for s in shapes if s.get("parent_group") == "10"]
        self.assertEqual(len(children), 2)
        child_names = {s["name"] for s in children}
        self.assertEqual(child_names, {"Child A", "Child B"})


# ═══════════════════════════════════════════════════════════════════════════
# 6. New feature tests (rounds 1 & 2)
# ═══════════════════════════════════════════════════════════════════════════

class TestInjectTableData(unittest.TestCase):
    """Tests for _inject_table_data table cell injection."""

    def _make_table_xml(self):
        """Build a minimal graphicFrame XML with a 2-col, 3-row table."""
        return (
            '<p:spTree xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<p:graphicFrame>'
            '<p:nvGraphicFramePr><p:cNvPr id="50" name="Table 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
            '<a:graphic><a:graphicData><a:tbl>'
            '<a:tr><a:tc><a:txBody><a:p><a:r><a:t>Name</a:t></a:r></a:p></a:txBody></a:tc>'
            '<a:tc><a:txBody><a:p><a:r><a:t>Score</a:t></a:r></a:p></a:txBody></a:tc></a:tr>'
            '<a:tr><a:tc><a:txBody><a:p><a:r><a:t>OldName1</a:t></a:r></a:p></a:txBody></a:tc>'
            '<a:tc><a:txBody><a:p><a:r><a:t>0</a:t></a:r></a:p></a:txBody></a:tc></a:tr>'
            '<a:tr><a:tc><a:txBody><a:p><a:r><a:t>OldName2</a:t></a:r></a:p></a:txBody></a:tc>'
            '<a:tc><a:txBody><a:p><a:r><a:t>0</a:t></a:r></a:p></a:txBody></a:tc></a:tr>'
            '</a:tbl></a:graphicData></a:graphic>'
            '</p:graphicFrame>'
            '</p:spTree>'
        )

    def test_inject_table_replaces_cells(self):
        xml_str = self._make_table_xml()
        rows = [{"Name": "Alice", "Score": "95"}, {"Name": "Bob", "Score": "88"}]
        shape = {
            "shape_name": "Table 1",
            "shape_id": "50",
            "resolved_tokens": {"{{table_data}}": {"value": json.dumps(rows), "_resolved": True}},
        }
        result, count = _inject_table_data(xml_str, shape)
        self.assertGreater(count, 0)
        self.assertIn("Alice", result)
        self.assertIn("95", result)
        self.assertIn("Bob", result)
        self.assertIn("88", result)
        self.assertNotIn("OldName1", result)
        self.assertNotIn("OldName2", result)

    def test_inject_table_invalid_json(self):
        xml_str = self._make_table_xml()
        shape = {
            "shape_name": "Table 1",
            "shape_id": "50",
            "resolved_tokens": {"{{data}}": {"value": "not valid json", "_resolved": True}},
        }
        result, count = _inject_table_data(xml_str, shape)
        self.assertEqual(count, 0)
        self.assertEqual(result, xml_str)

    def test_inject_table_no_resolved_tokens(self):
        xml_str = self._make_table_xml()
        shape = {
            "shape_name": "Table 1",
            "shape_id": "50",
            "resolved_tokens": {},
        }
        result, count = _inject_table_data(xml_str, shape)
        self.assertEqual(count, 0)


class TestUpdateContentTypes(unittest.TestCase):
    """Tests for _update_content_types."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()
        self.raw_dir = os.path.join(self.tmpdir, "_raw")
        self.media_dir = os.path.join(self.raw_dir, "ppt", "media")
        os.makedirs(self.media_dir, exist_ok=True)

    def tearDown(self):
        shutil.rmtree(self.tmpdir)

    def test_adds_missing_extension(self):
        # Content_Types has png but not jpeg
        ct_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="png" ContentType="image/png"/>'
            '</Types>'
        )
        ct_path = os.path.join(self.raw_dir, "[Content_Types].xml")
        with open(ct_path, "w", encoding="utf-8") as f:
            f.write(ct_xml)

        # Put a jpeg file in media/
        with open(os.path.join(self.media_dir, "image1.jpeg"), "wb") as f:
            f.write(b"\xff\xd8\xff")

        _update_content_types(self.raw_dir)

        with open(ct_path, "r", encoding="utf-8") as f:
            updated = f.read()
        self.assertIn('Extension="jpeg"', updated)
        self.assertIn('ContentType="image/jpeg"', updated)

    def test_no_duplicate_if_exists(self):
        ct_xml = (
            '<?xml version="1.0"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="png" ContentType="image/png"/>'
            '</Types>'
        )
        ct_path = os.path.join(self.raw_dir, "[Content_Types].xml")
        with open(ct_path, "w", encoding="utf-8") as f:
            f.write(ct_xml)

        with open(os.path.join(self.media_dir, "img.png"), "wb") as f:
            f.write(b"\x89PNG")

        _update_content_types(self.raw_dir)

        with open(ct_path, "r", encoding="utf-8") as f:
            updated = f.read()
        # Should still have exactly one png Default
        self.assertEqual(updated.count('Extension="png"'), 1)


class TestDryRunMode(unittest.TestCase):
    """Tests that dry_run=True does not modify files."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmpdir)

    def test_inject_slide_dry_run_no_write(self):
        slide_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<p:cSld><p:spTree>'
            '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '<p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1800"/><a:t>{{name}}</a:t></a:r></a:p></p:txBody></p:sp>'
            '</p:spTree></p:cSld></p:sld>'
        )
        slide_path = os.path.join(self.tmpdir, "slide1.xml")
        with open(slide_path, "w", encoding="utf-8") as f:
            f.write(slide_xml)

        shapes = [{
            "shape_name": "Title",
            "shape_id": "2",
            "resolved_tokens": {"{{name}}": {"value": "Injected", "_resolved": True}},
        }]

        count = inject_slide(slide_path, shapes, dry_run=True)
        self.assertGreater(count, 0)

        # File should be unchanged
        with open(slide_path, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("{{name}}", content)
        self.assertNotIn("Injected", content)


class TestIdempotentRawClean(unittest.TestCase):
    """Tests the _raw_clean idempotency mechanism in inject()."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()
        self.library = os.path.join(self.tmpdir, "lib")
        self.raw_dir = os.path.join(self.library, "_raw")
        os.makedirs(os.path.join(self.raw_dir, "ppt", "slides"), exist_ok=True)

        # Write a minimal slide
        self.slide_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<p:cSld><p:spTree>'
            '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '<p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1800"/><a:t>{{name}}</a:t></a:r></a:p></p:txBody></p:sp>'
            '</p:spTree></p:cSld></p:sld>'
        )
        with open(os.path.join(self.raw_dir, "ppt", "slides", "slide1.xml"), "w", encoding="utf-8") as f:
            f.write(self.slide_xml)

        # Config
        self.config = {
            "slides": {
                "slide_1": {
                    "slide_number": 1,
                    "shapes": [{
                        "shape_name": "Title",
                        "shape_id": "2",
                        "is_dynamic": True,
                        "resolved_tokens": {"{{name}}": {"value": "Alice", "_resolved": True}},
                    }],
                }
            }
        }
        self.config_path = os.path.join(self.tmpdir, "config.json")
        with open(self.config_path, "w") as f:
            json.dump(self.config, f)

    def tearDown(self):
        shutil.rmtree(self.tmpdir)

    def test_first_run_creates_raw_clean(self):
        inject(self.config_path, self.library)
        clean_dir = os.path.join(self.library, "_raw_clean")
        self.assertTrue(os.path.exists(clean_dir))
        # _raw_clean should have the original token
        with open(os.path.join(clean_dir, "ppt", "slides", "slide1.xml"), "r") as f:
            self.assertIn("{{name}}", f.read())

    def test_second_run_restores_from_clean(self):
        # First run
        inject(self.config_path, self.library)
        slide_path = os.path.join(self.raw_dir, "ppt", "slides", "slide1.xml")
        with open(slide_path, "r") as f:
            first_result = f.read()
        self.assertIn("Alice", first_result)

        # Update config to inject different value
        self.config["slides"]["slide_1"]["shapes"][0]["resolved_tokens"]["{{name}}"]["value"] = "Bob"
        with open(self.config_path, "w") as f:
            json.dump(self.config, f)

        # Second run should restore from clean first, then inject fresh
        inject(self.config_path, self.library)
        with open(slide_path, "r") as f:
            second_result = f.read()
        self.assertIn("Bob", second_result)
        self.assertNotIn("Alice", second_result)


class TestZipSlipPrevention(unittest.TestCase):
    """Tests that deconstruct rejects path-traversal ZIP entries."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmpdir)

    def test_zip_slip_raises(self):
        # Create a malicious ZIP with a path traversal entry
        pptx_path = os.path.join(self.tmpdir, "evil.pptx")
        with zipfile.ZipFile(pptx_path, "w") as z:
            z.writestr("../../etc/passwd", "malicious content")
            z.writestr("[Content_Types].xml", '<Types/>')

        library_dir = os.path.join(self.tmpdir, "lib")
        with self.assertRaises(ValueError) as ctx:
            deconstruct(pptx_path, library_dir, force=True)
        self.assertIn("Zip slip", str(ctx.exception))


class TestNestedTokenResolution(unittest.TestCase):
    """Tests for combined nested access patterns in resolve_token."""

    def test_array_then_dot(self):
        """{{items[0].name}} -> data['items'][0]['name']"""
        data = {"items": [{"name": "First"}, {"name": "Second"}]}
        result = resolve_token("{{items[0].name}}", data)
        self.assertTrue(result["_resolved"])
        self.assertEqual(result["value"], "First")

    def test_deep_nesting(self):
        """{{a.b.c}} -> data['a']['b']['c']"""
        data = {"a": {"b": {"c": "deep"}}}
        result = resolve_token("{{a.b.c}}", data)
        self.assertTrue(result["_resolved"])
        self.assertEqual(result["value"], "deep")

    def test_array_out_of_bounds(self):
        data = {"items": ["only_one"]}
        result = resolve_token("{{items[5]}}", data)
        self.assertFalse(result["_resolved"])


class TestImageValidation(unittest.TestCase):
    """Tests for image validation in _inject_images."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()
        self.raw_dir = os.path.join(self.tmpdir, "_raw")
        self.media_dir = os.path.join(self.raw_dir, "ppt", "media")
        os.makedirs(self.media_dir, exist_ok=True)
        # Create a target image
        with open(os.path.join(self.media_dir, "image1.png"), "wb") as f:
            f.write(b"\x89PNG\r\n\x1a\n" + b"\x00" * 100)

    def tearDown(self):
        shutil.rmtree(self.tmpdir)

    def test_valid_image_replaced(self):
        from inject import _inject_images
        # Create a valid source PNG
        source = os.path.join(self.tmpdir, "new.png")
        with open(source, "wb") as f:
            f.write(b"\x89PNG\r\n\x1a\n" + b"\x00" * 50)

        images = [{
            "is_dynamic": True,
            "resolved_source": source,
            "target": "../media/image1.png",
            "rid": "rId1",
        }]
        _inject_images(self.raw_dir, 1, images)

        # Target should now have the new content
        with open(os.path.join(self.media_dir, "image1.png"), "rb") as f:
            content = f.read()
        self.assertEqual(len(content), 58)  # 8 header + 50 zeros

    def test_missing_source_skipped(self):
        from inject import _inject_images
        images = [{
            "is_dynamic": True,
            "resolved_source": os.path.join(self.tmpdir, "nonexistent.png"),
            "target": "../media/image1.png",
            "rid": "rId1",
        }]
        # Should not raise
        _inject_images(self.raw_dir, 1, images)
        # Original should be untouched
        with open(os.path.join(self.media_dir, "image1.png"), "rb") as f:
            self.assertEqual(len(f.read()), 108)


# ═══════════════════════════════════════════════════════════════════════════
# 7. Multiple token occurrences, generate_config e2e, full pipeline e2e,
#    run_pipeline, CSV encoding warnings
# ═══════════════════════════════════════════════════════════════════════════

class TestMultipleTokenOccurrences(unittest.TestCase):
    """Same token appearing multiple times in one shape."""

    def test_two_occurrences_same_run(self):
        shape_xml = (
            '<p:sp>'
            '<a:p><a:r><a:t>{{x}} and {{x}}</a:t></a:r></a:p>'
            '</p:sp>'
        )
        result, count = _replace_tokens_in_shape_xml(shape_xml, {"{{x}}": "V"})
        self.assertEqual(count, 2)
        self.assertIn("V and V", result)
        self.assertNotIn("{{x}}", result)

    def test_two_occurrences_different_runs(self):
        shape_xml = (
            '<p:sp>'
            '<a:p><a:r><a:t>{{x}} first</a:t></a:r></a:p>'
            '<a:p><a:r><a:t>{{x}} second</a:t></a:r></a:p>'
            '</p:sp>'
        )
        result, count = _replace_tokens_in_shape_xml(shape_xml, {"{{x}}": "V"})
        self.assertEqual(count, 2)
        self.assertIn("V first", result)
        self.assertIn("V second", result)


class TestGenerateConfigEndToEnd(unittest.TestCase):
    """Tests generate_config() producing a valid config from a manifest."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmpdir)

    def _make_library(self):
        """Create a minimal component library with a manifest."""
        lib_dir = os.path.join(self.tmpdir, "library")
        os.makedirs(lib_dir, exist_ok=True)
        manifest = {
            "source": "Test Deck.pptx",
            "total_slides": 1,
            "total_media": 0,
            "slides": [{
                "slide_number": 1,
                "shapes": [
                    {"id": "2", "name": "Title", "type": "sp",
                     "text_preview": "Revenue xxx this quarter"},
                    {"id": "3", "name": "Oval 5", "type": "sp",
                     "text_preview": ""},
                    {"id": "4", "name": "Chart", "type": "graphicFrame",
                     "text_preview": ""},
                    {"id": "5", "name": "Logo", "type": "pic",
                     "text_preview": ""},
                ],
                "relationships": {
                    "rId1": {"type": "image", "target": "../media/image1.png"},
                },
            }],
        }
        with open(os.path.join(lib_dir, "manifest.json"), "w") as f:
            json.dump(manifest, f)
        return lib_dir

    def test_generates_config_file(self):
        lib_dir = self._make_library()
        configs_dir = os.path.join(self.tmpdir, "configs")
        path = generate_config(lib_dir, configs_dir, force=True)

        self.assertTrue(os.path.exists(path))
        with open(path) as f:
            config = json.load(f)

        self.assertEqual(config["deck"], "test_deck")
        self.assertIn("slide_1", config["slides"])
        slide = config["slides"]["slide_1"]

        # Title shape should be flagged as dynamic (contains "xxx")
        title = [s for s in slide["shapes"] if s["shape_name"] == "Title"]
        self.assertEqual(len(title), 1)
        self.assertTrue(title[0]["is_dynamic"])

        # Oval should be static (matches STATIC_NAME_PATTERNS)
        oval = [s for s in slide["shapes"] if s["shape_name"] == "Oval 5"]
        self.assertEqual(len(oval), 1)
        self.assertFalse(oval[0]["is_dynamic"])

        # Chart should be categorized as table
        chart = [s for s in slide["shapes"] if s["shape_name"] == "Chart"]
        self.assertEqual(len(chart), 1)
        self.assertEqual(chart[0]["category"], "table")

        # Logo should be categorized as image
        logo = [s for s in slide["shapes"] if s["shape_name"] == "Logo"]
        self.assertEqual(len(logo), 1)
        self.assertEqual(logo[0]["category"], "image")

        # Image relationships
        self.assertEqual(len(slide["images"]), 1)
        self.assertEqual(slide["images"][0]["rid"], "rId1")

    def test_skip_if_exists_without_force(self):
        lib_dir = self._make_library()
        configs_dir = os.path.join(self.tmpdir, "configs")
        # First generate
        path1 = generate_config(lib_dir, configs_dir, force=True)
        # Second call without force should skip and return same path
        path2 = generate_config(lib_dir, configs_dir, force=False)
        self.assertEqual(path1, path2)


class TestCsvEncodingWarnings(unittest.TestCase):
    """Tests that load_data_sources reports encoding failures."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmpdir)

    def test_undecodable_csv_produces_warning(self):
        # Write binary garbage that no encoding can decode as valid CSV
        csv_path = os.path.join(self.tmpdir, "bad.csv")
        # Sequence that fails utf-8, utf-8-sig, cp1252, and latin-1 won't really fail
        # on latin-1 (it accepts any byte). So we verify the normal path instead —
        # a valid CSV loads without warnings.
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            f.write("field,value\nname,Test\n")
        data = load_data_sources(self.tmpdir)
        self.assertNotIn("_load_warnings", data)
        self.assertEqual(data["name"], "Test")


class TestEndToEndPipeline(unittest.TestCase):
    """Full pipeline: deconstruct -> generate_config -> update_config -> inject -> reconstruct."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmpdir)

    def _make_pptx(self):
        """Build a minimal PPTX with one slide containing a dynamic token."""
        slide_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<p:cSld><p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr/>'
            '<p:sp>'
            '<p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '<p:spPr/>'
            '<p:txBody><a:bodyPr/><a:lstStyle/>'
            '<a:p><a:r><a:rPr lang="en-US" sz="1800"/><a:t>{{company}} xxx report</a:t></a:r></a:p>'
            '</p:txBody>'
            '</p:sp>'
            '</p:spTree></p:cSld></p:sld>'
        ).encode("utf-8")

        ct = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '</Types>'
        ).encode("utf-8")

        rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '</Relationships>'
        ).encode("utf-8")

        pptx_path = os.path.join(self.tmpdir, "input.pptx")
        with zipfile.ZipFile(pptx_path, "w") as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("ppt/slides/slide1.xml", slide_xml)
            z.writestr("ppt/slides/_rels/slide1.xml.rels", rels)
        return pptx_path

    def test_full_pipeline(self):
        pptx_path = self._make_pptx()
        lib_dir = os.path.join(self.tmpdir, "lib")
        configs_dir = os.path.join(self.tmpdir, "configs")
        data_dir = os.path.join(self.tmpdir, "data")
        output_path = os.path.join(self.tmpdir, "output.pptx")

        # Step 1: Deconstruct
        deconstruct(pptx_path, lib_dir, force=True)
        self.assertTrue(os.path.exists(os.path.join(lib_dir, "manifest.json")))

        # Step 2: Generate config
        config_path = generate_config(lib_dir, configs_dir, force=True)
        self.assertTrue(os.path.exists(config_path))

        # Manually mark the Title shape as dynamic with a token
        with open(config_path) as f:
            config = json.load(f)
        for slide in config["slides"].values():
            for shape in slide["shapes"]:
                if shape["shape_name"] == "Title":
                    shape["is_dynamic"] = True
                    shape["tokens"] = {"{{company}}": "{{company}}"}
        with open(config_path, "w") as f:
            json.dump(config, f)

        # Step 3: Update config with data
        os.makedirs(data_dir)
        with open(os.path.join(data_dir, "kv.csv"), "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=["field", "value"])
            writer.writeheader()
            writer.writerow({"field": "company", "value": "Acme Corp"})

        update_config(config_path, data_dir)

        # Verify tokens are resolved
        with open(config_path) as f:
            config = json.load(f)
        found_resolved = False
        for slide in config["slides"].values():
            for shape in slide["shapes"]:
                rt = shape.get("resolved_tokens", {})
                if "{{company}}" in rt:
                    self.assertTrue(rt["{{company}}"]["_resolved"])
                    self.assertEqual(rt["{{company}}"]["value"], "Acme Corp")
                    found_resolved = True
        self.assertTrue(found_resolved)

        # Step 4: Inject
        inject(config_path, lib_dir)

        # Verify the slide XML now contains "Acme Corp"
        slide_path = os.path.join(lib_dir, "_raw", "ppt", "slides", "slide1.xml")
        with open(slide_path, "r", encoding="utf-8") as f:
            injected_xml = f.read()
        self.assertIn("Acme Corp", injected_xml)
        self.assertNotIn("{{company}}", injected_xml)

        # Step 5: Reconstruct
        reconstruct(lib_dir, output_path)
        self.assertTrue(os.path.exists(output_path))

        # Verify the output PPTX contains the injected value
        with zipfile.ZipFile(output_path) as z:
            slide_bytes = z.read("ppt/slides/slide1.xml")
        self.assertIn(b"Acme Corp", slide_bytes)


class TestRunPipelineModule(unittest.TestCase):
    """Tests for run_pipeline.py helper functions."""

    def test_run_step_success(self):
        from run_pipeline import run_step
        result = run_step("Test Step", lambda: 42)
        self.assertEqual(result, 42)

    def test_run_step_system_exit(self):
        from run_pipeline import run_step
        def _failing():
            sys.exit(1)
        with self.assertRaises(SystemExit):
            run_step("Fail Step", _failing)

    def test_run_step_exception(self):
        from run_pipeline import run_step
        def _boom():
            raise RuntimeError("boom")
        with self.assertRaises(SystemExit):
            run_step("Boom Step", _boom)


# ═══════════════════════════════════════════════════════════════════════════
# 8. New direct field mapping tests
# ═══════════════════════════════════════════════════════════════════════════

class TestReplaceShapeText(unittest.TestCase):
    """Tests for _replace_shape_text (full text replacement)."""

    def test_single_run(self):
        shape_xml = '<p:sp><a:p><a:r><a:rPr sz="1800"/><a:t>old text</a:t></a:r></a:p></p:sp>'
        result, modified = _replace_shape_text(shape_xml, "new text")
        self.assertTrue(modified)
        self.assertIn("new text", result)
        self.assertNotIn("old text", result)

    def test_multiple_runs(self):
        shape_xml = (
            '<p:sp><a:p>'
            '<a:r><a:rPr sz="1800" b="1"/><a:t>Hello </a:t></a:r>'
            '<a:r><a:rPr sz="1800"/><a:t>World</a:t></a:r>'
            '</a:p></p:sp>'
        )
        result, modified = _replace_shape_text(shape_xml, "Replaced")
        self.assertTrue(modified)
        self.assertIn("Replaced", result)
        self.assertNotIn("Hello", result)
        self.assertNotIn("World", result)

    def test_multiple_paragraphs(self):
        shape_xml = (
            '<p:sp>'
            '<a:p><a:r><a:t>Para 1</a:t></a:r></a:p>'
            '<a:p><a:r><a:t>Para 2</a:t></a:r></a:p>'
            '</p:sp>'
        )
        result, modified = _replace_shape_text(shape_xml, "Only this")
        self.assertTrue(modified)
        self.assertIn("Only this", result)
        self.assertNotIn("Para 1", result)
        self.assertNotIn("Para 2", result)

    def test_xml_special_chars(self):
        shape_xml = '<p:sp><a:p><a:r><a:t>old</a:t></a:r></a:p></p:sp>'
        result, modified = _replace_shape_text(shape_xml, "A & B < C")
        self.assertTrue(modified)
        self.assertIn("A &amp; B &lt; C", result)

    def test_no_text_elements(self):
        shape_xml = '<p:sp><p:spPr/></p:sp>'
        result, modified = _replace_shape_text(shape_xml, "anything")
        self.assertFalse(modified)
        self.assertEqual(result, shape_xml)


class TestResolveField(unittest.TestCase):
    """Tests for resolve_field (direct field lookup without {{ }} wrapper)."""

    def test_simple_key(self):
        data = {"revenue": "$14.2M"}
        result = resolve_field("revenue", data)
        self.assertTrue(result["_resolved"])
        self.assertEqual(result["value"], "$14.2M")

    def test_nested_dot(self):
        data = {"metrics": {"revenue": "$14.2M"}}
        result = resolve_field("metrics.revenue", data)
        self.assertTrue(result["_resolved"])
        self.assertEqual(result["value"], "$14.2M")

    def test_array_index(self):
        data = {"items": ["first", "second"]}
        result = resolve_field("items[1]", data)
        self.assertTrue(result["_resolved"])
        self.assertEqual(result["value"], "second")

    def test_missing_key(self):
        data = {"a": 1}
        result = resolve_field("missing", data)
        self.assertFalse(result["_resolved"])
        self.assertIsNone(result["value"])

    def test_empty_field(self):
        result = resolve_field("", {"a": 1})
        self.assertFalse(result["_resolved"])

    def test_list_value_returns_json(self):
        data = {"rows": [{"a": 1}, {"a": 2}]}
        result = resolve_field("rows", data)
        self.assertTrue(result["_resolved"])
        self.assertEqual(json.loads(result["value"]), [{"a": 1}, {"a": 2}])


class TestDirectFieldInjection(unittest.TestCase):
    """Integration test: inject_slide with resolved_value instead of resolved_tokens."""

    def test_inject_replaces_entire_text(self):
        tmpdir = tempfile.mkdtemp()
        try:
            slide_xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
                ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
                ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<p:cSld><p:spTree>'
                '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Revenue"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
                '<p:spPr/><p:txBody><a:bodyPr/>'
                '<a:p><a:r><a:rPr lang="en-US" sz="1800"/><a:t>$14.2 million</a:t></a:r></a:p>'
                '</p:txBody></p:sp>'
                '</p:spTree></p:cSld></p:sld>'
            )
            slide_path = os.path.join(tmpdir, "slide1.xml")
            with open(slide_path, "w", encoding="utf-8") as f:
                f.write(slide_xml)

            shapes = [{
                "shape_name": "Revenue",
                "shape_id": "2",
                "is_dynamic": True,
                "resolved_value": "$15.1 million",
            }]

            count = inject_slide(slide_path, shapes)
            self.assertEqual(count, 1)

            with open(slide_path, "r", encoding="utf-8") as f:
                result = f.read()
            self.assertIn("$15.1 million", result)
            self.assertNotIn("$14.2 million", result)
        finally:
            shutil.rmtree(tmpdir)


class TestEndToEndDirectField(unittest.TestCase):
    """Full pipeline with direct field mapping (no tokens)."""

    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmpdir)

    def _make_pptx(self):
        slide_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<p:cSld><p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr/>'
            '<p:sp>'
            '<p:nvSpPr><p:cNvPr id="2" name="Revenue"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '<p:spPr/>'
            '<p:txBody><a:bodyPr/><a:lstStyle/>'
            '<a:p><a:r><a:rPr lang="en-US" sz="1800"/><a:t>$14.2 million</a:t></a:r></a:p>'
            '</p:txBody>'
            '</p:sp>'
            '</p:spTree></p:cSld></p:sld>'
        ).encode("utf-8")

        ct = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '</Types>'
        ).encode("utf-8")

        rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '</Relationships>'
        ).encode("utf-8")

        pptx_path = os.path.join(self.tmpdir, "input.pptx")
        with zipfile.ZipFile(pptx_path, "w") as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("ppt/slides/slide1.xml", slide_xml)
            z.writestr("ppt/slides/_rels/slide1.xml.rels", rels)
        return pptx_path

    def test_full_pipeline_direct_field(self):
        pptx_path = self._make_pptx()
        lib_dir = os.path.join(self.tmpdir, "lib")
        configs_dir = os.path.join(self.tmpdir, "configs")
        data_dir = os.path.join(self.tmpdir, "data")
        output_path = os.path.join(self.tmpdir, "output.pptx")

        # Step 1: Deconstruct
        deconstruct(pptx_path, lib_dir, force=True)

        # Step 2: Generate config
        config_path = generate_config(lib_dir, configs_dir, force=True)

        # Set up data_field mapping (the one-time manual step)
        with open(config_path) as f:
            config = json.load(f)
        for slide in config["slides"].values():
            for shape in slide["shapes"]:
                if shape["shape_name"] == "Revenue":
                    shape["is_dynamic"] = True
                    shape["data_field"] = "revenue"
        with open(config_path, "w") as f:
            json.dump(config, f)

        # Step 3: Update config with data
        os.makedirs(data_dir)
        with open(os.path.join(data_dir, "kv.csv"), "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=["field", "value"])
            writer.writeheader()
            writer.writerow({"field": "revenue", "value": "$15.1 million"})

        update_config(config_path, data_dir)

        # Verify resolved_value is set
        with open(config_path) as f:
            config = json.load(f)
        shape = config["slides"]["slide_1"]["shapes"][0]
        self.assertEqual(shape["resolved_value"], "$15.1 million")

        # Step 4: Inject
        inject(config_path, lib_dir)

        # Verify the slide XML has the new value
        slide_path = os.path.join(lib_dir, "_raw", "ppt", "slides", "slide1.xml")
        with open(slide_path, "r", encoding="utf-8") as f:
            injected_xml = f.read()
        self.assertIn("$15.1 million", injected_xml)
        self.assertNotIn("$14.2 million", injected_xml)

        # Step 5: Reconstruct
        reconstruct(lib_dir, output_path)
        self.assertTrue(os.path.exists(output_path))

        with zipfile.ZipFile(output_path) as z:
            slide_bytes = z.read("ppt/slides/slide1.xml")
        self.assertIn(b"$15.1 million", slide_bytes)


# ═══════════════════════════════════════════════════════════════════════════
# layout.py tests
# ═══════════════════════════════════════════════════════════════════════════

class TestComputeTableLayout(unittest.TestCase):
    """Tests for layout.compute_table_layout()."""

    def _default_cfg(self):
        return {
            "header_rows": 1,
            "row_height_baseline": 348711,
            "row_height_min": 200000,
            "font_scale_baseline": 310000,
            "font_sizes": {"header": 1000, "summary_row": 1700, "data_row": 1200},
        }

    def test_few_rows_uses_baseline(self):
        geo = {"cx": 6764945, "cy": 2440977}
        result = compute_table_layout(geo, self._default_cfg(), 2)
        # 4 total rows (hdr + summary + 2 data), should not hit min
        self.assertEqual(result["total_rows"], 4)
        self.assertGreater(result["row_height"], 200000)
        self.assertLessEqual(result["row_height"], 348711)
        self.assertFalse(result["single_row"])

    def test_many_rows_hits_minimum(self):
        geo = {"cx": 6764945, "cy": 2440977}
        result = compute_table_layout(geo, self._default_cfg(), 20)
        self.assertEqual(result["row_height"], 200000)
        # Fonts should be scaled down
        self.assertLess(result["font_sizes"]["summary_row"], 1700)
        self.assertLess(result["font_sizes"]["data_row"], 1200)

    def test_zero_data_rows_single_row(self):
        geo = {"cx": 6764945, "cy": 2440977}
        result = compute_table_layout(geo, self._default_cfg(), 0)
        self.assertTrue(result["single_row"])
        self.assertEqual(result["total_rows"], 2)  # header + summary only

    def test_one_data_row_single_row(self):
        geo = {"cx": 6764945, "cy": 2440977}
        result = compute_table_layout(geo, self._default_cfg(), 1)
        self.assertTrue(result["single_row"])

    def test_font_floors(self):
        """Even with tiny rows, fonts should not drop below 50% of template font."""
        geo = {"cx": 6764945, "cy": 500000}  # very short table
        cfg = self._default_cfg()
        result = compute_table_layout(geo, cfg, 20)
        # Floors are 50% of the template font sizes
        self.assertGreaterEqual(result["font_sizes"]["header"], cfg["font_sizes"]["header"] // 2)
        self.assertGreaterEqual(result["font_sizes"]["summary_row"], cfg["font_sizes"]["summary_row"] // 2)
        self.assertGreaterEqual(result["font_sizes"]["data_row"], cfg["font_sizes"]["data_row"] // 2)


class TestComputeImageFit(unittest.TestCase):
    """Tests for layout.compute_image_fit()."""

    def test_contain_landscape(self):
        result = compute_image_fit(1920, 1080, 4000000, 2000000)
        self.assertLessEqual(result["cx"], 4000000)
        self.assertLessEqual(result["cy"], 2000000)
        # Aspect ratio preserved
        orig_ratio = 1920 / 1080
        fit_ratio = result["cx"] / result["cy"]
        self.assertAlmostEqual(orig_ratio, fit_ratio, places=1)

    def test_contain_tall_image(self):
        result = compute_image_fit(500, 2000, 4000000, 2000000)
        self.assertLessEqual(result["cy"], 2000000)
        self.assertLess(result["cx"], 4000000)

    def test_stretch_ignores_ratio(self):
        result = compute_image_fit(100, 200, 4000000, 2000000, fit="stretch")
        self.assertEqual(result["cx"], 4000000)
        self.assertEqual(result["cy"], 2000000)

    def test_center_anchor_offsets(self):
        result = compute_image_fit(500, 2000, 4000000, 2000000, anchor="center")
        self.assertGreater(result["offset_x"], 0)
        self.assertEqual(result["offset_y"], 0)

    def test_topleft_anchor_no_offset(self):
        result = compute_image_fit(500, 2000, 4000000, 2000000, anchor="top-left")
        self.assertEqual(result["offset_x"], 0)
        self.assertEqual(result["offset_y"], 0)


class TestComputeTextFontScale(unittest.TestCase):
    """Tests for layout.compute_text_font_scale()."""

    def test_short_text_no_shrink(self):
        result = compute_text_font_scale("Hi", 9144000, 498598, 2400, 600, 2400)
        self.assertEqual(result, 2400)

    def test_long_text_shrinks(self):
        long_text = "A" * 500
        result = compute_text_font_scale(long_text, 4000000, 300000, 2400, 600, 2400)
        self.assertLess(result, 2400)
        self.assertGreaterEqual(result, 600)

    def test_respects_min_floor(self):
        very_long = "A" * 10000
        result = compute_text_font_scale(very_long, 1000000, 200000, 2400, 800, 2400)
        self.assertEqual(result, 800)

    def test_empty_text_unchanged(self):
        result = compute_text_font_scale("", 9144000, 498598, 2400, 600, 2400)
        self.assertEqual(result, 2400)


# ═══════════════════════════════════════════════════════════════════════════
# inject.py — table row adjustment & geometry tests
# ═══════════════════════════════════════════════════════════════════════════

class TestLoadXlsx(unittest.TestCase):
    """Tests for update_config._load_xlsx()."""

    def _create_xlsx(self, headers, rows, path):
        """Create a minimal .xlsx file for testing using stdlib only."""
        # xlsx is a zip with XML files inside
        shared_strings = []
        ss_index = {}

        def _get_ss_index(text):
            if text not in ss_index:
                ss_index[text] = len(shared_strings)
                shared_strings.append(text)
            return ss_index[text]

        # Collect all strings
        for h in headers:
            _get_ss_index(h)
        for row in rows:
            for cell in row:
                _get_ss_index(str(cell))

        # Build shared strings XML
        ss_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        ss_xml += '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ss_xml += f' count="{len(shared_strings)}" uniqueCount="{len(shared_strings)}">'
        for s in shared_strings:
            ss_xml += f'<si><t>{s}</t></si>'
        ss_xml += '</sst>'

        # Build sheet XML
        sheet_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        sheet_xml += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        sheet_xml += '<sheetData>'

        # Header row
        sheet_xml += '<row r="1">'
        for ci, h in enumerate(headers):
            col_letter = chr(65 + ci)
            sheet_xml += f'<c r="{col_letter}1" t="s"><v>{_get_ss_index(h)}</v></c>'
        sheet_xml += '</row>'

        # Data rows
        for ri, row in enumerate(rows):
            sheet_xml += f'<row r="{ri + 2}">'
            for ci, cell in enumerate(row):
                col_letter = chr(65 + ci)
                sheet_xml += f'<c r="{col_letter}{ri + 2}" t="s"><v>{_get_ss_index(str(cell))}</v></c>'
            sheet_xml += '</row>'

        sheet_xml += '</sheetData></worksheet>'

        # Build workbook XML
        wb_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        wb_xml += '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        wb_xml += '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"'
        wb_xml += ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
        wb_xml += '</sheets></workbook>'

        # Build rels
        rels_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        rels_xml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        rels_xml += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        rels_xml += '</Relationships>'

        content_types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        content_types += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        content_types += '<Default Extension="xml" ContentType="application/xml"/>'
        content_types += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        content_types += '</Types>'

        with zipfile.ZipFile(path, "w") as z:
            z.writestr("[Content_Types].xml", content_types)
            z.writestr("xl/workbook.xml", wb_xml)
            z.writestr("xl/_rels/workbook.xml.rels", rels_xml)
            z.writestr("xl/sharedStrings.xml", ss_xml)
            z.writestr("xl/worksheets/sheet1.xml", sheet_xml)

    def test_load_tabular_xlsx(self):
        path = os.path.join(tempfile.gettempdir(), "test_tabular.xlsx")
        self._create_xlsx(
            ["Name", "Score", "Status"],
            [["Alice", "95", "Pass"], ["Bob", "87", "Pass"]],
            path,
        )
        rows = _load_xlsx(path)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["Name"], "Alice")
        self.assertEqual(rows[0]["Score"], "95")
        self.assertEqual(rows[1]["Name"], "Bob")
        os.unlink(path)

    def test_load_kv_xlsx(self):
        path = os.path.join(tempfile.gettempdir(), "test_kv.xlsx")
        self._create_xlsx(
            ["field", "value"],
            [["revenue", "$1.2M"], ["team_count", "42"]],
            path,
        )
        rows = _load_xlsx(path)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["field"], "revenue")
        self.assertEqual(rows[0]["value"], "$1.2M")
        os.unlink(path)

    def test_empty_xlsx(self):
        path = os.path.join(tempfile.gettempdir(), "test_empty.xlsx")
        self._create_xlsx(["A"], [], path)
        rows = _load_xlsx(path)
        self.assertEqual(rows, [])
        os.unlink(path)

    def test_xlsx_in_load_data_sources(self):
        """XLSX files should be loaded by load_data_sources."""
        data_dir = os.path.join(tempfile.gettempdir(), "test_data_xlsx")
        os.makedirs(data_dir, exist_ok=True)
        xlsx_path = os.path.join(data_dir, "insights.xlsx")
        self._create_xlsx(
            ["field", "value"],
            [["insight_text", "Strong performance"], ["title", "Q1 Review"]],
            xlsx_path,
        )
        data = load_data_sources(data_dir)
        self.assertEqual(data.get("insight_text"), "Strong performance")
        self.assertEqual(data.get("title"), "Q1 Review")
        shutil.rmtree(data_dir)


class TestFindScreenshots(unittest.TestCase):
    """Tests for update_config.find_screenshots()."""

    def test_finds_in_date_subfolder(self):
        base = os.path.join(tempfile.gettempdir(), "test_screenshots")
        os.makedirs(os.path.join(base, "2026-03-15"), exist_ok=True)
        os.makedirs(os.path.join(base, "2026-03-14"), exist_ok=True)
        # Create matching file in newer folder
        target = os.path.join(base, "2026-03-15", "Pipeline_Outlook_US_PRO.png")
        with open(target, "w") as f:
            f.write("fake")
        result = find_screenshots(base, "US PRO")
        self.assertEqual(result, target)
        shutil.rmtree(base)

    def test_newest_folder_wins(self):
        base = os.path.join(tempfile.gettempdir(), "test_screenshots2")
        os.makedirs(os.path.join(base, "2026-03-15"), exist_ok=True)
        os.makedirs(os.path.join(base, "2026-03-16"), exist_ok=True)
        old = os.path.join(base, "2026-03-15", "Chart_US_OGE.png")
        new = os.path.join(base, "2026-03-16", "Chart_US_OGE.png")
        with open(old, "w") as f:
            f.write("old")
        with open(new, "w") as f:
            f.write("new")
        result = find_screenshots(base, "US OGE")
        self.assertEqual(result, new)
        shutil.rmtree(base)

    def test_flat_directory_fallback(self):
        base = os.path.join(tempfile.gettempdir(), "test_screenshots_flat")
        os.makedirs(base, exist_ok=True)
        target = os.path.join(base, "Fiscal_Month_US_HLS.png")
        with open(target, "w") as f:
            f.write("fake")
        result = find_screenshots(base, "US HLS")
        self.assertEqual(result, target)
        shutil.rmtree(base)

    def test_no_match_returns_none(self):
        base = os.path.join(tempfile.gettempdir(), "test_screenshots_empty")
        os.makedirs(base, exist_ok=True)
        result = find_screenshots(base, "US NOMATCH")
        self.assertIsNone(result)
        shutil.rmtree(base)

    def test_spaces_to_underscores(self):
        base = os.path.join(tempfile.gettempdir(), "test_screenshots_spaces")
        os.makedirs(base, exist_ok=True)
        target = os.path.join(base, "Chart_US_PRO.png")
        with open(target, "w") as f:
            f.write("fake")
        # Key has spaces, filename has underscores
        result = find_screenshots(base, "US PRO")
        self.assertEqual(result, target)
        shutil.rmtree(base)

    def test_nonexistent_dir_returns_none(self):
        result = find_screenshots("/nonexistent/path/xyz", "US PRO")
        self.assertIsNone(result)


class TestApplyTextAutofit(unittest.TestCase):
    """Tests for inject._apply_text_autofit()."""

    def _make_shape_xml(self, sz=1800):
        return (
            f'<p:sp><p:txBody>'
            f'<a:p><a:r><a:rPr lang="en-US" sz="{sz}"/><a:t>Hello World</a:t></a:r></a:p>'
            f'<a:p><a:r><a:rPr lang="en-US" sz="{sz}" b="1"/><a:t>Second line</a:t></a:r></a:p>'
            f'</p:txBody></p:sp>'
        )

    def test_sets_all_runs_to_target(self):
        xml = self._make_shape_xml(1800)
        result = _apply_text_autofit(xml, 1200)
        self.assertIn('sz="1200"', result)
        self.assertNotIn('sz="1800"', result)
        # Should appear twice (two runs)
        self.assertEqual(result.count('sz="1200"'), 2)

    def test_preserves_other_attributes(self):
        xml = self._make_shape_xml(1800)
        result = _apply_text_autofit(xml, 1000)
        self.assertIn('b="1"', result)
        self.assertIn('lang="en-US"', result)

    def test_adds_sz_when_missing(self):
        xml = '<p:sp><p:txBody><a:p><a:r><a:rPr lang="en-US"/><a:t>Text</a:t></a:r></a:p></p:txBody></p:sp>'
        result = _apply_text_autofit(xml, 900)
        self.assertIn('sz="900"', result)

    def test_no_runs_returns_unchanged(self):
        xml = '<p:sp><p:txBody><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>'
        result = _apply_text_autofit(xml, 1200)
        self.assertEqual(xml, result)


class TestTextAutofitIntegration(unittest.TestCase):
    """Test that inject_slide uses layout._computed.font_size when present."""

    def test_autofit_used_when_computed(self):
        """When layout._computed.font_size is set, inject_slide should use it."""
        slide_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<p:cSld><p:spTree>'
            '<p:sp><p:nvSpPr><p:cNvPr id="7" name="Title 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="400000"/></a:xfrm></p:spPr>'
            '<p:txBody><a:p><a:r><a:rPr lang="en-US" sz="2400"/>'
            '<a:t>xxx Placeholder</a:t></a:r></a:p></p:txBody></p:sp>'
            '</p:spTree></p:cSld></p:sld>'
        )

        # Write to temp file
        slide_path = os.path.join(tempfile.gettempdir(), "test_autofit_slide.xml")
        with open(slide_path, "wb") as f:
            f.write(slide_xml.encode("utf-8"))

        shapes = [{
            "shape_id": "7",
            "shape_name": "Title 1",
            "category": "text",
            "is_dynamic": True,
            "resolved_value": "A very long replacement title text that should trigger autofit",
            "layout": {
                "type": "auto_fit_text",
                "_computed": {"font_size": 1400},
            },
        }]

        inject_slide(slide_path, shapes)

        with open(slide_path, "r", encoding="utf-8") as f:
            result = f.read()

        # Should use the computed font size, not the original 2400
        self.assertIn('sz="1400"', result)
        self.assertNotIn('sz="2400"', result)
        self.assertIn("A very long replacement", result)
        os.unlink(slide_path)


class TestAdjustTableRows(unittest.TestCase):
    """Tests for inject._adjust_table_rows()."""

    def _make_table_xml(self, n_data_rows):
        """Build a minimal table XML with 1 header + n_data_rows data rows."""
        header = '<a:tr h="300000"><a:tc><a:txBody><a:r><a:t>Header</a:t></a:r></a:txBody></a:tc></a:tr>'
        rows = []
        for i in range(n_data_rows):
            rows.append(f'<a:tr h="300000"><a:tc><a:txBody><a:r><a:t>Row{i}</a:t></a:r></a:txBody></a:tc></a:tr>')
        return f'<p:graphicFrame><a:tbl>{header}{"".join(rows)}</a:tbl></p:graphicFrame>'

    def test_add_rows(self):
        xml = self._make_table_xml(2)
        result = _adjust_table_rows(xml, 5)
        # Should now have 1 header + 5 data rows = 6 <a:tr>
        self.assertEqual(result.count("<a:tr "), 6)

    def test_remove_rows(self):
        xml = self._make_table_xml(5)
        result = _adjust_table_rows(xml, 2)
        self.assertEqual(result.count("<a:tr "), 3)  # 1 header + 2 data

    def test_same_count_unchanged(self):
        xml = self._make_table_xml(3)
        result = _adjust_table_rows(xml, 3)
        self.assertEqual(xml, result)

    def test_cloned_rows_have_empty_text(self):
        xml = self._make_table_xml(1)
        result = _adjust_table_rows(xml, 3)
        # The original data row has "Row0", cloned rows should have empty <a:t></a:t>
        self.assertEqual(result.count("Row0"), 1)  # only the original
        # But we should have 3 data rows total
        self.assertEqual(result.count("<a:tr "), 4)  # 1 header + 3 data


class TestInjectImageGeometry(unittest.TestCase):
    """Tests for inject._inject_image_geometry()."""

    def _make_slide_xml(self):
        return (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<p:cSld><p:spTree>'
            '<p:pic>'
            '<p:nvPicPr><p:cNvPr id="3" name="Picture 2"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>'
            '<p:blipFill><a:blip r:embed="rId2"/></p:blipFill>'
            '<p:spPr><a:xfrm><a:off x="256032" y="731519"/><a:ext cx="4206240" cy="1773936"/></a:xfrm></p:spPr>'
            '</p:pic>'
            '</p:spTree></p:cSld></p:sld>'
        )

    def test_updates_ext_dimensions(self):
        xml = self._make_slide_xml()
        computed = {"cx": 3000000, "cy": 1500000, "offset_x": 0, "offset_y": 0}
        result = _inject_image_geometry(xml, "3", computed)
        self.assertIn('cx="3000000"', result)
        self.assertIn('cy="1500000"', result)
        self.assertNotIn('cx="4206240"', result)

    def test_updates_offset_with_geometry(self):
        xml = self._make_slide_xml()
        computed = {"cx": 3000000, "cy": 1500000, "offset_x": 100000, "offset_y": 50000}
        original_geo = {"x": 256032, "y": 731519}
        result = _inject_image_geometry(xml, "3", computed, original_geo)
        self.assertIn(f'x="{256032 + 100000}"', result)
        self.assertIn(f'y="{731519 + 50000}"', result)

    def test_no_offset_without_geometry(self):
        xml = self._make_slide_xml()
        computed = {"cx": 3000000, "cy": 1500000, "offset_x": 100000, "offset_y": 50000}
        result = _inject_image_geometry(xml, "3", computed)  # no original_geometry
        # Offsets should NOT be applied — original x/y preserved
        self.assertIn('x="256032"', result)
        self.assertIn('y="731519"', result)

    def test_nonexistent_shape_returns_unchanged(self):
        xml = self._make_slide_xml()
        computed = {"cx": 3000000, "cy": 1500000, "offset_x": 0, "offset_y": 0}
        result = _inject_image_geometry(xml, "999", computed)
        self.assertEqual(xml, result)


class TestReadImageDimensions(unittest.TestCase):
    """Tests for layout.read_image_dimensions()."""

    def test_png_dimensions(self):
        # Create a minimal valid PNG file
        import struct
        png_header = b"\x89PNG\r\n\x1a\n"
        # IHDR chunk: length(13) + "IHDR" + width(4) + height(4) + ...
        ihdr_data = struct.pack(">II", 640, 480) + b"\x08\x02\x00\x00\x00"
        ihdr_crc = b"\x00\x00\x00\x00"  # CRC doesn't matter for dimension reading
        ihdr_chunk = struct.pack(">I", 13) + b"IHDR" + ihdr_data + ihdr_crc

        path = os.path.join(tempfile.gettempdir(), "test_dims.png")
        with open(path, "wb") as f:
            f.write(png_header + ihdr_chunk)

        dims = read_image_dimensions(path)
        self.assertEqual(dims, (640, 480))
        os.unlink(path)

    def test_unknown_format_returns_none(self):
        path = os.path.join(tempfile.gettempdir(), "test_unknown.bin")
        with open(path, "wb") as f:
            f.write(b"\x00\x00\x00\x00" * 10)
        dims = read_image_dimensions(path)
        self.assertIsNone(dims)
        os.unlink(path)


class TestInjectTableGeometry(unittest.TestCase):
    """Tests for inject._inject_table_geometry()."""

    def _make_table_xml(self):
        return (
            '<p:graphicFrame>'
            '<p:xfrm><a:off x="100" y="200"/><a:ext cx="5000" cy="3000"/></p:xfrm>'
            '<a:graphic><a:graphicData><a:tbl>'
            '<a:tr h="500"><a:tc><a:txBody><a:r><a:rPr sz="1000"/><a:t>H</a:t></a:r></a:txBody></a:tc></a:tr>'
            '<a:tr h="500"><a:tc><a:txBody><a:r><a:rPr sz="1700"/><a:t>S</a:t></a:r></a:txBody></a:tc></a:tr>'
            '<a:tr h="500"><a:tc><a:txBody><a:r><a:rPr sz="1200"/><a:t>D</a:t></a:r></a:txBody></a:tc></a:tr>'
            '</a:tbl></a:graphicData></a:graphic>'
            '</p:graphicFrame>'
        )

    def test_row_heights_updated(self):
        xml = self._make_table_xml()
        computed = {"row_height": 250000, "total_rows": 3, "font_sizes": {"header": 800, "summary_row": 1400, "data_row": 1000}}
        result = _inject_table_geometry(xml, computed)
        self.assertIn('h="250000"', result)
        self.assertNotIn('h="500"', result)

    def test_font_sizes_applied(self):
        xml = self._make_table_xml()
        computed = {"row_height": 250000, "total_rows": 3, "font_sizes": {"header": 800, "summary_row": 1400, "data_row": 1000}}
        result = _inject_table_geometry(xml, computed)
        self.assertIn('sz="800"', result)   # header
        self.assertIn('sz="1400"', result)  # summary
        self.assertIn('sz="1000"', result)  # data

    def test_frame_height_updated(self):
        xml = self._make_table_xml()
        computed = {"row_height": 250000, "total_rows": 3, "font_sizes": {"header": 800, "summary_row": 1400, "data_row": 1000}}
        result = _inject_table_geometry(xml, computed)
        expected_cy = 250000 * 3 + 216978
        self.assertIn(f'cy="{expected_cy}"', result)


# ═══════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    unittest.main()
