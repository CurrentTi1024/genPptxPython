from __future__ import annotations

import json
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from pptx import Presentation

from genppt.config import GeneratorConfig
from genppt.data import normalize_table
from genppt.errors import TemplateError
from genppt.service import generate_presentation


ROOT = Path(__file__).resolve().parents[1]
FIXTURES = ROOT / "tests" / "fixtures"
CONFIG = GeneratorConfig()


class NormalizeTableTest(unittest.TestCase):
    def test_matrix_rows(self) -> None:
        result = normalize_table({"columns": ["A", "B"], "rows": [[1, 2]]}, CONFIG)
        self.assertEqual(result.labels, ["A", "B"])
        self.assertEqual(result.rows, [["1", "2"]])

    def test_object_rows(self) -> None:
        result = normalize_table(
            {"columns": ["name", "count"], "rows": [{"name": "华东", "count": 3}]},
            CONFIG,
        )
        self.assertEqual(result.rows, [["华东", "3"]])

    def test_row_width_mismatch_is_rejected(self) -> None:
        with self.assertRaises(TemplateError):
            normalize_table({"columns": ["A", "B"], "rows": [[1]]}, CONFIG)

    def test_non_finite_number_is_rejected(self) -> None:
        with self.assertRaises(TemplateError):
            normalize_table({"columns": ["A"], "rows": [[float("nan")]]}, CONFIG)

    def test_resource_limit_is_enforced(self) -> None:
        with self.assertRaises(TemplateError):
            normalize_table(
                {"columns": ["A", "B"], "rows": []},
                GeneratorConfig(max_columns=1),
            )


class GeneratePresentationTest(unittest.TestCase):
    def test_test_template_generates_dynamic_tables(self) -> None:
        data = json.loads((ROOT / "examples/test_data.json").read_text(encoding="utf-8"))
        with TemporaryDirectory() as directory:
            output = Path(directory) / "generated.pptx"
            generate_presentation(FIXTURES / "test2.pptx", data, output)
            presentation = Presentation(output)
            tables = [
                shape
                for slide in presentation.slides
                for shape in slide.shapes
                if shape.has_table
            ]
            self.assertEqual(len(presentation.slides), 2)
            self.assertEqual(len(tables), 3)
            self.assertEqual(len(tables[0].table.columns), 3)
            self.assertEqual(len(tables[0].table.rows), 3)
            self.assertEqual(tables[0].table.cell(0, 0).text, "区域")
            self.assertEqual(tables[0].table.cell(1, 0).text, "华东区")
            self.assertTrue(
                any(
                    "分类：资料不完整" in shape.text
                    for shape in presentation.slides[0].shapes
                    if getattr(shape, "has_text_frame", False)
                )
            )
            self.assertTrue(
                all(
                    "<foreach" not in shape.name
                    for slide in presentation.slides
                    for shape in slide.shapes
                )
            )

    def test_columns_can_grow_beyond_template_width(self) -> None:
        data = {
            "tables": [
                {
                    "title": "九列表格",
                    "tableName": "nine_columns",
                    "columns": [f"C{index}" for index in range(1, 10)],
                    "rows": [[f"V{index}" for index in range(1, 10)]],
                }
            ]
        }
        with TemporaryDirectory() as directory:
            output = Path(directory) / "generated.pptx"
            generate_presentation(FIXTURES / "test2.pptx", data, output)
            presentation = Presentation(output)
            tables = [shape for shape in presentation.slides[0].shapes if shape.has_table]
            self.assertEqual(len(tables), 1)
            self.assertEqual(len(tables[0].table.columns), 9)
            self.assertEqual(len(tables[0].table.rows), 2)
            self.assertEqual(tables[0].table.cell(1, 8).text, "V9")

    def test_two_short_tables_share_a_compact_page(self) -> None:
        data = {
            "tables": [
                {"title": "A", "tableName": "a", "columns": ["列"], "rows": [[1]]},
                {"title": "B", "tableName": "b", "columns": ["列"], "rows": [[2]]},
            ]
        }
        with TemporaryDirectory() as directory:
            output = Path(directory) / "generated.pptx"
            generate_presentation(FIXTURES / "test2.pptx", data, output)
            presentation = Presentation(output)
            tables = [shape for shape in presentation.slides[0].shapes if shape.has_table]
            self.assertEqual(len(presentation.slides), 1)
            self.assertEqual([shape.name for shape in tables], ["a", "b"])

    def test_next_table_moves_whole_to_new_page_when_remainder_is_too_small(self) -> None:
        data = {
            "tables": [
                {
                    "title": "第一张",
                    "tableName": "first",
                    "columns": ["序号", "值"],
                    "rows": [[index, index] for index in range(1, 8)],
                },
                {
                    "title": "第二张",
                    "tableName": "second",
                    "columns": ["序号", "值"],
                    "rows": [[index, index] for index in range(1, 4)],
                },
            ]
        }
        with TemporaryDirectory() as directory:
            output = Path(directory) / "generated.pptx"
            generate_presentation(FIXTURES / "test2.pptx", data, output)
            presentation = Presentation(output)
            tables_by_slide = [
                [shape for shape in slide.shapes if shape.has_table]
                for slide in presentation.slides
            ]
            self.assertEqual(len(presentation.slides), 2)
            self.assertEqual([shape.name for shape in tables_by_slide[0]], ["first"])
            self.assertEqual([shape.name for shape in tables_by_slide[1]], ["second"])
            self.assertEqual(len(tables_by_slide[1][0].table.rows), 4)

    def test_table_larger_than_fresh_page_is_split_with_repeated_header(self) -> None:
        data = {
            "tables": [
                {
                    "title": "长表",
                    "tableName": "long_table",
                    "columns": ["序号", "值"],
                    "rows": [[index, f"V{index}"] for index in range(1, 26)],
                }
            ]
        }
        with TemporaryDirectory() as directory:
            output = Path(directory) / "generated.pptx"
            generate_presentation(FIXTURES / "test2.pptx", data, output)
            presentation = Presentation(output)
            tables = [
                shape
                for slide in presentation.slides
                for shape in slide.shapes
                if shape.has_table
            ]
            self.assertGreater(len(tables), 1)
            self.assertTrue(all(table.table.cell(0, 0).text == "序号" for table in tables))
            self.assertEqual(sum(len(table.table.rows) - 1 for table in tables), 25)

    def test_multiple_template_slides_are_rejected_by_default(self) -> None:
        data = {"tables": [{"columns": ["A"], "rows": [[1]]}]}
        with TemporaryDirectory() as directory:
            with self.assertRaises(TemplateError):
                generate_presentation(FIXTURES / "test.pptx", data, Path(directory) / "out.pptx")

    def test_template_cannot_be_overwritten(self) -> None:
        with self.assertRaises(TemplateError):
            generate_presentation(
                FIXTURES / "test2.pptx", {"tables": []}, FIXTURES / "test2.pptx"
            )

    def test_non_table_foreach_is_not_claimed_by_table_generator(self) -> None:
        with TemporaryDirectory() as directory:
            with self.assertRaisesRegex(TemplateError, "no foreach table blocks"):
                generate_presentation(
                    FIXTURES / "test-img.pptx",
                    {"imgs": [], "tables": []},
                    Path(directory) / "out.pptx",
                )

    def test_existing_output_survives_validation_failure(self) -> None:
        with TemporaryDirectory() as directory:
            output = Path(directory) / "existing.pptx"
            output.write_bytes(b"existing-output")
            with self.assertRaises(TemplateError):
                generate_presentation(
                    FIXTURES / "test2.pptx",
                    {"tables": [{"columns": ["A", "B"], "rows": [[1]]}]},
                    output,
                )
            self.assertEqual(output.read_bytes(), b"existing-output")


if __name__ == "__main__":
    unittest.main()
