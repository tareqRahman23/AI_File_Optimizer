from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

from markitdown_gui.excel_rule_engine import (
    InferenceError,
    apply_template,
    infer_template,
    load_template,
    save_template,
)


class RuleEngineTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def write_excel(self, name: str, frame: pd.DataFrame) -> Path:
        path = self.root / name
        frame.to_excel(path, index=False)
        return path

    def write_multi_sheet_excel(
        self,
        name: str,
        sheets: dict[str, pd.DataFrame],
    ) -> Path:
        path = self.root / name
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            for sheet_name, frame in sheets.items():
                frame.to_excel(writer, sheet_name=sheet_name, index=False)
        return path

    def test_projection_rename_remove_and_sort(self) -> None:
        input_path = self.write_excel(
            "direct-input.xlsx",
            pd.DataFrame(
                {
                    "Student Name": ["Asha", "Bashir", "Chen"],
                    "Score": [71, 92, 84],
                    "Unused": ["x", "y", "z"],
                }
            ),
        )
        expected_path = self.write_excel(
            "direct-output.xlsx",
            pd.DataFrame(
                {
                    "Name": ["Bashir", "Chen", "Asha"],
                    "Marks": [92, 84, 71],
                }
            ),
        )

        template = infer_template(input_path, expected_path)
        self.assertEqual(template.sheets[0].operation, "project")
        self.assertTrue(any("Removed 1 column" in item for item in template.actions))
        self.assertTrue(any("descending" in item for item in template.actions))

        template_path = save_template(template, self.root / "direct-template.json")
        output_path = apply_template(
            load_template(template_path),
            input_path,
            self.root / "direct-reproduced.xlsx",
        )
        actual = pd.read_excel(output_path)
        expected = pd.read_excel(expected_path)
        pd.testing.assert_frame_equal(actual, expected)

    def test_group_sum_sort_and_style(self) -> None:
        input_path = self.write_excel(
            "group-input.xlsx",
            pd.DataFrame(
                {
                    "Student ID": [101, 101, 102, 102, 103],
                    "Marks": [40, 35, 44, 41, 68],
                    "Subject": ["A", "B", "A", "B", "A"],
                }
            ),
        )
        expected_path = self.write_excel(
            "group-output.xlsx",
            pd.DataFrame(
                {
                    "Student ID": [102, 101, 103],
                    "Total Marks": [85, 75, 68],
                }
            ),
        )
        workbook = load_workbook(expected_path)
        worksheet = workbook.active
        for cell in worksheet[1]:
            cell.font = Font(bold=True, color="FFFFFFFF")
            cell.fill = PatternFill("solid", fgColor="FF1F4E78")
        worksheet.column_dimensions["A"].width = 18
        worksheet.freeze_panes = "A2"
        workbook.save(expected_path)
        workbook.close()

        template = infer_template(input_path, expected_path)
        rule = template.sheets[0]
        self.assertEqual(rule.operation, "group")
        self.assertEqual(rule.group_by[0].source, "Student ID")
        self.assertEqual(rule.aggregations[0].function, "sum")

        output_path = apply_template(
            template,
            input_path,
            self.root / "group-reproduced.xlsx",
        )
        actual = pd.read_excel(output_path)
        expected = pd.read_excel(expected_path)
        pd.testing.assert_frame_equal(actual, expected)

        styled = load_workbook(output_path)
        self.assertTrue(styled.active["A1"].font.bold)
        self.assertEqual(styled.active.freeze_panes, "A2")
        styled.close()

    def test_simple_date_filter(self) -> None:
        input_path = self.write_excel(
            "filter-input.xlsx",
            pd.DataFrame(
                {
                    "Date": pd.to_datetime(
                        ["2018-12-31", "2019-01-01", "2019-01-02"]
                    ),
                    "Open": [8, 9, 10],
                    "Close": [80, 90, 100],
                    "Volume": [800, 900, 1000],
                }
            ),
        )
        expected_path = self.write_excel(
            "filter-output.xlsx",
            pd.DataFrame(
                {
                    "Date": pd.to_datetime(["2019-01-02", "2019-01-01"]),
                    "Close": [100, 90],
                    "Volume": [1000, 900],
                }
            ),
        )

        template = infer_template(input_path, expected_path)
        rule = template.sheets[0]
        self.assertEqual(rule.filters[0].source, "Date")
        self.assertEqual(rule.filters[0].operator, ">=")
        self.assertTrue(any("Filtered Date >=" in item for item in template.actions))

        output_path = apply_template(
            template,
            input_path,
            self.root / "filter-reproduced.xlsx",
        )
        actual = pd.read_excel(output_path)
        expected = pd.read_excel(expected_path)
        pd.testing.assert_frame_equal(actual, expected)

    def test_column_arithmetic_formula_and_excel_formula_export(self) -> None:
        input_path = self.write_excel(
            "formula-input.xlsx",
            pd.DataFrame(
                {
                    "Item": ["Pen", "Book", "Bag"],
                    "Quantity": [3, 2, 1],
                    "Unit Price": [10, 50, 80],
                }
            ),
        )
        expected_path = self.write_excel(
            "formula-output.xlsx",
            pd.DataFrame(
                {
                    "Item": ["Book", "Bag", "Pen"],
                    "Quantity": [2, 1, 3],
                    "Unit Price": [50, 80, 10],
                    "Total": [100, 80, 30],
                }
            ),
        )

        template = infer_template(input_path, expected_path)
        formula = template.sheets[0].formulas[0]
        self.assertEqual(formula.output, "Total")
        self.assertEqual(formula.operator, "*")
        self.assertTrue(any("Calculated Total" in item for item in template.actions))

        template_path = save_template(template, self.root / "formula-template.json")
        output_path = apply_template(
            load_template(template_path),
            input_path,
            self.root / "formula-reproduced.xlsx",
        )
        workbook = load_workbook(output_path, data_only=False)
        worksheet = workbook.active
        self.assertEqual(worksheet["D2"].value, '=IFERROR(B2*C2,"")')
        workbook.close()

    def test_expected_workbook_with_uncached_excel_formulas(self) -> None:
        input_path = self.write_excel(
            "uncached-input.xlsx",
            pd.DataFrame(
                {
                    "Item": ["Pen", "Book"],
                    "Quantity": [3, 2],
                    "Unit Price": [10, 50],
                }
            ),
        )
        expected_path = self.write_excel(
            "uncached-output.xlsx",
            pd.DataFrame(
                {
                    "Item": ["Pen", "Book"],
                    "Quantity": [3, 2],
                    "Unit Price": [10, 50],
                    "Total": [None, None],
                }
            ),
        )
        workbook = load_workbook(expected_path)
        worksheet = workbook.active
        worksheet["D2"] = "=B2*C2"
        worksheet["D3"] = "=B3*C3"
        workbook.save(expected_path)
        workbook.close()

        template = infer_template(input_path, expected_path)
        formula = template.sheets[0].formulas[0]
        self.assertEqual(formula.output, "Total")
        self.assertEqual(formula.operator, "*")

    def test_constant_formula(self) -> None:
        input_path = self.write_excel(
            "constant-formula-input.xlsx",
            pd.DataFrame({"Student ID": [1, 2], "Marks": [80, 90]}),
        )
        expected_path = self.write_excel(
            "constant-formula-output.xlsx",
            pd.DataFrame(
                {
                    "Student ID": [1, 2],
                    "Marks": [80, 90],
                    "Percent": [0.8, 0.9],
                }
            ),
        )
        template = infer_template(input_path, expected_path)
        formula = template.sheets[0].formulas[0]
        self.assertIn(formula.operator, {"*", "/"})

        output_path = apply_template(
            template,
            input_path,
            self.root / "constant-formula-reproduced.xlsx",
        )
        workbook = load_workbook(output_path, data_only=False)
        self.assertTrue(str(workbook.active["C2"].value).startswith("=IFERROR("))
        workbook.close()

    def test_group_then_formula(self) -> None:
        input_path = self.write_excel(
            "group-formula-input.xlsx",
            pd.DataFrame(
                {
                    "Student ID": [1, 1, 2, 2],
                    "Marks": [40, 30, 45, 35],
                    "Possible Marks": [50, 50, 50, 50],
                }
            ),
        )
        expected_path = self.write_excel(
            "group-formula-output.xlsx",
            pd.DataFrame(
                {
                    "Student ID": [2, 1],
                    "Total Marks": [80, 70],
                    "Total Possible": [100, 100],
                    "Percentage": [0.8, 0.7],
                }
            ),
        )
        template = infer_template(input_path, expected_path)
        rule = template.sheets[0]
        self.assertEqual(rule.operation, "group")
        self.assertEqual(rule.formulas[0].output, "Percentage")
        self.assertEqual(rule.formulas[0].operator, "/")

        output_path = apply_template(
            template,
            input_path,
            self.root / "group-formula-reproduced.xlsx",
        )
        workbook = load_workbook(output_path, data_only=False)
        self.assertEqual(workbook.active["D2"].value, '=IFERROR(B2/C2,"")')
        workbook.close()

    def test_unsupported_calculation_explains_failure(self) -> None:
        input_path = self.write_excel(
            "unsupported-input.xlsx",
            pd.DataFrame({"Student ID": [1, 2], "Name": ["Asha", "Bashir"]}),
        )
        expected_path = self.write_excel(
            "unsupported-output.xlsx",
            pd.DataFrame(
                {
                    "Student ID": [1, 2],
                    "Greeting": ["Hello Asha", "Hello Bashir"],
                }
            ),
        )
        with self.assertRaisesRegex(
            InferenceError,
            "No compatible worksheet pair",
        ):
            infer_template(input_path, expected_path)

    def test_analysis_matches_renamed_worksheet(self) -> None:
        input_path = self.write_multi_sheet_excel(
            "renamed-input.xlsx",
            {
                "Raw Data": pd.DataFrame(
                    {"Student ID": [1, 2], "Marks": [80, 90], "Unused": ["x", "y"]}
                )
            },
        )
        expected_path = self.write_multi_sheet_excel(
            "renamed-output.xlsx",
            {
                "Results": pd.DataFrame(
                    {"Student ID": [1, 2], "Marks": [80, 90]}
                )
            },
        )
        template = infer_template(input_path, expected_path)
        self.assertEqual(template.sheets[0].input_sheet, "Raw Data")
        self.assertEqual(template.sheets[0].output_sheet, "Results")
        self.assertTrue(any("Matched renamed sheet" in item for item in template.actions))

    def test_apply_missing_sheet_exports_partial_result_with_warnings(self) -> None:
        example_input = self.write_multi_sheet_excel(
            "partial-example-input.xlsx",
            {
                "Students": pd.DataFrame(
                    {"Student ID": [1, 2], "Marks": [80, 90]}
                ),
                "Sales": pd.DataFrame(
                    {"Item": ["A", "B"], "Quantity": [2, 3]}
                ),
            },
        )
        expected_output = self.write_multi_sheet_excel(
            "partial-example-output.xlsx",
            {
                "Students": pd.DataFrame(
                    {"Student ID": [1, 2], "Marks": [80, 90]}
                ),
                "Sales": pd.DataFrame(
                    {"Item": ["A", "B"], "Quantity": [2, 3]}
                ),
            },
        )
        template = infer_template(example_input, expected_output)
        new_input = self.write_multi_sheet_excel(
            "partial-new-input.xlsx",
            {
                "Students 2027": pd.DataFrame(
                    {"Student ID": [3, 4], "Marks": [70, 85]}
                ),
                "Notes": pd.DataFrame({"Text": ["hello"]}),
            },
        )
        output_path = apply_template(
            template,
            new_input,
            self.root / "partial-result.xlsx",
        )

        workbook = load_workbook(output_path, data_only=False)
        self.assertIn("Students", workbook.sheetnames)
        self.assertNotIn("Sales", workbook.sheetnames)
        self.assertIn("Transformation Warnings", workbook.sheetnames)
        warnings = [
            workbook["Transformation Warnings"].cell(row, 1).value
            for row in range(2, workbook["Transformation Warnings"].max_row + 1)
        ]
        self.assertTrue(any("Used 'Students 2027'" in warning for warning in warnings))
        self.assertTrue(any("Skipped 'Sales'" in warning for warning in warnings))
        self.assertTrue(any("Ignored input sheet 'Notes'" in warning for warning in warnings))
        workbook.close()


if __name__ == "__main__":
    unittest.main()
