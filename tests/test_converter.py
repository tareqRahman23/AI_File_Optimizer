import tempfile
import unittest
from types import ModuleType, SimpleNamespace
from pathlib import Path
from unittest.mock import patch

from markitdown_gui.converter import (
    convert_files,
    default_worker_count,
    estimate_tokens,
    plan_outputs,
    unique_output_path,
)


class OutputPathTests(unittest.TestCase):
    def test_uses_markdown_extension(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory)
            result = unique_output_path(Path("report.pdf"), output, set())
            self.assertEqual(result, output / "report.md")

    def test_does_not_overwrite_existing_file(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory)
            (output / "report.md").touch()
            result = unique_output_path(Path("report.pdf"), output, set())
            self.assertEqual(result, output / "report_2.md")

    def test_duplicate_stems_get_unique_names(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory)
            first = Path("a/report.pdf")
            second = Path("b/report.xlsx")
            plan = plan_outputs([first, second], output)
            self.assertEqual(plan[first], output / "report.md")
            self.assertEqual(plan[second], output / "report_2.md")


class TokenEstimateTests(unittest.TestCase):
    def test_token_estimate_uses_four_character_rule(self) -> None:
        self.assertEqual(estimate_tokens("abcdefghijkl"), 3)
        self.assertEqual(estimate_tokens(40), 10)
        self.assertEqual(estimate_tokens(""), 0)

    def test_default_worker_count_uses_available_file_count(self) -> None:
        self.assertEqual(default_worker_count(0), 0)
        self.assertEqual(default_worker_count(1), 1)
        self.assertGreaterEqual(default_worker_count(2), 1)


class BatchConversionTests(unittest.TestCase):
    def test_failure_does_not_stop_remaining_files(self) -> None:
        fake_module = ModuleType("markitdown")

        class FakeMarkItDown:
            def __init__(self, **_kwargs: object) -> None:
                pass

            def convert_local(self, source: str) -> object:
                if source.endswith("broken.pdf"):
                    raise ValueError("Unreadable document")
                return SimpleNamespace(text_content=f"# {Path(source).stem}")

        fake_module.MarkItDown = FakeMarkItDown

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            good = root / "good.pdf"
            broken = root / "broken.pdf"
            good.touch()
            broken.touch()
            output = root / "markdown"

            with patch.dict("sys.modules", {"markitdown": fake_module}):
                results = convert_files([broken, good], output, max_workers=1)

            self.assertFalse(results[0].succeeded)
            self.assertTrue(results[1].succeeded)
            self.assertGreaterEqual(results[1].tokens_saved_estimate, 0)
            self.assertGreaterEqual(results[1].token_savings_percent, 0)
            self.assertEqual((output / "good.md").read_text(encoding="utf-8"), "# good")


if __name__ == "__main__":
    unittest.main()
