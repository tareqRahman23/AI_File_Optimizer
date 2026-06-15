from __future__ import annotations

import os
from concurrent.futures import ProcessPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable


@dataclass(frozen=True)
class ConversionResult:
    source: Path
    output: Path | None
    error: str | None = None
    source_tokens_estimate: int = 0
    markdown_tokens_estimate: int = 0
    tokens_saved_estimate: int = 0
    token_savings_percent: float = 0.0

    @property
    def succeeded(self) -> bool:
        return self.error is None


def estimate_tokens(text_or_size: str | int) -> int:
    """Estimate token count with the common rough 4 characters/bytes per token rule."""
    length = text_or_size if isinstance(text_or_size, int) else len(text_or_size)
    return max(1, round(length / 4)) if length else 0


def unique_output_path(source: Path, output_dir: Path, reserved: set[Path]) -> Path:
    """Return an unused Markdown path without overwriting existing files."""
    candidate = output_dir / f"{source.stem}.md"
    suffix = 2

    while candidate.exists() or candidate in reserved:
        candidate = output_dir / f"{source.stem}_{suffix}.md"
        suffix += 1

    reserved.add(candidate)
    return candidate


def plan_outputs(files: Iterable[Path], output_dir: Path) -> dict[Path, Path]:
    reserved: set[Path] = set()
    return {
        source: unique_output_path(source, output_dir, reserved)
        for source in files
    }


def _markdown_text(result: object) -> str:
    for attribute in ("text_content", "markdown"):
        value = getattr(result, attribute, None)
        if isinstance(value, str):
            return value
    raise RuntimeError("MarkItDown returned no Markdown text.")


def default_worker_count(file_count: int) -> int:
    if file_count <= 0:
        return 0
    return max(1, min(file_count, os.cpu_count() or 1))


def _convert_one_file(source: Path, output: Path) -> ConversionResult:
    try:
        from markitdown import MarkItDown
    except ImportError as exc:
        raise RuntimeError(
            "MarkItDown is not installed in the app environment. Close this "
            "window and run install.bat. Use Python 3.10 through 3.13; "
            "Python 3.14 is not currently supported by all dependencies."
        ) from exc

    try:
        converter = MarkItDown(enable_plugins=False)
        conversion = converter.convert_local(str(source))
        markdown = _markdown_text(conversion)
        output.write_text(markdown, encoding="utf-8")
        source_tokens = estimate_tokens(source.stat().st_size)
        markdown_tokens = estimate_tokens(markdown)
        tokens_saved = max(0, source_tokens - markdown_tokens)
        return ConversionResult(
            source=source,
            output=output,
            source_tokens_estimate=source_tokens,
            markdown_tokens_estimate=markdown_tokens,
            tokens_saved_estimate=tokens_saved,
            token_savings_percent=(
                tokens_saved / source_tokens * 100 if source_tokens else 0.0
            ),
        )
    except Exception as exc:
        return ConversionResult(source=source, output=None, error=str(exc))


def convert_files(
    files: list[Path],
    output_dir: Path,
    on_status: Callable[[Path, str], None] | None = None,
    max_workers: int | None = None,
) -> list[ConversionResult]:
    try:
        import markitdown  # noqa: F401
    except ImportError as exc:
        raise RuntimeError(
            "MarkItDown is not installed in the app environment. Close this "
            "window and run install.bat. Use Python 3.10 through 3.13; "
            "Python 3.14 is not currently supported by all dependencies."
        ) from exc

    output_dir.mkdir(parents=True, exist_ok=True)
    outputs = plan_outputs(files, output_dir)
    worker_count = max_workers or default_worker_count(len(files))
    results: list[ConversionResult | None] = [None] * len(files)

    if worker_count <= 1:
        for index, source in enumerate(files):
            if on_status:
                on_status(source, "Converting")
            result = _convert_one_file(source, outputs[source])
            results[index] = result
            if on_status:
                on_status(source, "Done" if result.succeeded else "Failed")
        return [result for result in results if result is not None]

    with ProcessPoolExecutor(max_workers=worker_count) as executor:
        futures = {}
        for index, source in enumerate(files):
            if on_status:
                on_status(source, "Converting")
            future = executor.submit(_convert_one_file, source, outputs[source])
            futures[future] = (index, source)

        for future in as_completed(futures):
            index, source = futures[future]
            try:
                result = future.result()
            except Exception as exc:
                result = ConversionResult(source=source, output=None, error=str(exc))
            results[index] = result
            if on_status:
                on_status(source, "Done" if result.succeeded else "Failed")

    return [result for result in results if result is not None]
