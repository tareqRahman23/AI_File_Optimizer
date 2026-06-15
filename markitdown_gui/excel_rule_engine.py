from __future__ import annotations

import itertools
import json
import re
from collections import Counter
from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Callable

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter
from pandas.testing import assert_frame_equal


SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}
AGGREGATIONS = ("sum", "mean", "min", "max", "count", "median", "first", "last")
FORMULA_OPERATORS = ("+", "-", "*", "/")
SIMPLE_EXCEL_FORMULA = re.compile(
    r"^\s*(?P<left>\$?[A-Z]+\$?\d+|-?\d+(?:\.\d+)?)\s*"
    r"(?P<operator>[+\-*/])\s*"
    r"(?P<right>\$?[A-Z]+\$?\d+|-?\d+(?:\.\d+)?)\s*$",
    re.IGNORECASE,
)
CELL_REFERENCE = re.compile(r"^\$?([A-Z]+)\$?(\d+)$", re.IGNORECASE)


class InferenceError(ValueError):
    """Raised when an example pair cannot be represented by supported rules."""


@dataclass
class ColumnRule:
    source: str
    output: str


@dataclass
class AggregateRule:
    source: str
    output: str
    function: str


@dataclass
class SortRule:
    column: str
    ascending: bool


@dataclass
class FilterRule:
    source: str
    operator: str
    value: Any
    value_type: str


@dataclass
class FormulaRule:
    output: str
    operator: str
    left_kind: str
    left_value: Any
    right_kind: str
    right_value: Any


@dataclass
class SheetRule:
    input_sheet: str
    output_sheet: str
    operation: str
    columns: list[ColumnRule] = field(default_factory=list)
    group_by: list[ColumnRule] = field(default_factory=list)
    aggregations: list[AggregateRule] = field(default_factory=list)
    formulas: list[FormulaRule] = field(default_factory=list)
    filters: list[FilterRule] = field(default_factory=list)
    sort: list[SortRule] = field(default_factory=list)
    output_columns: list[str] = field(default_factory=list)
    style: dict[str, Any] = field(default_factory=dict)


@dataclass
class TransformationTemplate:
    version: int
    sheets: list[SheetRule]
    actions: list[str]

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "TransformationTemplate":
        if data.get("version") != 1:
            raise ValueError("Unsupported template version.")
        sheets = []
        for raw_sheet in data.get("sheets", []):
            sheets.append(
                SheetRule(
                    input_sheet=raw_sheet["input_sheet"],
                    output_sheet=raw_sheet["output_sheet"],
                    operation=raw_sheet["operation"],
                    columns=[ColumnRule(**item) for item in raw_sheet.get("columns", [])],
                    group_by=[ColumnRule(**item) for item in raw_sheet.get("group_by", [])],
                    aggregations=[
                        AggregateRule(**item)
                        for item in raw_sheet.get("aggregations", [])
                    ],
                    formulas=[
                        FormulaRule(**item) for item in raw_sheet.get("formulas", [])
                    ],
                    filters=[
                        FilterRule(**item) for item in raw_sheet.get("filters", [])
                    ],
                    sort=[SortRule(**item) for item in raw_sheet.get("sort", [])],
                    output_columns=raw_sheet.get("output_columns", []),
                    style=raw_sheet.get("style", {}),
                )
            )
        return cls(version=1, sheets=sheets, actions=data.get("actions", []))


def normalize_name(value: object) -> str:
    return " ".join(str(value).strip().lower().replace("_", " ").split())


def validate_excel_path(path: str | Path, label: str) -> Path:
    resolved = Path(path).expanduser().resolve()
    if not resolved.exists():
        raise FileNotFoundError(f"{label} file not found: {resolved}")
    if resolved.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ValueError(f"{label} must be an .xlsx or .xlsm workbook.")
    return resolved


def clean_frame(frame: pd.DataFrame) -> pd.DataFrame:
    cleaned = frame.dropna(axis=0, how="all").dropna(axis=1, how="all").copy()
    cleaned.columns = [str(column).strip() for column in cleaned.columns]
    if len(set(cleaned.columns)) != len(cleaned.columns):
        raise InferenceError("Duplicate column headers are not supported.")
    return cleaned.reset_index(drop=True)


def unwrap_iferror(formula: str) -> str:
    text = formula.strip()
    if text.upper().startswith("=IFERROR(") and text.endswith(',"")'):
        return text[len("=IFERROR(") : -len(',"")')]
    return text[1:] if text.startswith("=") else text


def excel_operand_value(
    token: str,
    row_number: int,
    row_values: dict[int, object],
) -> float | None:
    reference = CELL_REFERENCE.match(token)
    if reference:
        column_letters, referenced_row = reference.groups()
        if int(referenced_row) != row_number:
            return None
        value = row_values.get(column_index_from_string(column_letters))
        try:
            return float(value)
        except (TypeError, ValueError):
            return None
    try:
        return float(token)
    except ValueError:
        return None


def evaluate_simple_excel_formula(
    formula: str,
    row_number: int,
    row_values: dict[int, object],
) -> object:
    expression = unwrap_iferror(formula)
    match = SIMPLE_EXCEL_FORMULA.match(expression)
    if not match:
        return formula
    left = excel_operand_value(match.group("left"), row_number, row_values)
    right = excel_operand_value(match.group("right"), row_number, row_values)
    if left is None or right is None:
        return formula
    operator = match.group("operator")
    if operator == "+":
        return left + right
    if operator == "-":
        return left - right
    if operator == "*":
        return left * right
    if operator == "/":
        return left / right if right != 0 else None
    return formula


def read_workbook(path: str | Path) -> dict[str, pd.DataFrame]:
    frames: dict[str, pd.DataFrame] = {}
    formula_workbook = load_workbook(path, data_only=False, read_only=True)
    value_workbook = load_workbook(path, data_only=True, read_only=True)
    try:
        for sheet_name in formula_workbook.sheetnames:
            formula_sheet = formula_workbook[sheet_name]
            value_sheet = value_workbook[sheet_name]
            formula_rows = formula_sheet.iter_rows(values_only=True)
            value_rows = value_sheet.iter_rows(values_only=True)
            headers = list(next(formula_rows, ()))
            next(value_rows, ())
            active_columns = [
                index
                for index, header in enumerate(headers, start=1)
                if header is not None
            ]
            if not active_columns:
                continue

            rows: list[list[object]] = []
            for row_number, (formula_row, value_row) in enumerate(
                zip(formula_rows, value_rows),
                start=2,
            ):
                row_values: dict[int, object] = {}
                pending_formulas: dict[int, str] = {}
                for column in active_columns:
                    formula_value = (
                        formula_row[column - 1]
                        if column <= len(formula_row)
                        else None
                    )
                    cached_value = (
                        value_row[column - 1] if column <= len(value_row) else None
                    )
                    if isinstance(formula_value, str) and formula_value.startswith("="):
                        if cached_value is not None:
                            row_values[column] = cached_value
                        else:
                            pending_formulas[column] = formula_value
                    else:
                        row_values[column] = formula_value

                for _ in range(len(pending_formulas) + 1):
                    progress = False
                    for column, formula in list(pending_formulas.items()):
                        evaluated = evaluate_simple_excel_formula(
                            formula,
                            row_number,
                            row_values,
                        )
                        if evaluated != formula:
                            row_values[column] = evaluated
                            del pending_formulas[column]
                            progress = True
                    if not progress:
                        break
                row_values.update(pending_formulas)
                rows.append([row_values.get(column) for column in active_columns])

            frame = pd.DataFrame(
                rows,
                columns=[str(headers[column - 1]).strip() for column in active_columns],
            )
            cleaned = clean_frame(frame)
            if cleaned.shape[1] > 0:
                frames[sheet_name] = cleaned
    finally:
        formula_workbook.close()
        value_workbook.close()
    if not frames:
        raise InferenceError("The workbook contains no tabular sheets with headers.")
    return frames


def canonical_scalar(value: object) -> object:
    if pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.isoformat()
    if hasattr(value, "isoformat") and not isinstance(value, str):
        try:
            return value.isoformat()
        except (TypeError, ValueError):
            pass
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return round(float(value), 10)
    return value


def series_signature(series: pd.Series) -> list[object]:
    return [canonical_scalar(value) for value in series.tolist()]


def series_multiset_signature(series: pd.Series) -> list[str]:
    return sorted(repr(canonical_scalar(value)) for value in series.tolist())


def series_multiset_is_subset(subset: pd.Series, superset: pd.Series) -> bool:
    subset_counts = Counter(
        repr(canonical_scalar(value)) for value in subset.tolist()
    )
    superset_counts = Counter(
        repr(canonical_scalar(value)) for value in superset.tolist()
    )
    return all(
        count <= superset_counts.get(value, 0)
        for value, count in subset_counts.items()
    )


def frames_equal(left: pd.DataFrame, right: pd.DataFrame) -> bool:
    if list(left.columns) != list(right.columns):
        return False
    try:
        assert_frame_equal(
            left.reset_index(drop=True),
            right.reset_index(drop=True),
            check_dtype=False,
            check_names=True,
            check_exact=False,
            rtol=1e-8,
            atol=1e-8,
        )
        return True
    except AssertionError:
        return False


def series_equal(left: pd.Series, right: pd.Series) -> bool:
    return frames_equal(
        pd.DataFrame({"value": left.reset_index(drop=True)}),
        pd.DataFrame({"value": right.reset_index(drop=True)}),
    )


def header_similarity(left: str, right: str) -> float:
    return SequenceMatcher(None, normalize_name(left), normalize_name(right)).ratio()


def choose_unique_candidate(
    output_column: str,
    candidates: list[str],
) -> str | None:
    if not candidates:
        return None
    exact = [
        column
        for column in candidates
        if normalize_name(column) == normalize_name(output_column)
    ]
    if len(exact) == 1:
        return exact[0]
    ranked = sorted(
        ((header_similarity(output_column, column), column) for column in candidates),
        reverse=True,
    )
    if len(ranked) == 1 or ranked[0][0] > ranked[1][0] + 0.15:
        return ranked[0][1]
    return None


def infer_sort(
    unsorted_frame: pd.DataFrame,
    expected_frame: pd.DataFrame,
) -> list[SortRule] | None:
    if frames_equal(unsorted_frame, expected_frame):
        return []

    columns = list(expected_frame.columns)
    max_keys = min(3, len(columns))
    for key_count in range(1, max_keys + 1):
        for sort_columns in itertools.combinations(columns, key_count):
            for directions in itertools.product((True, False), repeat=key_count):
                try:
                    candidate = unsorted_frame.sort_values(
                        list(sort_columns),
                        ascending=list(directions),
                        kind="mergesort",
                        na_position="last",
                    ).reset_index(drop=True)
                except TypeError:
                    continue
                if frames_equal(candidate, expected_frame):
                    return [
                        SortRule(column=column, ascending=ascending)
                        for column, ascending in zip(sort_columns, directions)
                    ]
    return None


def infer_direct_rule(
    input_frame: pd.DataFrame,
    output_frame: pd.DataFrame,
) -> SheetRule | None:
    mappings: list[ColumnRule] = []
    used_sources: set[str] = set()
    unresolved_outputs: list[str] = []
    for output_column in output_frame.columns:
        candidates = []
        output_series = output_frame[output_column]
        for input_column in input_frame.columns:
            if input_column in used_sources:
                continue
            input_series = input_frame[input_column]
            if series_multiset_is_subset(output_series, input_series):
                candidates.append(input_column)
        source_column = choose_unique_candidate(output_column, candidates)
        if source_column is None:
            unresolved_outputs.append(output_column)
            continue
        mappings.append(ColumnRule(source=source_column, output=output_column))
        used_sources.add(source_column)

    formulas: list[FormulaRule] = []
    for output_column in unresolved_outputs:
        formula = infer_formula_rule(input_frame, output_frame[output_column], output_column)
        if formula is None:
            return None
        formulas.append(formula)

    candidate = project_columns(input_frame, mappings, formulas)
    sort_rules = infer_sort(candidate, output_frame)
    filters: list[FilterRule] = []
    if sort_rules is None and len(output_frame) < len(input_frame):
        inferred = infer_single_filter(
            input_frame,
            output_frame,
            mappings,
            formulas,
        )
        if inferred is not None:
            filters, sort_rules = inferred
    if sort_rules is None:
        return None

    return SheetRule(
        input_sheet="",
        output_sheet="",
        operation="project",
        columns=mappings,
        formulas=formulas,
        filters=filters,
        sort=sort_rules,
        output_columns=list(output_frame.columns),
    )


def project_columns(
    input_frame: pd.DataFrame,
    mappings: list[ColumnRule],
    formulas: list[FormulaRule] | None = None,
) -> pd.DataFrame:
    result = input_frame[[rule.source for rule in mappings]].copy()
    result.columns = [rule.output for rule in mappings]
    for formula in formulas or []:
        result[formula.output] = evaluate_formula(input_frame, formula)
    return result


def numeric_series(series: pd.Series) -> pd.Series | None:
    converted = pd.to_numeric(series, errors="coerce")
    if converted.notna().sum() != series.notna().sum():
        return None
    return converted


def formula_operand(frame: pd.DataFrame, kind: str, value: Any) -> pd.Series | float:
    if kind == "column":
        converted = numeric_series(frame[str(value)])
        if converted is None:
            raise TypeError(f"Formula column is not numeric: {value}")
        return converted
    if kind == "constant":
        return float(value)
    raise ValueError(f"Unsupported formula operand kind: {kind}")


def evaluate_formula(frame: pd.DataFrame, rule: FormulaRule) -> pd.Series:
    left = formula_operand(frame, rule.left_kind, rule.left_value)
    right = formula_operand(frame, rule.right_kind, rule.right_value)
    if rule.operator == "+":
        return left + right
    if rule.operator == "-":
        return left - right
    if rule.operator == "*":
        return left * right
    if rule.operator == "/":
        return left / right
    raise ValueError(f"Unsupported formula operator: {rule.operator}")


def formula_matches(candidate: pd.Series, expected: pd.Series) -> bool:
    candidate = candidate.replace([float("inf"), float("-inf")], pd.NA)
    candidate_values = sorted(
        (
            None if pd.isna(value) else round(float(value), 10)
            for value in candidate.tolist()
        ),
        key=repr,
    )
    expected_values = sorted(
        (
            None if pd.isna(value) else round(float(value), 10)
            for value in expected.tolist()
        ),
        key=repr,
    )
    return candidate_values == expected_values


def infer_formula_rule(
    input_frame: pd.DataFrame,
    expected: pd.Series,
    output_column: str,
) -> FormulaRule | None:
    numeric_columns = [
        column
        for column in input_frame.columns
        if numeric_series(input_frame[column]) is not None
    ]
    expected_numeric = numeric_series(expected)
    if expected_numeric is None:
        return None

    candidates: list[FormulaRule] = []
    for left_column, right_column in itertools.permutations(numeric_columns, 2):
        for operator in FORMULA_OPERATORS:
            rule = FormulaRule(
                output=output_column,
                operator=operator,
                left_kind="column",
                left_value=left_column,
                right_kind="column",
                right_value=right_column,
            )
            try:
                if formula_matches(evaluate_formula(input_frame, rule), expected_numeric):
                    candidates.append(rule)
            except (TypeError, ValueError, ZeroDivisionError):
                continue

    expected_values = expected_numeric.dropna().unique().tolist()
    if len(expected_values) <= 200:
        for column in numeric_columns:
            source = numeric_series(input_frame[column])
            source_values = source.dropna()
            if source_values.empty:
                continue
            anchor = source_values.iloc[0]
            constants: set[float] = set()
            for output_value in expected_values:
                constants.update(
                    {
                        float(output_value - anchor),
                        float(anchor - output_value),
                        float(output_value / anchor) if anchor != 0 else 0.0,
                        float(anchor / output_value) if output_value != 0 else 0.0,
                    }
                )
            for constant in constants:
                if not pd.notna(constant):
                    continue
                for operator in FORMULA_OPERATORS:
                    for constant_on_left in (False, True):
                        rule = FormulaRule(
                            output=output_column,
                            operator=operator,
                            left_kind="constant" if constant_on_left else "column",
                            left_value=constant if constant_on_left else column,
                            right_kind="column" if constant_on_left else "constant",
                            right_value=column if constant_on_left else constant,
                        )
                        try:
                            if formula_matches(
                                evaluate_formula(input_frame, rule),
                                expected_numeric,
                            ):
                                candidates.append(rule)
                        except (TypeError, ValueError, ZeroDivisionError):
                            continue

    unique: dict[tuple[Any, ...], FormulaRule] = {}
    for rule in candidates:
        key = (
            rule.operator,
            rule.left_kind,
            rule.left_value,
            rule.right_kind,
            rule.right_value,
        )
        unique[key] = rule
    candidates = list(unique.values())
    if not candidates:
        return None

    candidates.sort(
        key=lambda rule: (
            int(rule.left_kind == "column") + int(rule.right_kind == "column"),
            header_similarity(output_column, str(rule.left_value))
            + header_similarity(output_column, str(rule.right_value)),
            rule.operator in {"+", "*"},
        ),
        reverse=True,
    )
    return candidates[0]


def encode_filter_value(value: object) -> tuple[object, str]:
    if isinstance(value, pd.Timestamp):
        return value.isoformat(), "datetime"
    if isinstance(value, datetime):
        return value.isoformat(), "datetime"
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time()).isoformat(), "datetime"
    if isinstance(value, bool):
        return value, "boolean"
    if isinstance(value, (int, float)):
        return value, "number"
    return str(value), "string"


def decode_filter_value(rule: FilterRule) -> object:
    if rule.value_type == "datetime":
        return pd.Timestamp(rule.value)
    return rule.value


def apply_filters(frame: pd.DataFrame, filters: list[FilterRule]) -> pd.DataFrame:
    result = frame
    for rule in filters:
        value = decode_filter_value(rule)
        series = result[rule.source]
        if rule.value_type == "datetime":
            series = pd.to_datetime(series, errors="coerce")
        if rule.operator == ">=":
            mask = series >= value
        elif rule.operator == "<=":
            mask = series <= value
        elif rule.operator == "==":
            mask = series == value
        else:
            raise ValueError(f"Unsupported filter operator: {rule.operator}")
        result = result[mask.fillna(False)]
    return result.reset_index(drop=True)


def infer_single_filter(
    input_frame: pd.DataFrame,
    output_frame: pd.DataFrame,
    mappings: list[ColumnRule],
    formulas: list[FormulaRule] | None = None,
) -> tuple[list[FilterRule], list[SortRule]] | None:
    for mapping in mappings:
        expected_series = output_frame[mapping.output].dropna()
        if expected_series.empty:
            continue

        possible_values: list[tuple[str, object]] = []
        converted_numbers = pd.to_numeric(expected_series, errors="coerce")
        is_numeric = converted_numbers.notna().all()
        is_datetime = all(
            isinstance(value, (pd.Timestamp, datetime, date))
            for value in expected_series
        )
        if is_datetime:
            converted_dates = pd.to_datetime(expected_series, errors="coerce")
            possible_values.extend(
                [(">=", converted_dates.min()), ("<=", converted_dates.max())]
            )
        elif is_numeric:
            possible_values.extend(
                [(">=", converted_numbers.min()), ("<=", converted_numbers.max())]
            )
        if expected_series.nunique(dropna=True) == 1:
            possible_values.append(("==", expected_series.iloc[0]))

        for operator, raw_value in possible_values:
            value, value_type = encode_filter_value(raw_value)
            filter_rule = FilterRule(
                source=mapping.source,
                operator=operator,
                value=value,
                value_type=value_type,
            )
            try:
                filtered_input = apply_filters(input_frame, [filter_rule])
            except (KeyError, TypeError, ValueError):
                continue
            if len(filtered_input) != len(output_frame):
                continue
            candidate = project_columns(filtered_input, mappings, formulas)
            sort_rules = infer_sort(candidate, output_frame)
            if sort_rules is not None:
                return [filter_rule], sort_rules
    return None


def possible_group_keys(
    input_frame: pd.DataFrame,
    output_frame: pd.DataFrame,
) -> list[tuple[ColumnRule, ...]]:
    key_candidates: list[ColumnRule] = []
    for output_column in output_frame.columns:
        output_values = {
            repr(canonical_scalar(value))
            for value in output_frame[output_column].dropna().tolist()
        }
        for input_column in input_frame.columns:
            input_values = {
                repr(canonical_scalar(value))
                for value in input_frame[input_column].dropna().tolist()
            }
            if output_values and output_values.issubset(input_values):
                if (
                    normalize_name(output_column) == normalize_name(input_column)
                    or header_similarity(output_column, input_column) >= 0.6
                ):
                    key_candidates.append(
                        ColumnRule(source=input_column, output=output_column)
                    )

    possible: list[tuple[ColumnRule, ...]] = []
    for key_count in range(1, min(3, len(key_candidates) + 1)):
        for rules in itertools.combinations(key_candidates, key_count):
            if len({rule.source for rule in rules}) != len(rules):
                continue
            if len({rule.output for rule in rules}) != len(rules):
                continue
            grouped_count = input_frame.groupby(
                [rule.source for rule in rules],
                dropna=False,
                sort=False,
            ).ngroups
            if grouped_count == len(output_frame):
                possible.append(rules)
    return possible


def aggregate_series(
    input_frame: pd.DataFrame,
    key_sources: list[str],
    source_column: str,
    function: str,
) -> pd.DataFrame:
    grouped = input_frame.groupby(key_sources, dropna=False, sort=False)
    if function == "count":
        result = grouped[source_column].count()
    else:
        result = grouped[source_column].agg(function)
    return result.reset_index(name="__value__")


def align_expected_by_keys(
    output_frame: pd.DataFrame,
    key_rules: tuple[ColumnRule, ...],
) -> pd.DataFrame:
    return output_frame.set_index([rule.output for rule in key_rules], drop=False)


def infer_group_rule(
    input_frame: pd.DataFrame,
    output_frame: pd.DataFrame,
) -> SheetRule | None:
    solutions: list[SheetRule] = []

    for key_rules in possible_group_keys(input_frame, output_frame):
        key_sources = [rule.source for rule in key_rules]
        key_outputs = [rule.output for rule in key_rules]
        expected_indexed = align_expected_by_keys(output_frame, key_rules)
        aggregate_rules: list[AggregateRule] = []
        unresolved_outputs: list[str] = []

        for output_column in output_frame.columns:
            if output_column in key_outputs:
                continue
            matching: list[AggregateRule] = []
            expected_series = expected_indexed[output_column]

            for input_column in input_frame.columns:
                if input_column in key_sources:
                    continue
                for function in AGGREGATIONS:
                    try:
                        aggregate = aggregate_series(
                            input_frame,
                            key_sources,
                            input_column,
                            function,
                        )
                        aggregate = aggregate.rename(
                            columns={
                                rule.source: rule.output
                                for rule in key_rules
                            }
                        )
                        aggregate_indexed = aggregate.set_index(key_outputs)
                        candidate_series = aggregate_indexed.loc[
                            expected_series.index, "__value__"
                        ]
                    except (KeyError, TypeError, ValueError):
                        continue
                    if series_equal(
                        candidate_series.reset_index(drop=True),
                        expected_series.reset_index(drop=True),
                    ):
                        matching.append(
                            AggregateRule(
                                source=input_column,
                                output=output_column,
                                function=function,
                            )
                        )

            if not matching:
                unresolved_outputs.append(output_column)
                continue
            matching.sort(
                key=lambda rule: (
                    header_similarity(output_column, rule.source),
                    rule.function == "sum",
                ),
                reverse=True,
            )
            if (
                len(matching) > 1
                and header_similarity(output_column, matching[0].source)
                == header_similarity(output_column, matching[1].source)
                and matching[0].function == matching[1].function
            ):
                unresolved_outputs.append(output_column)
                continue
            aggregate_rules.append(matching[0])

        rule = SheetRule(
            input_sheet="",
            output_sheet="",
            operation="group",
            group_by=list(key_rules),
            aggregations=aggregate_rules,
            output_columns=key_outputs + [
                item.output for item in aggregate_rules
            ],
        )
        candidate = apply_sheet_rule(input_frame, rule, apply_sort=False)
        formulas: list[FormulaRule] = []
        for output_column in unresolved_outputs:
            formula = infer_formula_rule(
                candidate,
                output_frame[output_column],
                output_column,
            )
            if formula is None:
                formulas = []
                break
            formulas.append(formula)
        if unresolved_outputs and not formulas:
            continue
        rule.formulas = formulas
        rule.output_columns = list(output_frame.columns)
        candidate = apply_sheet_rule(input_frame, rule, apply_sort=False)
        sort_rules = infer_sort(candidate, output_frame)
        if sort_rules is None:
            continue
        rule.sort = sort_rules
        solutions.append(rule)

    if len(solutions) == 1:
        return solutions[0]
    if len(solutions) > 1:
        ranked = sorted(
            solutions,
            key=lambda rule: (
                len(rule.group_by),
                sum(
                    header_similarity(item.output, item.source)
                    for item in rule.group_by + rule.aggregations
                ),
            ),
            reverse=True,
        )
        if len(ranked[0].group_by) > len(ranked[1].group_by):
            return ranked[0]
        if (
            sum(
                header_similarity(item.output, item.source)
                for item in ranked[0].group_by + ranked[0].aggregations
            )
            > sum(
                header_similarity(item.output, item.source)
                for item in ranked[1].group_by + ranked[1].aggregations
            )
            + 0.2
        ):
            return ranked[0]
    return None


def apply_sheet_rule(
    input_frame: pd.DataFrame,
    rule: SheetRule,
    apply_sort: bool = True,
) -> pd.DataFrame:
    required = {
        item.source
        for item in rule.columns + rule.group_by + rule.aggregations + rule.filters
    }
    if rule.operation == "project":
        for formula in rule.formulas:
            if formula.left_kind == "column":
                required.add(str(formula.left_value))
            if formula.right_kind == "column":
                required.add(str(formula.right_value))
    missing = sorted(required - set(input_frame.columns))
    if missing:
        raise ValueError(
            f"Sheet '{rule.input_sheet}' is missing required columns: "
            + ", ".join(missing)
        )

    filtered_frame = apply_filters(input_frame, rule.filters)

    if rule.operation == "project":
        result = project_columns(filtered_frame, rule.columns, rule.formulas)
    elif rule.operation == "group":
        key_sources = [item.source for item in rule.group_by]
        grouped = filtered_frame.groupby(key_sources, dropna=False, sort=False)
        key_frame = grouped.size().reset_index(name="__row_count__").drop(
            columns="__row_count__"
        )
        result = key_frame
        for aggregate in rule.aggregations:
            values = aggregate_series(
                filtered_frame,
                key_sources,
                aggregate.source,
                aggregate.function,
            )
            result = result.merge(values, on=key_sources, how="left")
            result = result.rename(columns={"__value__": aggregate.output})
        result = result.rename(
            columns={item.source: item.output for item in rule.group_by}
        )
        for formula in rule.formulas:
            result[formula.output] = evaluate_formula(result, formula)
    else:
        raise ValueError(f"Unsupported operation: {rule.operation}")

    result = result[rule.output_columns]
    if apply_sort and rule.sort:
        result = result.sort_values(
            [item.column for item in rule.sort],
            ascending=[item.ascending for item in rule.sort],
            kind="mergesort",
            na_position="last",
        )
    return result.reset_index(drop=True)


def color_to_text(color) -> str | None:
    if color is None:
        return None
    if color.type == "rgb" and color.rgb:
        if color.rgb in {"00000000", "000000"}:
            return None
        return color.rgb
    return None


def extract_style(path: Path, sheet_name: str, columns: list[str]) -> dict[str, Any]:
    workbook = load_workbook(path, read_only=False, data_only=False)
    worksheet = workbook[sheet_name]
    header_style: dict[str, Any] = {}
    if worksheet.max_column:
        cell = worksheet.cell(1, 1)
        header_style = {
            "font_name": cell.font.name,
            "font_size": cell.font.sz,
            "bold": cell.font.bold,
            "italic": cell.font.italic,
            "font_color": color_to_text(cell.font.color),
            "fill_color": color_to_text(cell.fill.fgColor),
            "horizontal": cell.alignment.horizontal,
            "vertical": cell.alignment.vertical,
        }
    widths = {}
    number_formats = {}
    for index, column_name in enumerate(columns, start=1):
        letter = worksheet.cell(1, index).column_letter
        widths[column_name] = worksheet.column_dimensions[letter].width
        if worksheet.max_row >= 2:
            number_formats[column_name] = worksheet.cell(2, index).number_format
    style = {
        "header": header_style,
        "column_widths": widths,
        "number_formats": number_formats,
        "freeze_panes": str(worksheet.freeze_panes) if worksheet.freeze_panes else None,
        "auto_filter": bool(worksheet.auto_filter.ref),
    }
    workbook.close()
    return style


def describe_rule(
    rule: SheetRule,
    input_columns: list[str],
) -> list[str]:
    used_sources = {
        item.source
        for item in rule.columns + rule.group_by + rule.aggregations + rule.filters
    }
    if rule.operation == "project":
        for formula in rule.formulas:
            if formula.left_kind == "column":
                used_sources.add(str(formula.left_value))
            if formula.right_kind == "column":
                used_sources.add(str(formula.right_value))
    actions = []
    removed = [column for column in input_columns if column not in used_sources]
    if removed:
        actions.append(
            f"Removed {len(removed)} column(s): {', '.join(removed)}"
        )

    renamed = [
        (item.source, item.output)
        for item in rule.columns + rule.group_by + rule.aggregations
        if item.source != item.output
    ]
    if renamed:
        actions.append(
            "Renamed columns: "
            + ", ".join(f"{source} -> {output}" for source, output in renamed)
        )

    if rule.operation == "group":
        actions.append(
            "Grouped by "
            + ", ".join(item.output for item in rule.group_by)
        )
        for item in rule.aggregations:
            actions.append(
                f"{item.function.title()} {item.source} as {item.output}"
            )

    for item in rule.formulas:
        left = item.left_value
        right = item.right_value
        actions.append(
            f"Calculated {item.output} = {left} {item.operator} {right}"
        )

    for item in rule.filters:
        value = item.value
        if item.value_type == "datetime":
            value = pd.Timestamp(value).date().isoformat()
        actions.append(f"Filtered {item.source} {item.operator} {value}")

    source_order = [
        item.source for item in rule.columns
    ] if rule.operation == "project" else []
    if source_order and source_order != [
        column for column in input_columns if column in used_sources
    ]:
        actions.append("Reordered output columns")

    if rule.sort:
        actions.append(
            "Sorted by "
            + ", ".join(
                f"{item.column} {'ascending' if item.ascending else 'descending'}"
                for item in rule.sort
            )
        )
    if not actions:
        actions.append("Copied data without structural changes")
    return actions


def infer_sheet_rule(
    input_frame: pd.DataFrame,
    output_frame: pd.DataFrame,
) -> SheetRule | None:
    rule = infer_direct_rule(input_frame, output_frame)
    if rule is None:
        rule = infer_group_rule(input_frame, output_frame)
    return rule


def required_columns(rule: SheetRule) -> set[str]:
    required = {
        item.source
        for item in rule.columns + rule.group_by + rule.aggregations + rule.filters
    }
    if rule.operation == "project":
        for formula in rule.formulas:
            if formula.left_kind == "column":
                required.add(str(formula.left_value))
            if formula.right_kind == "column":
                required.add(str(formula.right_value))
    return required


def schema_match_score(frame: pd.DataFrame, rule: SheetRule) -> float:
    required = required_columns(rule)
    if not required:
        return 0.0
    available = set(frame.columns)
    if not required.issubset(available):
        return -1.0
    extra_penalty = len(available - required) * 0.001
    return 1.0 - extra_penalty


def unique_sheet_name(name: str, used_names: set[str]) -> str:
    base = name[:31] or "Output"
    candidate = base
    suffix = 2
    while candidate in used_names:
        suffix_text = f" ({suffix})"
        candidate = f"{base[:31 - len(suffix_text)]}{suffix_text}"
        suffix += 1
    used_names.add(candidate)
    return candidate


def infer_template(
    input_path: str | Path,
    expected_output_path: str | Path,
    progress: Callable[[str], None] | None = None,
) -> TransformationTemplate:
    input_file = validate_excel_path(input_path, "Input")
    output_file = validate_excel_path(expected_output_path, "Expected output")
    input_frames = read_workbook(input_file)
    output_frames = read_workbook(output_file)

    rules: list[SheetRule] = []
    all_actions: list[str] = []
    unused_inputs = set(input_frames)
    for output_sheet, output_frame in output_frames.items():
        if progress:
            progress(f"Analyzing sheet: {output_sheet}")
        candidate_names = sorted(
            unused_inputs,
            key=lambda name: (
                name != output_sheet,
                -header_similarity(name, output_sheet),
                name,
            ),
        )
        matches: list[tuple[str, SheetRule]] = []
        for input_sheet in candidate_names:
            candidate_rule = infer_sheet_rule(input_frames[input_sheet], output_frame)
            if candidate_rule is not None:
                matches.append((input_sheet, candidate_rule))
                if input_sheet == output_sheet:
                    break

        if not matches:
            all_actions.extend(
                [
                    f"[{output_sheet}]",
                    "[WARN] Skipped: no input sheet could safely reproduce this output.",
                ]
            )
            continue

        input_sheet, rule = matches[0]
        input_frame = input_frames[input_sheet]

        rule.input_sheet = input_sheet
        rule.output_sheet = output_sheet
        rule.style = extract_style(output_file, output_sheet, rule.output_columns)
        reproduced = apply_sheet_rule(input_frame, rule)
        if not frames_equal(reproduced, output_frame):
            raise InferenceError(
                f"Detected rules for sheet '{output_sheet}' did not reproduce the "
                "expected output exactly. No template was created."
            )

        sheet_actions = describe_rule(rule, list(input_frame.columns))
        all_actions.append(f"[{output_sheet}]")
        if input_sheet != output_sheet:
            all_actions.append(
                f"[WARN] Matched renamed sheet: {input_sheet} -> {output_sheet}"
            )
        all_actions.extend(f"[OK] {action}" for action in sheet_actions)
        rules.append(rule)
        unused_inputs.discard(input_sheet)

    for input_sheet in sorted(unused_inputs):
        all_actions.extend(
            [
                f"[{input_sheet}]",
                "[WARN] Ignored: no expected-output sheet could be matched safely.",
            ]
        )
    if not rules:
        raise InferenceError(
            "No compatible worksheet pair could be inferred. Check that at least "
            "one input sheet can reproduce one expected-output sheet."
        )
    return TransformationTemplate(version=1, sheets=rules, actions=all_actions)


def save_template(template: TransformationTemplate, path: str | Path) -> Path:
    output = Path(path).expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(
        json.dumps(template.to_dict(), indent=2, ensure_ascii=True),
        encoding="utf-8",
    )
    return output


def load_template(path: str | Path) -> TransformationTemplate:
    template_path = Path(path).expanduser().resolve()
    if not template_path.exists():
        raise FileNotFoundError(f"Template not found: {template_path}")
    return TransformationTemplate.from_dict(
        json.loads(template_path.read_text(encoding="utf-8"))
    )


def apply_style(worksheet, rule: SheetRule) -> None:
    style = rule.style
    header = style.get("header", {})
    fill_color = header.get("fill_color")
    font_color = header.get("font_color")
    for cell in worksheet[1]:
        cell.font = Font(
            name=header.get("font_name") or "Calibri",
            size=header.get("font_size") or 11,
            bold=bool(header.get("bold")),
            italic=bool(header.get("italic")),
            color=font_color,
        )
        if fill_color:
            cell.fill = PatternFill("solid", fgColor=fill_color)
        cell.alignment = Alignment(
            horizontal=header.get("horizontal"),
            vertical=header.get("vertical"),
        )

    widths = style.get("column_widths", {})
    formats = style.get("number_formats", {})
    for index, column_name in enumerate(rule.output_columns, start=1):
        letter = worksheet.cell(1, index).column_letter
        if widths.get(column_name):
            worksheet.column_dimensions[letter].width = widths[column_name]
        number_format = formats.get(column_name)
        if number_format:
            for row_number in range(2, worksheet.max_row + 1):
                worksheet.cell(row_number, index).number_format = number_format

    if style.get("freeze_panes"):
        worksheet.freeze_panes = style["freeze_panes"]
    if style.get("auto_filter") and worksheet.max_row:
        worksheet.auto_filter.ref = worksheet.dimensions


def excel_formula_operand(
    rule: FormulaRule,
    kind: str,
    value: Any,
    row_number: int,
    source_to_output: dict[str, str],
    output_positions: dict[str, int],
) -> str | None:
    if kind == "constant":
        return format(float(value), ".15g")
    output_name = source_to_output.get(str(value))
    if output_name is None or output_name not in output_positions:
        return None
    column_letter = get_column_letter(output_positions[output_name])
    return f"{column_letter}{row_number}"


def write_excel_formulas(worksheet, rule: SheetRule) -> None:
    if not rule.formulas:
        return
    source_to_output = {item.source: item.output for item in rule.columns}
    if rule.operation == "group":
        source_to_output.update(
            {item.output: item.output for item in rule.group_by + rule.aggregations}
        )
    output_positions = {
        column_name: index
        for index, column_name in enumerate(rule.output_columns, start=1)
    }
    for formula in rule.formulas:
        output_index = output_positions.get(formula.output)
        if output_index is None:
            continue
        for row_number in range(2, worksheet.max_row + 1):
            left = excel_formula_operand(
                formula,
                formula.left_kind,
                formula.left_value,
                row_number,
                source_to_output,
                output_positions,
            )
            right = excel_formula_operand(
                formula,
                formula.right_kind,
                formula.right_value,
                row_number,
                source_to_output,
                output_positions,
            )
            if left is None or right is None:
                break
            worksheet.cell(row_number, output_index).value = (
                f"=IFERROR({left}{formula.operator}{right},\"\")"
            )


def apply_template(
    template: TransformationTemplate,
    input_path: str | Path,
    output_path: str | Path,
    progress: Callable[[str], None] | None = None,
) -> Path:
    input_file = validate_excel_path(input_path, "Input")
    output_file = Path(output_path).expanduser().resolve()
    if input_file == output_file:
        raise ValueError("Choose a different output path to preserve the input file.")
    input_frames = read_workbook(input_file)
    generated: dict[str, pd.DataFrame] = {}
    applied_rules: list[tuple[SheetRule, str]] = []
    warnings: list[str] = []
    unused_inputs = set(input_frames)
    used_output_names: set[str] = set()

    for rule in template.sheets:
        candidates = []
        for sheet_name in unused_inputs:
            score = schema_match_score(input_frames[sheet_name], rule)
            if score >= 0:
                candidates.append((sheet_name == rule.input_sheet, score, sheet_name))
        if not candidates:
            warnings.append(
                f"Skipped '{rule.output_sheet}': no sheet contains required columns "
                f"{', '.join(sorted(required_columns(rule)))}."
            )
            continue
        candidates.sort(reverse=True)
        input_sheet = candidates[0][2]
        if progress:
            progress(f"Applying template to: {input_sheet}")
        try:
            result = apply_sheet_rule(input_frames[input_sheet], rule)
        except Exception as exc:
            warnings.append(
                f"Skipped '{rule.output_sheet}' using '{input_sheet}': {exc}"
            )
            unused_inputs.discard(input_sheet)
            continue
        output_sheet = unique_sheet_name(rule.output_sheet, used_output_names)
        generated[output_sheet] = result
        applied_rules.append((rule, output_sheet))
        unused_inputs.discard(input_sheet)
        if input_sheet != rule.input_sheet:
            warnings.append(
                f"Used '{input_sheet}' for expected sheet '{rule.input_sheet}' "
                "because the original sheet was missing."
            )

    for sheet_name in sorted(unused_inputs):
        warnings.append(
            f"Ignored input sheet '{sheet_name}': no remaining template rule matched."
        )
    if not generated:
        raise ValueError(
            "No worksheet could be transformed. The workbook does not contain the "
            "minimum columns required by any template rule."
        )

    output_file.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for sheet_name, frame in generated.items():
            frame.to_excel(writer, sheet_name=sheet_name, index=False)
        if warnings:
            pd.DataFrame({"Warning": warnings}).to_excel(
                writer,
                sheet_name="Transformation Warnings",
                index=False,
            )

    workbook = load_workbook(output_file)
    for rule, output_sheet in applied_rules:
        worksheet = workbook[output_sheet]
        apply_style(worksheet, rule)
        write_excel_formulas(worksheet, rule)
    workbook.save(output_file)
    workbook.close()
    return output_file
