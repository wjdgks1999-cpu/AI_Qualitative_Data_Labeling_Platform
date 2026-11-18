"""응답 파일과 코드프레임을 읽기 전용으로 검수한다."""

from __future__ import annotations

import argparse
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
import re

import pandas as pd

REQUIRED_COLUMNS = ("제품", "문항", "코드", "Net", "Sub-Net")
NO_RESPONSE_LABELS = ("없다", "없음", "모름/무응답", "모름", "무응답")


@dataclass(frozen=True)
class Finding:
    """검수 항목과 Excel 기준 위치를 담으며 응답 원문은 보관하지 않는다."""

    code: str
    message: str
    row: int | None = None
    column: int | None = None
    question: str | None = None


@dataclass
class ValidationReport:
    """검수 문제와 사람이 확인할 무응답 범주의 Net 목록을 담는다."""

    findings: list[Finding] = field(default_factory=list)
    no_response_nets: list[str | None] = field(default_factory=list)


def is_missing(value: object) -> bool:
    """결측값과 공백 문자열을 비어 있는 셀로 판정한다."""
    return bool(pd.isna(value)) or (isinstance(value, str) and not value.strip())


def normalize_code(value: object) -> int | None:
    """비음수 정수 코드를 통일하며 1, 1.0, '1.0'은 같은 코드로 취급한다.

    비어 있는 셀은 None을 반환하고 소수·음수·비숫자는 거부한다.
    """
    if is_missing(value):
        return None
    text = str(value).strip()
    if not re.fullmatch(r"[0-9]+(?:\.0+)?", text):
        raise ValueError("코드는 비음수 정수여야 합니다.")
    return int(Decimal(text))


def validate_frames(data: pd.DataFrame, codes: pd.DataFrame) -> ValidationReport:
    """제품·문항별 코드와 Net을 비교하고 입력 DataFrame은 변경하지 않는다.

    data의 열 제목은 제품·문항명, 첫 행은 subnet 등의 세부 항목명이다.
    동일한 문항의 여러 열은 중복 열 제목을 그대로 유지해야 한다.
    """
    report = ValidationReport()
    missing = [name for name in REQUIRED_COLUMNS if name not in codes.columns]
    if missing:
        report.findings.append(Finding("CODEFRAME_COLUMNS", "코드프레임 필수 열 누락: " + ", ".join(missing)))
        return report
    if data.empty or len(data.columns) < 2:
        report.findings.append(Finding("DATA_LAYOUT", "제품·문항 열과 세부 항목 행이 필요합니다."))
        return report
    if any(not isinstance(name, str) or not name.strip() for name in data.columns):
        report.findings.append(Finding("DATA_HEADERS", "제품·문항 열 제목은 비어 있지 않은 문자열이어야 합니다."))
        return report
    if len(data) < 2:
        report.findings.append(Finding("NO_RESPONSES", "세부 항목 행 다음에 응답 행이 없습니다."))

    product = data.columns[0]
    selected = codes.loc[codes["제품"] == product].copy()
    if selected.empty:
        report.findings.append(Finding("PRODUCT_MISMATCH", "데이터의 제품명과 일치하는 코드프레임이 없습니다."))
        return report

    questions = list(dict.fromkeys(data.columns[1:]))
    code_questions = set(selected["문항"].dropna())
    for question in questions:
        if question not in code_questions:
            report.findings.append(Finding("QUESTION_NOT_IN_CODES", "코드프레임에 문항이 없습니다.", question=question))
    for question in sorted(code_questions - set(questions), key=str):
        report.findings.append(Finding("QUESTION_NOT_IN_DATA", "응답 데이터에 문항이 없습니다.", question=str(question)))

    allowed: dict[str, set[int]] = {}
    seen: set[tuple[str, int]] = set()
    for position in range(len(codes)):
        record = codes.iloc[position]
        if is_missing(record["제품"]) or record["제품"] != product:
            continue
        question = record["문항"]
        excel_row = position + 2
        if is_missing(question):
            report.findings.append(Finding("CODEFRAME_QUESTION", "코드프레임의 문항이 비어 있습니다.", row=excel_row))
        else:
            try:
                number = normalize_code(record["코드"])
            except ValueError:
                number = None
            if number is None:
                report.findings.append(Finding("CODEFRAME_CODE", "코드프레임에 유효한 정수 코드가 필요합니다.", row=excel_row, question=str(question)))
            else:
                key = (question, number)
                if key in seen:
                    report.findings.append(Finding("DUPLICATE_CODE", "동일 문항에서 코드가 반복됩니다.", row=excel_row, question=str(question)))
                seen.add(key)
                allowed.setdefault(question, set()).add(number)
        if is_missing(record["Net"]):
            report.findings.append(Finding("MISSING_NET", "상위 Net이 없는 Sub-Net이 있습니다.", row=excel_row, question=None if is_missing(question) else str(question)))
        if isinstance(record["Sub-Net"], str) and record["Sub-Net"] in NO_RESPONSE_LABELS:
            net = None if is_missing(record["Net"]) else str(record["Net"])
            if net not in report.no_response_nets:
                report.no_response_nets.append(net)

    for question in questions:
        positions = [index for index, name in enumerate(data.columns) if index > 0 and name == question]
        subnet_positions = [index for index in positions if "subnet" in str(data.iloc[0, index])]
        if not subnet_positions:
            report.findings.append(Finding("SUBNET_COLUMN_MISSING", "문항에 subnet 세부 항목 열이 없습니다.", question=question))
        for column in subnet_positions:
            for row in range(1, len(data)):
                try:
                    number = normalize_code(data.iloc[row, column])
                except ValueError:
                    report.findings.append(Finding("INVALID_SUBNET", "Sub-Net 값이 정수 코드가 아닙니다.", row=row + 2, column=column + 1, question=question))
                    continue
                if number is not None and question in code_questions and number not in allowed.get(question, set()):
                    report.findings.append(Finding("UNKNOWN_CODE", "코드프레임에 없는 Sub-Net 코드입니다.", row=row + 2, column=column + 1, question=question))
    return report


def validate_workbooks(data_path: str | Path, codes_path: str | Path) -> ValidationReport:
    """각 Excel의 첫 시트를 읽고 시트 개수와 내용 검수를 합친다.

    응답 파일의 원래 중복 열 제목을 보존하기 위해 header=None으로 읽는다.
    파일이나 시트의 내용을 저장·수정하지 않는다.
    """
    sheet_findings = []
    with pd.ExcelFile(data_path) as workbook:
        if len(workbook.sheet_names) != 1:
            sheet_findings.append(Finding("DATA_SHEETS", "응답 파일은 시트 1개를 권장합니다. 첫 시트만 검수했습니다."))
        raw = pd.read_excel(workbook, sheet_name=0, header=None)
    with pd.ExcelFile(codes_path) as workbook:
        if len(workbook.sheet_names) > 2:
            sheet_findings.append(Finding("CODEFRAME_SHEETS", "코드프레임의 시트가 2개를 초과합니다. 첫 시트만 검수했습니다."))
        codes = pd.read_excel(workbook, sheet_name=0)
    if raw.empty:
        report = ValidationReport([Finding("DATA_LAYOUT", "응답 파일이 비어 있습니다.")])
    else:
        data = raw.iloc[1:].reset_index(drop=True).copy()
        data.columns = raw.iloc[0].tolist()
        report = validate_frames(data, codes)
    report.findings[:0] = sheet_findings
    return report


def format_report(report: ValidationReport) -> str:
    """응답 원문 없이 검수 항목과 Excel 행·열 위치를 표시한다."""
    lines = [f"검수 항목 {len(report.findings)}건"]
    for finding in report.findings:
        location = []
        if finding.question is not None:
            location.append(f"문항={finding.question}")
        if finding.row is not None:
            location.append(f"행={finding.row}")
        if finding.column is not None:
            location.append(f"열={finding.column}")
        lines.append(f"[{finding.code}] {finding.message} {' '.join(location)}".rstrip())
    if report.no_response_nets:
        nets = ", ".join("미지정" if value is None else value for value in report.no_response_nets)
        lines.append(f"수동 확인: 없음·모름·무응답 범주의 Net = {nets}")
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    """두 Excel 경로를 받아 검수하며 문제 발견 시 종료 코드 1을 반환한다."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("data", type=Path, help="응답 데이터 Excel")
    parser.add_argument("codeframe", type=Path, help="코드프레임 Excel")
    args = parser.parse_args(argv)
    try:
        report = validate_workbooks(args.data, args.codeframe)
    except Exception as error:
        print(f"Excel 검수 실행 실패: {type(error).__name__}. 파일 형식과 읽기 권한을 확인하세요.")
        return 2
    print(format_report(report))
    return 1 if report.findings else 0


if __name__ == "__main__":
    raise SystemExit(main())
