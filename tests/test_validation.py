"""검수 규칙과 실제 Excel 읽기의 경계 조건을 합성 자료로 검증한다."""

import pandas as pd
import pytest

from validation.validate_excel import format_report, main, normalize_code, validate_frames, validate_workbooks


@pytest.mark.parametrize("value, expected", [(1, 1), (1.0, 1), ("1.0", 1), ("001", 1), (0, 0), (None, None), (float("nan"), None), (pd.NA, None), ("", None)])
def test_code_normalization(value, expected):
    """동일 정수 표기와 빈 셀을 일관되게 판정한다."""
    assert normalize_code(value) == expected


@pytest.mark.parametrize("value", [1.5, "1.2", "문자", -1, True, "1e2"])
def test_invalid_codes_are_not_truncated(value):
    """소수와 비숫자를 정수로 잘라서 받아들이지 않는다."""
    with pytest.raises(ValueError):
        normalize_code(value)


def test_valid_frames_preserve_inputs(response_frame, codeframe):
    """반복·단일 문항 열을 검수하고 입력과 수동 확인 범주를 보존한다."""
    before_data, before_codes = response_frame.copy(deep=True), codeframe.copy(deep=True)
    report = validate_frames(response_frame, codeframe)
    assert report.findings == []
    assert report.no_response_nets == ["기타"]
    pd.testing.assert_frame_equal(response_frame, before_data)
    pd.testing.assert_frame_equal(codeframe, before_codes)


def test_workbooks_preserve_bytes_and_duplicate_headers(workbooks):
    """Excel 중복 제목을 손실 없이 읽고 두 원본 파일의 바이트를 보존한다."""
    before = [path.read_bytes() for path in workbooks]
    assert validate_workbooks(*workbooks).findings == []
    assert [path.read_bytes() for path in workbooks] == before


def test_real_question_suffix_is_not_removed(response_frame, codeframe, tmp_path):
    """실제 문항명 Q1.1을 pandas의 중복 열 접미사로 오인하지 않는다."""
    response_frame.columns = ["합성제품", "Q1.1", "Q1.1", "Q2"]
    codeframe.loc[codeframe["문항"] == "Q1", "문항"] = "Q1.1"
    data_path, codes_path = tmp_path / "data.xlsx", tmp_path / "codes.xlsx"
    pd.DataFrame([list(response_frame.columns), *response_frame.values.tolist()]).to_excel(data_path, index=False, header=False)
    codeframe.to_excel(codes_path, index=False)
    assert validate_workbooks(data_path, codes_path).findings == []


def test_invalid_value_does_not_stop_later_checks(response_frame, codeframe):
    """잘못된 Sub-Net 뒤의 미등록 코드와 Net 누락까지 계속 보고한다."""
    response_frame.iloc[1, 2] = "잘못된 합성 값"
    response_frame.iloc[2, 2] = 88
    codeframe.loc[0, "Net"] = None
    report = validate_frames(response_frame, codeframe)
    findings = {finding.code: finding for finding in report.findings}
    assert {"INVALID_SUBNET", "UNKNOWN_CODE", "MISSING_NET"} <= findings.keys()
    assert (findings["INVALID_SUBNET"].row, findings["INVALID_SUBNET"].column) == (3, 3)
    assert findings["UNKNOWN_CODE"].row == 4
    assert "잘못된 합성 값" not in format_report(report)


def test_duplicate_codes_use_normalized_numbers(response_frame, codeframe):
    """같은 문항의 1과 1.0을 중복 코드로 보고한다."""
    extra = codeframe.iloc[[0]].copy()
    extra["코드"] = 1
    report = validate_frames(response_frame, pd.concat([codeframe, extra], ignore_index=True))
    assert [finding.code for finding in report.findings] == ["DUPLICATE_CODE"]


def test_question_differences_are_reported_both_ways(response_frame, codeframe):
    """데이터 전용 문항과 코드프레임 전용 문항을 각각 보고한다."""
    codeframe.loc[codeframe["문항"] == "Q2", "문항"] = "Q3"
    codes = {finding.code for finding in validate_frames(response_frame, codeframe).findings}
    assert {"QUESTION_NOT_IN_CODES", "QUESTION_NOT_IN_DATA"} <= codes


def test_product_mismatch_and_missing_columns(response_frame, codeframe):
    """제품 불일치나 필수 열 누락을 후속 예외 없이 보고한다."""
    codeframe["제품"] = "다른합성제품"
    assert validate_frames(response_frame, codeframe).findings[0].code == "PRODUCT_MISMATCH"
    assert validate_frames(response_frame, codeframe.drop(columns="코드")).findings[0].code == "CODEFRAME_COLUMNS"


def test_missing_subnet_and_codeframe_values(response_frame, codeframe):
    """세부 항목 누락과 코드프레임 결측값을 명시적으로 보고한다."""
    response_frame.iloc[0, 3] = "text"
    codeframe.loc[0, "코드"] = None
    codeframe.loc[1, "문항"] = None
    codes = {finding.code for finding in validate_frames(response_frame, codeframe).findings}
    assert {"SUBNET_COLUMN_MISSING", "CODEFRAME_CODE", "CODEFRAME_QUESTION"} <= codes


def test_no_response_categories_are_manual_review(response_frame, codeframe):
    """무응답 범주의 Net을 자동 재분류하거나 오류로 확정하지 않는다."""
    codeframe.loc[2, "Net"] = "별도분류"
    report = validate_frames(response_frame, codeframe)
    assert report.findings == []
    assert report.no_response_nets == ["별도분류", "기타"]
    assert codeframe.loc[2, "Net"] == "별도분류"


def test_sheet_count_warnings_keep_first_sheet_checks(workbooks):
    """시트 개수가 다르면 알리면서 기존처럼 첫 시트 검수를 계속한다."""
    for path in workbooks:
        with pd.ExcelWriter(path, engine="openpyxl", mode="a") as writer:
            pd.DataFrame({"합성": [1]}).to_excel(writer, sheet_name="추가1", index=False)
            pd.DataFrame({"합성": [2]}).to_excel(writer, sheet_name="추가2", index=False)
    assert {finding.code for finding in validate_workbooks(*workbooks).findings} == {"DATA_SHEETS", "CODEFRAME_SHEETS"}


def test_empty_inputs_and_cli_exit_codes(tmp_path, workbooks, response_frame, codeframe, capsys):
    """빈 입력과 명령 실행의 성공·검수 문제·실행 실패를 구분한다."""
    assert validate_frames(pd.DataFrame(), codeframe).findings[0].code == "DATA_LAYOUT"
    assert validate_frames(response_frame.iloc[:1], codeframe).findings[0].code == "NO_RESPONSES"
    assert main([str(path) for path in workbooks]) == 0
    codeframe["제품"] = "다른제품"
    codeframe.to_excel(workbooks[1], index=False)
    assert main([str(path) for path in workbooks]) == 1
    assert main([str(tmp_path / "missing.xlsx"), str(workbooks[1])]) == 2
    assert "합성 응답" not in capsys.readouterr().out
