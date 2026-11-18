"""업무 자료와 무관한 합성 데이터 및 임시 파일만 제공한다."""

import pandas as pd
import pytest


@pytest.fixture
def response_frame():
    """반복 문항 열과 단일 Sub-Net 열을 포함한 합성 응답을 만든다."""
    return pd.DataFrame(
        [["id", "text", "subnet", "subnet"], [1, "합성 응답 가", 1.0, 3], [2, "합성 응답 나", 2, None]],
        columns=["합성제품", "Q1", "Q1", "Q2"],
    )


@pytest.fixture
def codeframe():
    """두 문항의 정상 코드와 수동 확인용 무응답 범주를 만든다."""
    return pd.DataFrame(
        [["합성제품", "Q1", "1.0", "특성", "맛"], ["합성제품", "Q1", 2, "특성", "향"],
         ["합성제품", "Q1", 9, "기타", "모름"], ["합성제품", "Q2", 3, "기타", "없음"]],
        columns=["제품", "문항", "코드", "Net", "Sub-Net"],
    )


@pytest.fixture
def workbooks(tmp_path, response_frame, codeframe):
    """임시 폴더에 원래 중복 헤더 형식을 가진 합성 Excel 두 개를 저장한다."""
    data_path = tmp_path / "2. synthetic_data.xlsx"
    codes_path = tmp_path / "3. synthetic_codes.xlsx"
    raw = pd.DataFrame([list(response_frame.columns), *response_frame.values.tolist()])
    raw.to_excel(data_path, header=False, index=False)
    codeframe.to_excel(codes_path, index=False)
    return data_path, codes_path


@pytest.fixture(autouse=True)
def isolate_plot_cache(monkeypatch, tmp_path):
    """그림 검증의 캐시와 백엔드를 임시 폴더 및 비대화형 모드로 제한한다."""
    monkeypatch.setenv("MPLCONFIGDIR", str(tmp_path / "matplotlib"))
    monkeypatch.setenv("MPLBACKEND", "Agg")
