"""DB·NLP 대체 객체로 입출력 계약과 실패 시 부작용을 검증한다."""

from types import SimpleNamespace
from unittest.mock import Mock
import sys

import pandas as pd
import pytest

from preprocessing.database import fetch_responses, read_db_config
from preprocessing import preprocess_text as module


def identity(text):
    """외부 호출 없이 텍스트를 그대로 돌려주는 테스트용 처리기다."""
    return text


def test_cleaning_returns_modified_text_and_removes_complete_patterns():
    """정제 결과 반영과 구두점 제거 전에 HTML·이메일·URL을 처리하는지 확인한다."""
    text = '<b>합성 문장</b> sample@example.test https://example.test/a?q=1 ㄱㄴ  ²√'
    assert module.clean_text(text) == "합성 문장 2 sqrt"


def test_processing_order_columns_and_index():
    """정제된 값이 단계별로 전달되고 네 출력 열과 원래 인덱스를 보존한다."""
    spacing = Mock(side_effect=lambda text: text + " 띄어쓰기")
    spellcheck = Mock(side_effect=lambda text: text + " 맞춤법")
    tokenize = Mock(return_value="합성 명사")
    source = pd.Series(["합성!", "문장?"], index=[10, 20])
    result = module.preprocess_responses(source, spacing=spacing, spellcheck=spellcheck, tokenize=tokenize)
    assert list(result.columns) == ["text", "corrected_text", "spell_check_text", "tokenized_text"]
    assert result.index.tolist() == [10, 20]
    assert result["text"].tolist() == ["합성", "문장"]
    assert spacing.call_args_list[0].args == ("합성",)
    assert spellcheck.call_args_list[0].args == ("합성 띄어쓰기",)
    assert tokenize.call_args_list[0].args == ("합성 띄어쓰기 맞춤법",)
    assert source.tolist() == ["합성!", "문장?"]


@pytest.mark.parametrize("responses", [[], [None], ["정상 합성", 1], [float("nan")]])
def test_bad_input_is_rejected_before_processing(responses):
    """입력 오류가 있으면 어떤 외부 처리기도 호출하지 않는다."""
    processor = Mock()
    with pytest.raises(ValueError):
        module.preprocess_responses(responses, spacing=processor, spellcheck=processor, tokenize=processor)
    processor.assert_not_called()


def test_bad_processor_result_is_not_silently_kept():
    """NLP의 비정상 반환값이 다음 단계로 전달되지 않도록 한다."""
    later = Mock()
    with pytest.raises(ValueError, match="corrected_text"):
        module.preprocess_responses(["합성"], spacing=Mock(return_value=None), spellcheck=later, tokenize=later)
    later.assert_not_called()


def test_nlp_construction_is_opt_in_and_reused(monkeypatch):
    """실제 라이브러리 대신 대체 객체로 동의 검사와 객체 재사용을 확인한다."""
    spacing_factory = Mock(return_value=identity)
    okt = SimpleNamespace(nouns=Mock(return_value=["합성", "명사"]))
    okt_factory = Mock(return_value=okt)
    checker = SimpleNamespace(check=Mock(return_value=SimpleNamespace(result=True, checked="합성")))
    monkeypatch.setitem(sys.modules, "pykospacing", SimpleNamespace(Spacing=spacing_factory))
    monkeypatch.setitem(sys.modules, "hanspell", SimpleNamespace(spell_checker=checker))
    monkeypatch.setitem(sys.modules, "konlpy", SimpleNamespace())
    monkeypatch.setitem(sys.modules, "konlpy.tag", SimpleNamespace(Okt=okt_factory))
    with pytest.raises(ValueError):
        module.create_nlp_processors()
    spacing_factory.assert_not_called()
    spacing, spellcheck, tokenize = module.create_nlp_processors(allow_external_spellcheck=True)
    result = module.preprocess_responses(["합성 가", "합성 나"], spacing=spacing, spellcheck=spellcheck, tokenize=tokenize)
    assert result["tokenized_text"].tolist() == ["합성 명사", "합성 명사"]
    spacing_factory.assert_called_once()
    okt_factory.assert_called_once()
    checker.check.return_value = SimpleNamespace(result=False)
    with pytest.raises(ValueError):
        spellcheck("합성")


def test_db_config_does_not_reveal_values():
    """환경변수 누락·포트 오류가 설정 값을 노출하지 않는다."""
    environment = {"MYSQL_HOST": "synthetic.invalid", "MYSQL_USER": "synthetic", "MYSQL_PASSWORD": "synthetic-only", "MYSQL_DATABASE": "synthetic"}
    assert read_db_config(environment)["port"] == 3306
    environment["MYSQL_PORT"] = "invalid-port-fixture"
    with pytest.raises(ValueError) as error:
        read_db_config(environment)
    assert "invalid-port-fixture" not in str(error.value)
    with pytest.raises(ValueError, match="MYSQL_HOST"):
        read_db_config({})


@pytest.mark.parametrize("failure", [None, "cursor", "execute", "fetchall", "close"])
def test_db_connection_is_read_only_and_always_closed(failure):
    """조회 성공과 각 실패 지점에서 연결을 닫고 쓰기 명령을 호출하지 않는다."""
    cursor = Mock()
    cursor.fetchall.return_value = [("합성 응답",)]
    connection = Mock()
    connection.cursor.return_value = cursor
    if failure == "cursor":
        connection.cursor.side_effect = RuntimeError("합성 실패")
    elif failure:
        getattr(cursor, failure).side_effect = RuntimeError("합성 실패")
    connect = Mock(return_value=connection)
    if failure:
        with pytest.raises(RuntimeError):
            fetch_responses(config={}, connect=connect)
    else:
        assert fetch_responses(config={}, connect=connect) == ["합성 응답"]
    connection.close.assert_called_once()
    connection.commit.assert_not_called()
    if failure != "cursor":
        cursor.close.assert_called_once()
        assert cursor.execute.call_args.args == ("SELECT `response` FROM `food_project`",)


def test_bad_table_is_rejected_before_connection():
    """테이블 식별자로 SQL 구문이 유입되는 경우 연결 전에 차단한다."""
    connect = Mock()
    with pytest.raises(ValueError):
        fetch_responses("table; DELETE FROM table", config={}, connect=connect)
    connect.assert_not_called()


def test_excel_output_preserves_existing_file(tmp_path):
    """기존 출력은 기본적으로 보존하고 명시적 덮어쓰기만 허용한다."""
    path = tmp_path / "output.xlsx"
    data = module.preprocess_responses(["합성!"], spacing=identity, spellcheck=identity, tokenize=identity)
    module.save_preprocessed_data(data, path)
    before = path.read_bytes()
    with pytest.raises(FileExistsError):
        module.save_preprocessed_data(data, path)
    assert path.read_bytes() == before
    data.loc[0, "tokenized_text"] = "변경 합성"
    module.save_preprocessed_data(data, path, overwrite=True)
    pd.testing.assert_frame_equal(pd.read_excel(path), data)


def test_serialization_failure_preserves_existing_output(tmp_path, monkeypatch):
    """Excel 변환 자체가 실패하면 덮어쓰기 허용 상태에서도 기존 파일을 보존한다."""
    path = tmp_path / "output.xlsx"
    path.write_bytes(b"synthetic-existing-content")
    monkeypatch.setattr(pd.DataFrame, "to_excel", Mock(side_effect=RuntimeError("합성 실패")))
    with pytest.raises(RuntimeError):
        module.save_preprocessed_data(pd.DataFrame(), path, overwrite=True)
    assert path.read_bytes() == b"synthetic-existing-content"


def test_cli_rejects_side_effects_without_consent_or_with_existing_file(tmp_path, monkeypatch):
    """외부 전송 미허용이나 출력 충돌이면 NLP와 DB를 시작하지 않는다."""
    fetch = Mock()
    create = Mock()
    monkeypatch.setattr(module, "fetch_responses", fetch)
    monkeypatch.setattr(module, "create_nlp_processors", create)
    with pytest.raises(SystemExit) as error:
        module.main([])
    assert error.value.code == 2
    path = tmp_path / "exists.xlsx"
    path.write_bytes(b"synthetic")
    assert module.main(["--allow-external-spellcheck", "--output", str(path)]) == 2
    fetch.assert_not_called()
    create.assert_not_called()


def test_cli_failure_leaves_no_output_or_private_error(tmp_path, monkeypatch, capsys):
    """서비스 실패 시 부분 산출물과 원문이 들어간 예외 메시지를 남기지 않는다."""
    monkeypatch.setattr(module, "fetch_responses", Mock(return_value=["합성"] ))
    monkeypatch.setattr(module, "create_nlp_processors", Mock(return_value=(identity, Mock(side_effect=RuntimeError("private-error-fixture")), identity)))
    path = tmp_path / "not_created.xlsx"
    assert module.main(["--allow-external-spellcheck", "--output", str(path)]) == 2
    assert not path.exists()
    assert "private-error-fixture" not in capsys.readouterr().out
