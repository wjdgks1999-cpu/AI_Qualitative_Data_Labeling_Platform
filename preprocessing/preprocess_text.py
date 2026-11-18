"""MySQL 응답을 정제하고 군집 탐색용 Excel을 명시적으로 저장한다."""

from __future__ import annotations

import argparse
from collections.abc import Callable, Sequence
from io import BytesIO
from pathlib import Path
import re

import pandas as pd

from preprocessing.database import fetch_responses

PUNCTUATION = "/-'?!.,#$%'()*+-/:;<=>@[\\]^_`{|}~" + '"“”’' + "∞θ÷α•à−β∅³π‘₹´°£€\\×™√²—–&"
PUNCTUATION_MAPPING = {
    "‘": "'", "₹": "e", "´": "'", "°": "", "€": "e", "™": "tm",
    "√": " sqrt ", "×": "x", "²": "2", "—": "-", "–": "-", "’": "'",
    "_": "-", "`": "'", "“": '"', "”": '"', "£": "e", "∞": "infinity",
    "θ": "theta", "÷": "/", "α": "alpha", "•": ".", "à": "a", "−": "-",
    "β": "beta", "∅": "", "³": "3", "π": "pi",
}


def clean_text(text: str) -> str:
    """HTML·이메일·URL을 먼저 제거한 뒤 기존 문자 치환과 공백 정리를 적용한다."""
    if not isinstance(text, str):
        raise ValueError("응답은 결측값이 없는 문자열이어야 합니다.")
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", "", text)
    text = re.sub(r"(?:https?|ftp)://[^\s<]+", " ", text)
    for source, replacement in PUNCTUATION_MAPPING.items():
        text = text.replace(source, replacement)
    for punctuation in PUNCTUATION:
        text = text.replace(punctuation, "")
    text = re.sub(r"[ㄱ-ㅎㅏ-ㅣ]+", "", text)
    text = re.sub(r"[^\w\s\n]", "", text)
    return re.sub(r"\s+", " ", text).strip()


def preprocess_responses(responses: Sequence[str] | pd.Series, *, spacing: Callable[[str], str], spellcheck: Callable[[str], str], tokenize: Callable[[str], str]) -> pd.DataFrame:
    """주입된 처리기로 정제→띄어쓰기→맞춤법→명사 추출을 수행한다.

    외부 호출 여부는 처리기에 달려 있다. 결측·비문자열 응답은 외부 처리
    전에 거부하며 행을 임의로 삭제하거나 문자열로 변환하지 않는다.
    """
    source = responses.copy() if isinstance(responses, pd.Series) else pd.Series(list(responses), dtype=object)
    if source.empty:
        raise ValueError("전처리할 응답이 없습니다.")
    if not source.map(lambda value: isinstance(value, str)).all():
        raise ValueError("응답은 결측값이 없는 문자열이어야 합니다.")
    data = pd.DataFrame({"text": source.map(clean_text)})
    for input_column, output_column, processor in (
        ("text", "corrected_text", spacing),
        ("corrected_text", "spell_check_text", spellcheck),
        ("spell_check_text", "tokenized_text", tokenize),
    ):
        data[output_column] = data[input_column].map(processor)
        if not data[output_column].map(lambda value: isinstance(value, str)).all():
            raise ValueError(f"{output_column} 처리기가 문자열을 반환하지 않았습니다.")
    return data


def create_nlp_processors(*, allow_external_spellcheck: bool = False) -> tuple[Callable, Callable, Callable]:
    """실제 NLP 처리기를 한 번씩 만들며 외부 맞춤법 검사 사용을 명시적으로 요구한다."""
    if not allow_external_spellcheck:
        raise ValueError("맞춤법 검사는 외부 서비스에 문장을 보냅니다. 외부 전송 허용이 필요합니다.")
    from pykospacing import Spacing
    from hanspell import spell_checker
    from konlpy.tag import Okt

    spacing = Spacing()
    okt = Okt()

    def spellcheck(text: str) -> str:
        """외부 맞춤법 검사 결과를 반환하고 실패 결과는 조용히 대체하지 않는다."""
        result = spell_checker.check(text)
        if not result.result:
            raise ValueError("맞춤법 검사에 실패했습니다. 입력 길이와 서비스 상태를 확인하세요.")
        return result.checked

    def tokenize(text: str) -> str:
        """재사용하는 Okt 객체로 명사를 추출해 공백으로 연결한다."""
        return " ".join(okt.nouns(text))

    return spacing, spellcheck, tokenize


def check_output_path(path: str | Path, *, overwrite: bool = False) -> Path:
    """Excel 출력 위치와 덮어쓰기 허용 여부를 쓰기 전에 검사한다."""
    target = Path(path)
    if target.suffix.lower() != ".xlsx":
        raise ValueError("출력 확장자는 .xlsx여야 합니다.")
    if not target.parent.is_dir():
        raise ValueError("출력 폴더가 존재하지 않습니다.")
    if target.exists() and (not overwrite or not target.is_file()):
        raise FileExistsError("출력 파일이 이미 있습니다. 다른 경로나 명시적 덮어쓰기를 선택하세요.")
    return target


def save_preprocessed_data(data: pd.DataFrame, path: str | Path, *, overwrite: bool = False) -> Path:
    """Excel 직렬화 성공 후 저장하며 기본적으로 기존 파일을 덮어쓰지 않는다."""
    target = check_output_path(path, overwrite=overwrite)
    buffer = BytesIO()
    data.to_excel(buffer, index=False, engine="openpyxl")
    with target.open("wb" if overwrite else "xb") as output:
        output.write(buffer.getvalue())
    return target


def main(argv: list[str] | None = None) -> int:
    """명시적 실행에서만 DB 조회·외부 검사·Excel 저장을 수행한다."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--table", default="food_project", help="response 열이 있는 MySQL 테이블")
    parser.add_argument("--output", type=Path, default=Path("preprocessing_data.xlsx"))
    parser.add_argument("--overwrite", action="store_true", help="기존 출력 파일 덮어쓰기 허용")
    parser.add_argument("--allow-external-spellcheck", action="store_true", help="문장을 외부 맞춤법 검사 서비스로 전송하는 것을 허용")
    args = parser.parse_args(argv)
    if not args.allow_external_spellcheck:
        parser.error("실제 맞춤법 검사를 사용하려면 --allow-external-spellcheck가 필요합니다.")
    try:
        target = check_output_path(args.output, overwrite=args.overwrite)
        spacing, spellcheck, tokenize = create_nlp_processors(allow_external_spellcheck=True)
        responses = fetch_responses(args.table)
        result = preprocess_responses(responses, spacing=spacing, spellcheck=spellcheck, tokenize=tokenize)
        save_preprocessed_data(result, target, overwrite=args.overwrite)
    except Exception as error:
        print(f"전처리 실패: {type(error).__name__}. 입력·환경변수·NLP 환경·출력 경로를 확인하세요.")
        return 2
    print(f"전처리 완료: {len(result)}개 응답을 지정한 Excel에 저장했습니다.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
