"""환경변수의 접속 설정을 사용해 MySQL의 응답 열만 조회한다."""

from __future__ import annotations

from collections.abc import Callable, Mapping
import os
import re


def read_db_config(environ: Mapping[str, str] | None = None) -> dict:
    """접속 환경변수를 읽고 누락·포트 오류를 값 노출 없이 보고한다."""
    source = os.environ if environ is None else environ
    names = {"host": "MYSQL_HOST", "user": "MYSQL_USER", "password": "MYSQL_PASSWORD", "database": "MYSQL_DATABASE"}
    missing = [name for name in names.values() if not source.get(name)]
    if missing:
        raise ValueError("필수 환경변수 누락: " + ", ".join(missing))
    try:
        port = int(source.get("MYSQL_PORT", "3306"))
    except ValueError:
        raise ValueError("MYSQL_PORT는 정수여야 합니다.") from None
    if not 1 <= port <= 65535:
        raise ValueError("MYSQL_PORT 범위를 확인하세요.")
    return {**{key: source[name] for key, name in names.items()}, "port": port}


def fetch_responses(table_name: str = "food_project", *, config: Mapping | None = None, connect: Callable | None = None) -> list:
    """지정 테이블의 response를 SELECT하고 성공·실패 모두 연결을 닫는다.

    테이블 식별자는 영문·숫자·밑줄만 허용한다. connect 인자로 테스트용
    연결을 주입할 수 있으며, INSERT·UPDATE·COMMIT은 수행하지 않는다.
    """
    if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", table_name):
        raise ValueError("테이블명은 영문 또는 밑줄로 시작하는 영문·숫자·밑줄이어야 합니다.")
    settings = read_db_config() if config is None else dict(config)
    if connect is None:
        import mysql.connector

        connect = mysql.connector.connect
    connection = connect(**settings)
    cursor = None
    try:
        cursor = connection.cursor()
        cursor.execute(f"SELECT `response` FROM `{table_name}`")
        return [record[0] for record in cursor.fetchall()]
    finally:
        try:
            if cursor is not None:
                cursor.close()
        finally:
            connection.close()
