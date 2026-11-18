"""모듈을 가져오는 것만으로 업무 처리가 시작되지 않는지 확인한다."""

import os
from pathlib import Path
import subprocess
import sys


def test_imports_do_not_start_io_or_load_optional_nlp(tmp_path):
    """새 프로세스에서 파일 입출력·DB·외부 연결을 차단한 채 모듈을 가져온다."""
    project = Path(__file__).resolve().parents[1]
    script = '''
import socket
import sys
from unittest.mock import Mock
import pandas as pd
pd.ExcelFile = Mock(side_effect=AssertionError("Excel 읽기 발생"))
pd.read_excel = Mock(side_effect=AssertionError("Excel 읽기 발생"))
pd.DataFrame.to_excel = Mock(side_effect=AssertionError("Excel 쓰기 발생"))
socket.socket.connect = Mock(side_effect=AssertionError("외부 연결 발생"))
sys.path.insert(0, sys.argv[1])
import validation.validate_excel
import preprocessing.database
import preprocessing.preprocess_text
import exploration.cluster_texts
assert not any(name in sys.modules for name in ("mysql.connector", "pykospacing", "hanspell", "konlpy", "matplotlib.pyplot"))
'''
    result = subprocess.run([sys.executable, "-X", "utf8", "-B", "-c", script, str(project)], cwd=tmp_path, env=os.environ.copy(), capture_output=True, text=True, encoding="utf-8", timeout=60)
    assert result.returncode == 0, result.stderr
    assert list(tmp_path.iterdir()) == []
