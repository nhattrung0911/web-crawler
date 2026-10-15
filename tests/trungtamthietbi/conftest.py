import pathlib
import sys
import tempfile

import pytest

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[2] / "_vn" / "trungtamthietbi_com"))


@pytest.fixture
def tmp_path():
    """Thu muc tam rieng (tempfile) - tranh pytest-of-<user> trong %TEMP% hay bi loi ACL tren Windows."""
    with tempfile.TemporaryDirectory(prefix="ttb_", ignore_cleanup_errors=True) as d:
        yield pathlib.Path(d)


@pytest.fixture(autouse=True)
def _isolate_runtime_dirs(tmp_path, monkeypatch):
    """Khong test nao duoc ghi log/checkpoint vao out_trungtamthietbi/ THAT cua crawler."""
    import trungtamthietbi as t
    iso = tmp_path / "_iso_out"
    monkeypatch.setattr(t, "OUT_DIR", iso)
    monkeypatch.setattr(t, "LOG_FILE", iso / "trungtamthietbi.log")
