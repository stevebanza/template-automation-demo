"""The Streamlit checker shows STR01 for a damaged upload instead of a traceback."""
from pathlib import Path

import pytest

AppTest = pytest.importorskip("streamlit.testing.v1").AppTest
APP = str(Path(__file__).resolve().parents[1] / "app.py")


def _upload(at, name: str, content: bytes):
    uploader = at.main.file_uploader[0]
    if not hasattr(uploader, "set_value"):
        pytest.skip("this Streamlit version cannot simulate uploads")
    uploader.set_value((name, content, "application/octet-stream"))
    return at.run(timeout=30)


OLE_FILE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + bytes(64)  # how password-protected .xlsx look


@pytest.mark.parametrize("content", [b"not a workbook", OLE_FILE],
                         ids=["not-excel", "password-protected"])
def test_damaged_upload_shows_str01(content):
    at = _upload(AppTest.from_file(APP).run(timeout=30), "XYZ_2025.xlsx", content)
    assert not at.exception
    assert at.metric[0].value == "Returned to country"
    table = at.dataframe[0].value
    assert list(table["Rule"]) == ["STR01"]
    assert list(table["Level"]) == ["Must fix"]
