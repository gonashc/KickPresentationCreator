from pathlib import Path

from streamlit.testing.v1 import AppTest


def test_app_renders_upload_workflow():
    app = AppTest.from_file(Path(__file__).parents[1] / "app.py")
    app.run(timeout=20)

    assert not app.exception
    assert app.title[0].value == "Build a client-ready kickoff deck"
    assert len(app.get("file_uploader")) == 2
    assert app.button[0].disabled is True

