from __future__ import annotations

import importlib

from fastapi.testclient import TestClient

app_module = importlib.import_module("red_movie_xlsx_tool.app")


def test_schema_endpoint_returns_default_questions_and_separator() -> None:
    client = TestClient(app_module.app)
    response = client.get("/api/schema")
    assert response.status_code == 200
    data = response.json()
    assert data["title"]
    assert len(data["questions"]) == 18
    assert data["multi_separator"] == "┋"
    assert data["multi_separator_codepoint"] == "U+250B"


def test_generate_endpoint_creates_downloadable_split_250b_file(tmp_path, monkeypatch) -> None:
    monkeypatch.setattr(app_module, "GENERATED_DIR", tmp_path)
    client = TestClient(app_module.app)
    schema = client.get("/api/schema").json()

    response = client.post(
        "/api/generate",
        json={"count": 5, "seed": "api", "questions": schema["questions"]},
    )

    assert response.status_code == 200
    data = response.json()
    assert data["filename"].startswith("red_movie_wjx_split_250b_")
    assert data["row_count"] == 5
    assert data["question_count"] == 18
    assert data["multi_separator"] == "┋"
    assert data["multi_separator_codepoint"] == "U+250B"
    assert data["summary"]["multi_separator"] == "┋"
    assert data["summary"]["multi_separator_codepoint"] == "U+250B"
    assert data["summary"]["unique_answer_rows"] >= 1
    assert (tmp_path / data["filename"]).exists()

    download = client.get(data["download_url"])
    assert download.status_code == 200
    assert download.content.startswith(b"PK")


def test_generate_endpoint_rejects_legacy_pipe_mode_fields(tmp_path, monkeypatch) -> None:
    monkeypatch.setattr(app_module, "GENERATED_DIR", tmp_path)
    client = TestClient(app_module.app)
    schema = client.get("/api/schema").json()

    response = client.post(
        "/api/generate",
        json={"count": 1, "seed": "api", "multi_mode": "joined", "multi_separator": "|", "questions": schema["questions"]},
    )

    assert response.status_code == 200
    data = response.json()
    assert data["multi_separator"] == "┋"
    assert data["summary"]["multi_separator"] == "┋"


def test_download_rejects_path_traversal() -> None:
    client = TestClient(app_module.app)
    response = client.get("/api/download/..%5Csecret.xlsx")
    assert response.status_code == 404
