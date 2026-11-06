"""
FastAPI web endpoints test.
"""
from fastapi.testclient import TestClient
from pp_to_excel.web import app

client = TestClient(app)


def test_index_page():
    res = client.get("/")
    assert res.status_code == 200
    assert "SMT Pick & Place to Excel" in res.text
    assert "Modül 1" in res.text
    assert "Modül 2" in res.text


def test_analyze_bom_endpoint():
    with open("canıasbom feka.xlsx", "rb") as f:
        res = client.post("/api/analyze-bom", files={"bom_file": ("canıasbom feka.xlsx", f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")})
    assert res.status_code == 200
    data = res.json()
    assert data["success"] is True
    assert "ST1" in data["detected_prefixes"]
    assert data["smd_item_count"] > 0
