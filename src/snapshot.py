"""
법령 스냅샷 저장/로드.
스냅샷 파일: data/{법령명}.json

구조:
{
  "법령명": "산업안전보건법",
  "공포일자": "20230101",
  "MST": "12345",
  "조회일": "2026-05-28",
  "조문": [{"번호": "1", "제목": "목적", "내용": "..."}, ...]
}
"""
import json
import os


def load(law_name: str, data_dir: str) -> dict | None:
    path = os.path.join(data_dir, f"{law_name}.json")
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def save(snapshot: dict, data_dir: str) -> None:
    os.makedirs(data_dir, exist_ok=True)
    law_name = snapshot["법령명"]
    safe_name = law_name.replace("/", "_").replace("\\", "_")
    path = os.path.join(data_dir, f"{safe_name}.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(snapshot, f, ensure_ascii=False, indent=2)
