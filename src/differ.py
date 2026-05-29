"""
두 법령 스냅샷 간 조문 diff.
"""
import re


def _sort_key(번호: str) -> tuple:
    m = re.match(r"^(\d+)", 번호)
    return (int(m.group(1)), 번호) if m else (0, 번호)


def diff(old: dict, new: dict) -> dict:
    """
    반환:
    {
      "추가": [{"번호", "제목", "내용"}, ...],
      "삭제": [{"번호", "제목", "내용"}, ...],
      "변경": [{"번호", "이전": {...}, "현재": {...}}, ...],
    }
    """
    old_map = {a["번호"]: a for a in old.get("조문", [])}
    new_map = {a["번호"]: a for a in new.get("조문", [])}

    old_keys = set(old_map)
    new_keys = set(new_map)

    added   = sorted([new_map[k] for k in new_keys - old_keys], key=lambda x: _sort_key(x["번호"]))
    deleted = sorted([old_map[k] for k in old_keys - new_keys], key=lambda x: _sort_key(x["번호"]))
    changed = []
    for k in sorted(old_keys & new_keys, key=_sort_key):
        o, n = old_map[k], new_map[k]
        if o.get("제목") != n.get("제목") or o.get("내용") != n.get("내용"):
            changed.append({"번호": k, "이전": o, "현재": n})

    return {"추가": added, "삭제": deleted, "변경": changed}
