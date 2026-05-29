"""
XML 파싱 모듈.

lawSearch.do?target=law 응답 구조:
  <LawSearch>
    <law>
      <법령명한글>...</법령명한글>
      <법령일련번호>...</법령일련번호>
      <공포일자>...</공포일자>
      <시행일자>...</시행일자>
    </law>
    ...
  </LawSearch>

lawService.do?target=oldAndNew 실제 응답 구조:
  <OldAndNewService>
    <구조문_기본정보>
      <공포일자>20240206</공포일자>
      <시행일자>20250807</시행일자>
    </구조문_기본정보>
    <신조문_기본정보>
      <공포일자>20251001</공포일자>
      <시행일자>20251001</시행일자>
    </신조문_기본정보>
    <구조문목록>
      <조문 no="1"><![CDATA[제2조(정의) ...]]></조문>
      <조문 no="2"><![CDATA[1. ∼ 5. (생  략)]]></조문>
      ...
    </구조문목록>
    <신조문목록>
      <조문 no="1"><![CDATA[제2조(정의) ...]]></조문>
      <조문 no="2"><![CDATA[1. ∼ 5. (현행과 같음)]]></조문>
      ...
    </신조문목록>
  </OldAndNewService>

  - 각 <조문 no="N">은 CDATA로 텍스트를 감싼 단일 단락
  - <P>...</P> 는 변경된 텍스트를 마킹하는 인라인 태그
  - 구법의 (생  략) ↔ 신법의 (현행과 같음) 은 동일 내용
  - 구조문목록 / 신조문목록의 no 속성으로 1:1 대응
"""
import re
import xml.etree.ElementTree as ET
from typing import Optional

_CONTENT_TAGS = {"조문내용", "항내용", "호내용", "목내용", "단서내용"}

# <P>변경텍스트</P> → [변경텍스트] 치환
_P_TAG_RE = re.compile(r'</?P>', re.IGNORECASE)
# 조문 시작 패턴: "제N조" 또는 "제N조의M"
_ARTICLE_START_RE = re.compile(r'(제\d+조(?:의\d+)?)(?:\(([^)]+)\))?')
_HAS_P_RE = re.compile(r'<P>', re.IGNORECASE)
# 별표/별지/부칙 시작 패턴
_ANNEX_START_RE = re.compile(r'^(별표\s*\d+|별지\s+제?\d+호|부칙(?:\s*제?\d+조)?)', re.IGNORECASE)


def _fmt_date(date_str: str) -> str:
    """20251001 → 2025/10/01"""
    if date_str and len(date_str) == 8 and date_str.isdigit():
        return f"{date_str[:4]}/{date_str[4:6]}/{date_str[6:]}"
    return date_str


def _parse_xml(xml_text: str) -> ET.Element:
    xml_text = xml_text.lstrip("﻿")  # UTF-8 BOM 제거
    try:
        return ET.fromstring(xml_text)
    except ET.ParseError:
        return ET.fromstring(xml_text.encode("utf-8"))


def _text(elem: Optional[ET.Element]) -> str:
    if elem is None or elem.text is None:
        return ""
    return elem.text.strip()


def _find(root: ET.Element, *tags: str) -> Optional[ET.Element]:
    for tag in tags:
        elem = root.find(f".//{tag}")
        if elem is not None:
            return elem
    return None


def _collect_content(unit: ET.Element) -> str:
    return " ".join(
        elem.text.strip()
        for elem in unit.iter()
        if elem.tag in _CONTENT_TAGS and elem.text
    )


def parse_law_search(xml_text: str) -> list[dict]:
    """
    법령 검색 결과 XML 파싱.
    반환: [{"name": 법령명, "mst": 법령일련번호, "공포일자": str,
            "시행일자": str, "제개정구분명": str}, ...]
    """
    root = _parse_xml(xml_text)
    laws: list[dict] = []
    for item in root.iter("law"):
        name = _text(_find(item, "법령명한글", "법령명_한글", "법령명"))
        mst  = _text(_find(item, "법령일련번호", "MST", "LST", "일련번호"))
        date = _text(_find(item, "공포일자", "개정일자"))
        ef   = _text(_find(item, "시행일자"))
        kind = _text(_find(item, "제개정구분명"))
        if name and mst:
            laws.append({
                "name": name,
                "mst": mst,
                "공포일자": date,
                "시행일자": ef,
                "제개정구분명": kind,
            })
    return laws


def parse_law_articles(xml_text: str) -> list[dict]:
    """
    법령 전체 조문 파싱 (내용 포함).
    '조문여부'가 '조문'인 항목만 포함 (장·절 구분 제외).
    반환: [{"번호": str, "제목": str, "내용": str}, ...]
    """
    root = _parse_xml(xml_text)
    articles: list[dict] = []

    for unit in root.iter("조문단위"):
        kind = _text(unit.find("조문여부"))
        if kind and kind != "조문":
            continue
        num   = _text(unit.find("조문번호"))
        title = _text(unit.find("조문제목"))
        if not num:
            continue
        articles.append({
            "번호": num,
            "제목": title,
            "내용": _collect_content(unit),
        })

    return articles


def _strip_p(text: str) -> str:
    """<P>...</P> 태그 제거 (텍스트는 유지)."""
    return _P_TAG_RE.sub("", text or "").strip()


def _normalize(text: str) -> str:
    """비교용 정규화: P태그·(생략)·(현행과 같음)·공백 제거."""
    text = _P_TAG_RE.sub("", text or "")
    text = re.sub(r'\(생\s*략\)', '', text)
    text = re.sub(r'\(현행과\s*같음\)', '', text)
    return re.sub(r'\s+', ' ', text).strip()


def _read_jo_map(parent: Optional[ET.Element]) -> dict[tuple, str]:
    """
    구조문목록 또는 신조문목록에서 {(sort_key, no): text} 딕셔너리 반환.
    정수 no → (0, int) 순서 우선; 별표/별지 등 비정수 → (1, str) 후순위.
    """
    result: dict[tuple, str] = {}
    if parent is None:
        return result
    for jo in parent.findall("조문"):
        no_str = jo.get("no")
        if no_str and jo.text:
            try:
                key: tuple = (0, int(no_str))
            except ValueError:
                key = (1, no_str)
            result[key] = jo.text
    return result


def parse_old_new(xml_text: str) -> tuple[dict, list[dict]]:
    """
    신구법비교 API 파싱 (target=oldAndNew).

    반환:
      meta:     {"공포일자": str, "시행일자": str}
      articles: [{"조문번호": str, "조문명": str, "개정일": str,
                  "시행일자": str, "구법내용": str, "신법내용": str}, ...]

    처리 흐름:
      1. 구조문목록 / 신조문목록을 no 속성 기준으로 1:1 대응
      2. "제N조" 패턴으로 조문 단위 그룹핑
      3. 조문 전체 텍스트를 비교하여 실제 변경된 조문만 추출
      4. (생  략) ↔ (현행과 같음) 은 동일 내용으로 간주
    """
    root = _parse_xml(xml_text)

    # 메타데이터 (신조문 기준, fallback 구조문)
    new_info = root.find("신조문_기본정보")
    old_info = root.find("구조문_기본정보")
    meta = {
        "공포일자": _fmt_date(_text(_find(new_info or root, "공포일자"))),
        "시행일자": _fmt_date(_text(_find(new_info or root, "시행일자"))),
        "구_공포일자": _fmt_date(_text(_find(old_info or root, "공포일자"))),
    }

    # 신구법비교 데이터 존재 여부 확인 (타법개정 시 N)
    존재여부 = root.find("신구법존재여부")
    if 존재여부 is not None and 존재여부.text == "N":
        meta["신구법존재여부"] = "N"
        return meta, []
    meta["신구법존재여부"] = "Y"

    old_map = _read_jo_map(root.find("구조문목록"))
    new_map = _read_jo_map(root.find("신조문목록"))
    all_nos = sorted(set(old_map) | set(new_map))

    # 조문 단위 그룹핑: "제N조" 패턴으로 새 조문 시작 감지
    groups: list[dict] = []
    current: dict | None = None

    for no in all_nos:
        old_raw = old_map.get(no, "")
        new_raw = new_map.get(no, "")
        ref_text = old_raw or new_raw

        m_art = _ARTICLE_START_RE.search(ref_text)
        is_article = m_art and ref_text.lstrip().startswith(m_art.group(0))

        m_annex = (not is_article) and _ANNEX_START_RE.match(ref_text.lstrip())

        if is_article:
            if current is not None:
                groups.append(current)
            current = {
                "num":   m_art.group(1),
                "title": m_art.group(2) or "",
                "old":   [old_raw],
                "new":   [new_raw],
            }
        elif m_annex:
            # 별표/별지/부칙: 별도 그룹으로 처리
            if current is not None:
                groups.append(current)
            current = {
                "num":   m_annex.group(1).strip(),
                "title": "",
                "old":   [old_raw],
                "new":   [new_raw],
            }
        else:
            if current is None:
                current = {"num": "", "title": "", "old": [], "new": []}
            current["old"].append(old_raw)
            current["new"].append(new_raw)

    if current is not None:
        groups.append(current)

    # 변경된 조문만 추출
    articles: list[dict] = []
    for g in groups:
        old_texts = g["old"]
        new_texts = g["new"]
        old_full = "\n".join(t for t in old_texts if t)
        new_full = "\n".join(t for t in new_texts if t)

        if _normalize(old_full) == _normalize(new_full):
            continue  # 실질 변경 없음

        # 개정: <P> 태그 포함 항만 표시 (항상 조문 헤더 포함)
        # 신설·삭제: 전체 내용 표시
        if old_full.strip() and new_full.strip():
            f_old, f_new = [], []
            for i, (o, n) in enumerate(zip(old_texts, new_texts)):
                if i == 0 or _HAS_P_RE.search(o) or _HAS_P_RE.search(n):
                    f_old.append(o)
                    f_new.append(n)
            old_combined = "\n".join(t for t in f_old if t)
            new_combined = "\n".join(t for t in f_new if t)
        else:
            old_combined = old_full
            new_combined = new_full

        articles.append({
            "조문번호": g["num"],
            "조문명":  g["title"],
            "개정일":  meta["공포일자"],
            "시행일자": meta["시행일자"],
            "구법내용": old_combined,
            "신법내용": new_combined,
        })

    return meta, articles
