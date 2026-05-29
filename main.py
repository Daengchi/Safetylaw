"""
안전보건 법규 제개정사항 모니터링 - CLI

사용 예시:
  python main.py                        # laws.json 전체 법령 확인 후 리포트 생성
  python main.py --law 산업안전보건법   # 특정 법령만 확인
  python main.py --no-report            # 스냅샷 업데이트만 (리포트 생성 안 함)
  python main.py --output ./output      # 출력 디렉토리 지정
  python main.py --debug                # 신구법비교 원본 XML 저장 (파싱 문제 진단용)
"""
import argparse
import json
import os
import sys
from datetime import datetime

from dotenv import load_dotenv

from src.api_client import APIError, LawAPIClient
from src import parser, snapshot, reporter


DATA_DIR   = os.path.join(os.path.dirname(__file__), "data")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "output")
LAWS_FILE  = os.path.join(os.path.dirname(__file__), "..", "laws.json")


def _norm(s: str) -> str:
    """법령명 비교용: 공백 제거."""
    return s.replace(" ", "")


def _pick_law(laws: list[dict], query: str) -> dict | None:
    """검색 결과에서 입력 법령명과 가장 일치하는 항목 반환."""
    for law in laws:
        if law["name"] == query:
            return law
    for law in laws:
        if _norm(query) == _norm(law["name"]):
            return law
    for law in laws:
        if query in law["name"] or _norm(query) in _norm(law["name"]):
            return law
    return laws[0] if laws else None


def _find_related(laws: list[dict], base_name: str) -> list[dict]:
    """
    검색 결과에서 base_name 법률 + 시행령 + 시행규칙을 순서대로 추출.
    각각 없으면 포함 안 함. 공백 정규화 비교 포함.
    """
    result = []
    base_norm = _norm(base_name)

    # 법률 (정확 일치 우선 → 공백 정규화 일치 → 부분 일치)
    main = next((l for l in laws if l["name"] == base_name), None)
    if main is None:
        main = next(
            (l for l in laws if _norm(l["name"]) == base_norm
             and "시행령" not in l["name"] and "시행규칙" not in l["name"]),
            None,
        )
    if main is None:
        main = next(
            (l for l in laws if base_norm in _norm(l["name"])
             and "시행령" not in l["name"] and "시행규칙" not in l["name"]),
            None,
        )
    if main:
        result.append(main)

    # 시행령 (정확 일치 → 공백 정규화 → startswith)
    enf = next((l for l in laws if l["name"] == base_name + " 시행령"), None)
    if enf is None:
        enf = next(
            (l for l in laws if "시행령" in l["name"]
             and _norm(l["name"]).startswith(base_norm)),
            None,
        )
    if enf:
        result.append(enf)

    # 시행규칙 (정확 일치 → 공백 정규화 → startswith)
    rul = next((l for l in laws if l["name"] == base_name + " 시행규칙"), None)
    if rul is None:
        rul = next(
            (l for l in laws if "시행규칙" in l["name"]
             and _norm(l["name"]).startswith(base_norm)),
            None,
        )
    if rul:
        result.append(rul)

    return result


def _discover_laws(client: LawAPIClient, law_name: str) -> list[dict]:
    """
    법령명에 대한 관련 법령 목록 반환.
    - 법령명에 시행령/시행규칙이 포함된 경우: 단독 처리
    - 그 외: 법률 + 시행령 + 시행규칙 포함 목록 반환 (존재하는 것만)
    반환: [{"name": str, "mst": str, "공포일자": str}, ...]
    """
    print(f"  [{law_name}] 관련 법령 탐색 중...")
    try:
        search_xml = client.search_law(law_name)
    except APIError as e:
        print(f"    오류: {e}")
        return []

    laws = parser.parse_law_search(search_xml)
    if not laws:
        print(f"    '{law_name}' 검색 결과 없음")
        return []

    # laws.json에 시행령/시행규칙을 직접 지정한 경우: 단독 처리
    if "시행령" in law_name or "시행규칙" in law_name:
        matched = _pick_law(laws, law_name)
        if matched:
            kind = "시행령" if "시행령" in matched["name"] else "시행규칙"
            print(f"    {kind} (단독): {matched['name']}  시행일자: {matched['시행일자']}")
            return [matched]
        return []

    related = _find_related(laws, law_name)

    _KIND = lambda n: "시행규칙" if "시행규칙" in n else ("시행령" if "시행령" in n else "법률")
    for l in related:
        print(f"    {_KIND(l['name']):<6}: {l['name']}  시행일자: {l['시행일자']}")

    if not related:
        print(f"    '{law_name}' 매칭 실패")

    return related


def _check_law(client: LawAPIClient, law_info: dict, debug: bool = False) -> dict:
    """
    law_info: {"name": str, "mst": str, "공포일자": str}
    법령 1개를 확인하고 결과 dict 반환.
    반환: {"law_name", "old_date", "new_date", "status", "articles"}
    """
    law_name  = law_info["name"]
    mst       = law_info["mst"]
    new_date  = law_info["시행일자"]
    safe_name = law_name.replace("/", "_").replace("\\", "_")

    result: dict = {"law_name": law_name, "old_date": "", "new_date": new_date,
                    "status": "오류", "articles": [], "article_count": "-"}

    # Step 1: 기존 스냅샷 로드 (시행일자 변경 감지용)
    old_snap = snapshot.load(safe_name, DATA_DIR)
    old_date = old_snap.get("시행일자", "") if old_snap else ""
    result["old_date"] = old_date

    # Step 2: 시행일자 비교 — 동일하면 변경 없음
    if old_snap and old_date == new_date:
        print(f"    [{law_name}] 변경 없음 (시행일자: {new_date})")
        result["status"]        = "변경 없음"
        result["article_count"] = old_snap.get("개정_조문_수", "-")
        if old_snap.get("조문비교"):                       # 유효한 캐시 있음
            result["articles"] = old_snap["조문비교"]
        elif old_snap.get("신구법존재여부") == "N":         # 타법개정 미제공 확인됨
            result["타법개정"] = True
        else:                                               # 캐시 없음 → API 호출
            print(f"    [{law_name}] 신구법비교 초기 캐시 로드...")
            try:
                xml = client.get_old_new(mst)
                if debug:
                    debug_path = os.path.join(DATA_DIR, f"debug_oldNew_{safe_name}.xml")
                    with open(debug_path, "w", encoding="utf-8") as f:
                        f.write(xml)
                    print(f"      원본 XML 저장: {debug_path}")
                meta, articles = parser.parse_old_new(xml)
                result["articles"]      = articles
                result["article_count"] = len(articles)
                update = {
                    **old_snap,
                    "시행일자":    new_date,
                    "조문비교":    articles,
                    "개정_조문_수": len(articles),
                }
                if meta.get("신구법존재여부") == "N":
                    update["신구법존재여부"] = "N"
                    result["타법개정"] = True
                snapshot.save(update, DATA_DIR)
            except APIError as e:
                print(f"      캐시 로드 오류: {e}")
        return result

    # Step 3: 신구법비교 조회
    print(f"    [{law_name}] 신구법비교 조회 중...")
    try:
        xml = client.get_old_new(mst)
    except APIError as e:
        print(f"      오류: {e}")
        return result

    if debug:
        debug_path = os.path.join(DATA_DIR, f"debug_oldNew_{safe_name}.xml")
        with open(debug_path, "w", encoding="utf-8") as f:
            f.write(xml)
        print(f"      원본 XML 저장: {debug_path}")

    meta, articles = parser.parse_old_new(xml)
    print(f"      {len(articles)}개 조문 파싱")

    # Step 4: 스냅샷 갱신
    snap_data = {
        "법령명":     law_name,
        "공포일자":   law_info["공포일자"],   # 참고용
        "시행일자":   new_date,              # 변경 감지 기준
        "MST":       mst,
        "조회일":     datetime.now().strftime("%Y-%m-%d"),
        "개정_조문_수": len(articles),
        "조문비교":   articles,
    }
    if meta.get("신구법존재여부") == "N":
        snap_data["신구법존재여부"] = "N"
        result["타법개정"] = True
    snapshot.save(snap_data, DATA_DIR)

    status = "신규 등록" if old_snap is None else "개정됨"
    print(f"      {status}")

    result["status"]        = status
    result["articles"]      = articles
    result["article_count"] = len(articles)
    return result


def main() -> None:
    load_dotenv()

    arg_parser = argparse.ArgumentParser(
        description="안전보건 법규 제개정사항 모니터링",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    arg_parser.add_argument("--law", "-l", metavar="법령명", help="특정 법령만 확인 (미입력 시 laws.json 전체)")
    arg_parser.add_argument("--output", "-o", metavar="경로", default=OUTPUT_DIR, help="리포트 출력 디렉토리")
    arg_parser.add_argument("--no-report", action="store_true", help="스냅샷 업데이트만 수행 (리포트 생성 안 함)")
    arg_parser.add_argument("--debug", action="store_true", help="신구법비교 원본 XML을 data/ 에 저장 (파싱 문제 진단용)")
    args = arg_parser.parse_args()

    api_key = os.getenv("LAW_API_KEY")
    if not api_key:
        print("오류: LAW_API_KEY 환경변수가 없습니다.")
        print("  .env 파일에 다음을 추가하세요:  LAW_API_KEY=인증키")
        sys.exit(1)

    # 모니터링 대상 법령 목록
    if args.law:
        law_names = [args.law]
    else:
        if not os.path.exists(LAWS_FILE):
            print(f"오류: {LAWS_FILE} 파일이 없습니다.")
            print("  laws.json에 모니터링할 법령 목록을 작성하세요.")
            sys.exit(1)
        with open(LAWS_FILE, encoding="utf-8") as f:
            law_names = json.load(f)

    if not law_names:
        print("모니터링 대상 법령이 없습니다.")
        sys.exit(0)

    client     = LawAPIClient(api_key)
    results    = []
    law_groups = []   # 법규 목록표용: [{"parent": str, "laws": [law_info, ...]}]

    print(f"\n총 {len(law_names)}개 항목 확인 시작\n{'─' * 50}")
    for law_name in law_names:
        related = _discover_laws(client, law_name)
        if not related:
            results.append({"law_name": law_name, "old_date": "", "new_date": "", "status": "오류", "diff": {}})
            continue
        law_groups.append({"parent": law_name, "laws": related})
        for law_info in related:
            results.append(_check_law(client, law_info, debug=args.debug))
    print(f"{'─' * 50}")

    # 결과 요약 출력
    changed   = [r for r in results if r["status"] == "개정됨"]
    new_regs  = [r for r in results if r["status"] == "신규 등록"]
    no_change = [r for r in results if r["status"] == "변경 없음"]
    errors    = [r for r in results if r["status"] == "오류"]

    print(f"\n확인 완료: 신규 등록 {len(new_regs)}건 / 개정 {len(changed)}건 / 변경 없음 {len(no_change)}건 / 오류 {len(errors)}건")

    if args.no_report:
        print("--no-report 옵션으로 리포트 생성을 건너뜁니다.")
        return

    os.makedirs(args.output, exist_ok=True)
    date_str    = datetime.now().strftime("%Y%m%d_%H%M")
    output_path = os.path.join(args.output, f"법규_제개정_모니터링_{date_str}.xlsx")

    print(f"\nExcel 리포트 생성 중...")
    reporter.generate(results, output_path, law_groups=law_groups)
    print(f"완료!  ->  {output_path}")


if __name__ == "__main__":
    main()
