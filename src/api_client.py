"""
국가법령정보센터 Open API 클라이언트
- 재시도 3회 (대기 시간 점진적 증가)
- 요청 간 0.5초 딜레이 (API 제한 대응)
- UTF-8 → EUC-KR 인코딩 fallback
"""
import time
import requests

BASE_SEARCH_URL = "http://www.law.go.kr/DRF/lawSearch.do"
BASE_SERVICE_URL = "http://www.law.go.kr/DRF/lawService.do"

MAX_RETRIES = 3
REQUEST_DELAY = 0.5
REQUEST_TIMEOUT = 30


class APIError(Exception):
    pass


class LawAPIClient:
    def __init__(self, api_key: str) -> None:
        self.api_key = api_key
        self.session = requests.Session()
        self._last_request_time: float = 0

    def _wait(self) -> None:
        elapsed = time.time() - self._last_request_time
        if elapsed < REQUEST_DELAY:
            time.sleep(REQUEST_DELAY - elapsed)

    def _decode(self, response: requests.Response) -> str:
        try:
            return response.content.decode("utf-8")
        except UnicodeDecodeError:
            return response.content.decode("euc-kr")

    def _get(self, url: str, params: dict) -> str:
        params = {**params, "OC": self.api_key, "type": "XML"}
        self._wait()

        for attempt in range(MAX_RETRIES):
            try:
                resp = self.session.get(url, params=params, timeout=REQUEST_TIMEOUT)
                resp.raise_for_status()
                self._last_request_time = time.time()
                return self._decode(resp)
            except requests.RequestException as e:
                if attempt < MAX_RETRIES - 1:
                    wait = attempt + 1
                    print(f"  [!] 요청 실패 (시도 {attempt + 1}/{MAX_RETRIES}): {e}. {wait}초 후 재시도...")
                    time.sleep(wait)
                else:
                    raise APIError(f"API 요청 최종 실패: {e}") from e

    def search_law(self, query: str) -> str:
        return self._get(BASE_SEARCH_URL, {"target": "law", "query": query, "display": 20, "page": 1})

    def get_law_articles(self, mst: str) -> str:
        return self._get(BASE_SERVICE_URL, {"target": "law", "MST": mst})

    def get_old_new(self, mst: str) -> str:
        """신구법비교 조회 (target=oldAndNew)."""
        return self._get(BASE_SERVICE_URL, {"target": "oldAndNew", "MST": mst})
