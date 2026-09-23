import os
import re
import sys
import json
import urllib.parse
import requests
from bs4 import BeautifulSoup

# 콘솔 출력 UTF-8 설정
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

WORKSPACE_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
SETTING_FILE = os.path.join(WORKSPACE_ROOT, "setting", "contents_set_url.md")
COOKIE_FILE = os.path.join(WORKSPACE_ROOT, "setting", "lcdms_cookie.txt")
INPUT_DIR = os.path.join(WORKSPACE_ROOT, "input")

# ==============================================================================
# [매우 중요] 절대 안전 보안 수칙 (Strict Read-Only Enforcement)
# 1. 시스템의 어떤 데이터도 '수정', '삭제', '등록', '저장', '편집'하지 않습니다.
# 2. 오직 HTTP GET 요청을 통한 웹페이지 '조회'와 첨부파일 '다운로드'만 수행합니다.
# 3. 로그인 및 세션 유지를 위한 조회 API 외에 일체의 쓰기/변형 요청은 없습니다.
# ==============================================================================

def load_subject_urls():
    """setting/contents_set_url.md에서 과목명과 URL 매핑 로드"""
    if not os.path.exists(SETTING_FILE):
        return {}
    
    mapping = {}
    with open(SETTING_FILE, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or not line.startswith("-"):
                continue
            match = re.match(r"^-\s*([^:]+)\s*:\s*(https?://[^\s]+)", line)
            if match:
                subject = match.group(1).strip()
                url = match.group(2).strip()
                mapping[subject] = url
    return mapping

def load_cookies():
    """setting/lcdms_cookie.txt에서 쿠키 로드"""
    cookies = {}
    if os.path.exists(COOKIE_FILE):
        with open(COOKIE_FILE, "r", encoding="utf-8") as f:
            raw = f.read().replace("\n", " ").strip()
            for item in raw.split(";"):
                item = item.strip()
                if "=" in item:
                    k, v = item.split("=", 1)
                    cookies[k.strip()] = v.strip()
    return cookies

def js_escape_encode(s: str) -> str:
    """LCDMS JsCommon.Base64_encode와 동일한 escape(encodeURIComponent(s)) 인코딩"""
    return urllib.parse.quote(urllib.parse.quote(s, safe="-_.!~*'()"), safe="*+-./@")

def init_session(cookies: dict) -> requests.Session:
    """인증 세션 초기화 및 로그인 핸드셰이크"""
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36",
        "Referer": "https://home.iscu.ac.kr/cdms/index.jsp",
    })
    session.cookies.update(cookies)
    # SSO 토큰 바인딩을 위한 login.scu 초기화 (조회용 핸드셰이크)
    session.post("https://home.iscu.ac.kr/cdms/login/login.scu", data={"p": ""}, timeout=10)
    return session

def parse_contents_seqno(base_url: str) -> str:
    """URL에서 contents_seqno 파라미터 추출"""
    parsed = urllib.parse.urlparse(base_url)
    params = urllib.parse.parse_qs(parsed.query)
    return params.get("contents_seqno", [""])[0]

def fetch_and_download(subject: str, weeks: list[int]):
    """
    지정 과목 및 주차의 원고/SB 파일 다운로드 (엄격한 Read-Only 모드)
    """
    mapping = load_subject_urls()
    if subject not in mapping:
        print(f"[오류] 'setting/contents_set_url.md'에서 '{subject}' 과목을 찾을 수 없습니다.")
        print(f"등록된 과목 목록: {list(mapping.keys())}")
        return False

    base_url = mapping[subject]
    contents_seqno = parse_contents_seqno(base_url)
    if not contents_seqno:
        print(f"[오류] URL에서 contents_seqno를 찾을 수 없습니다: {base_url}")
        return False

    cookies = load_cookies()
    if not cookies:
        print(f"[오류] 'setting/lcdms_cookie.txt' 쿠키 파일이 비어있거나 없습니다.")
        return False

    os.makedirs(INPUT_DIR, exist_ok=True)
    session = init_session(cookies)

    print("\n" + "="*60)
    print(f"[*] LCDMS 원고/SB 다운로더 (Read-Only 모드)")
    print(f"[*] 과목명: {subject} (콘텐츠번호: {contents_seqno})")
    print(f"[*] 대상 주차: {weeks}")
    print(f"[*] 보안 원칙: 조회 및 다운로드 외 일체 쓰기/변경 금지")
    print("="*60)

    # 1. 진행 단계 메타데이터 조회 API 호출
    api_url = "https://home.iscu.ac.kr/cdms/progress/devstatu/devstatuDetailS.scu"
    try:
        resp = session.post(api_url, data={"contents_seqno": contents_seqno}, timeout=15)
        meta_data = resp.json()
        result_list = meta_data.get("resultList", [])
    except Exception as e:
        print(f"[오류] 진행 메타데이터 조회 실패: {e}")
        return False

    # 주차별 매핑 딕셔너리 구축
    week_map = {}
    for row in result_list:
        w = row.get("seq") or row.get("weekNo")
        if w is not None:
            try:
                week_map[int(w)] = row
            except ValueError:
                pass

    total_downloaded = 0

    for week in weeks:
        print(f"\n>> [{week}주차 점검 시작]")
        if week not in week_map:
            print(f"[-] [{week}주차] 시스템에 등록된 주차 데이터가 없습니다.")
            continue

        w_info = week_map[week]
        ws_seqno = w_info.get("wsSeqno")

        # 단계 탐색: 
        # 원고: 0401 (단건파일)
        # SB: 0403 (SB최초) 또는 0416 (SB최종)
        stages_to_check = [
            {"label": "원고", "codes": ["0401"], "mod": "dev"},
            {"label": "SB", "codes": ["0403", "0416"], "mod": "contents"},
        ]

        found_in_week = 0

        for target in stages_to_check:
            label = target["label"]
            codes = target["codes"]
            mod = target["mod"]

            # 해당되는 단계 찾기
            stage_num = None
            stage_seqno = None
            stage_name = None

            for i in range(1, 26):
                cd = w_info.get(f"s{i}WsStageCd")
                if cd in codes:
                    stage_num = w_info.get(f"s{i}StageNum") or i
                    stage_seqno = w_info.get(f"s{i}Seqno")
                    stage_name = w_info.get(f"s{i}WsStageCdNm", label)
                    break

            if not stage_num or not stage_seqno:
                print(f"[-] [{week}주차 {label}] 해당 단계 정보가 정의되어 있지 않습니다.")
                continue

            # 상세 조회 페이지(devstatuStepCommonDetail.scu) GET 요청
            detail_url = (
                f"https://home.iscu.ac.kr/cdms/progress/devstatu/devstatuStepCommonDetail.scu"
                f"?contents_seqno={contents_seqno}&ws_seqno={ws_seqno}&stage_num={stage_num}&seqno={stage_seqno}&weekNo={week}"
            )

            try:
                r_det = session.get(detail_url, timeout=15)
                r_det.encoding = "utf-8"

                # fileDownload(...) 함수 인자 추출
                # 예: fileDownload('원파일명', '다운파일명', '경로', '구분', '아이디', '번호')
                m = re.search(r"fileDownload\(\s*'([^']*)'\s*,\s*'([^']*)'\s*,\s*'([^']*)'", r_det.text)
                if not m:
                    print(f"[-] [{week}주차 {label}] 등록된 첨부파일이 없습니다.")
                    continue

                org_file = m.group(1).strip()
                down_file = m.group(2).strip()
                attach_path = m.group(3).strip()

                if not org_file or not down_file:
                    print(f"[-] [{week}주차 {label}] 첨부파일 정보를 찾을 수 없습니다.")
                    continue

                # 파일 다운로드 URL 구성 (userDown.scu - Read-Only GET)
                down_url = (
                    f"https://home.iscu.ac.kr/cdms/userDown.scu"
                    f"?cmd={mod}&filepath={attach_path}&filename={js_escape_encode(down_file)}&orgfilename={js_escape_encode(org_file)}"
                )

                print(f"[+] [{week}주차 {label}] 파일 발견: {org_file}")
                print(f"    -> 다운로드 진행 중...")

                r_file = session.get(down_url, stream=True, timeout=60)
                if r_file.status_code == 200:
                    out_path = os.path.join(INPUT_DIR, org_file)
                    with open(out_path, "wb") as f_out:
                        for chunk in r_file.iter_content(chunk_size=16384):
                            f_out.write(chunk)
                    fsize = os.path.getsize(out_path)
                    print(f"    -> [다운로드 완료] {out_path} ({fsize:,} bytes)")
                    found_in_week += 1
                    total_downloaded += 1
                else:
                    print(f"[!] [{week}주차 {label}] 다운로드 실패 (상태코드: {r_file.status_code})")

            except Exception as e:
                print(f"[!] [{week}주차 {label}] 파일 조회 중 오류 발생: {e}")

        if found_in_week == 0:
            print(f"[-] [{week}주차] 등록된 원고 및 SB 파일이 전혀 없습니다.")

    print("\n" + "="*60)
    print(f"[*] 다운로드 작업 완료: 총 {total_downloaded}개 파일 저장됨 (저장위치: {INPUT_DIR})")
    print("="*60)
    return True

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("사용법: python lcdms_downloader.py <과목명> <주차1> [주차2] ...")
        print("예시: python lcdms_downloader.py 고양이관리학 9 10")
    else:
        subj = sys.argv[1]
        w_list = [int(w.replace("주차", "").replace("주", "")) for w in sys.argv[2:]]
        fetch_and_download(subj, w_list)
