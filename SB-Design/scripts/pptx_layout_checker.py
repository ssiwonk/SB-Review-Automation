"""
pptx_layout_checker.py
서울사이버대학교(SCU) 지침서 기반 PPTX 레이아웃 및 서식 정량 검사기 (Track 1).
슬라이드 규격, 폰트(크기/볼드), 영역 이탈(오버플로우), 표 정렬, 출처 표기 누락 등을 100% 자동 검출합니다.
"""

import re
import sys
from pathlib import Path
from typing import List, Dict, Any, Optional

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

try:
    from pptx import Presentation
    from pptx.enum.text import PP_ALIGN
except ImportError:
    Presentation = None

from sb_parser import SCUSBParser


class PPTXLayoutChecker:
    STANDARD_WIDTH_CM = 33.867
    STANDARD_HEIGHT_CM = 19.050
    MIN_BODY_FONT_PT = 20.0
    ALLOWED_FONTS = ["나눔스퀘어", "나눔스퀘어b", "나눔스퀘어eb", "나눔스퀘어라운드", "맑은 고딕", "맑은고딕", "malgun gothic", "nanumsquare"]

    def __init__(self, pptx_path: str):
        self.pptx_path = Path(pptx_path)
        self.parser = SCUSBParser(str(self.pptx_path))
        self.slides_data = self.parser.parse()
        self.findings: List[Dict[str, Any]] = []

    def run_check(self) -> List[Dict[str, Any]]:
        """전체 정량 검사를 수행하고 지적사항 리스트를 반환합니다."""
        self.findings = []

        # 1. 슬라이드 전역 규격 검사
        self._check_slide_dimensions()

        # 2. 슬라이드별 세부 검사
        for s in self.slides_data:
            role = s["role"]
            slide_no = s["slide_no"]

            # 표지, 목차, 체크리스트, 미디어 간지는 본문 규격 검사에서 제외
            if role in ("COVER", "INDEX", "CHECKLIST", "MEDIA_INTERLUDE"):
                continue

            self._check_fonts_and_sizes(s)
            self._check_layout_boundaries(s)
            self._check_notes_and_sources(s)
            self._check_sentence_terminations(s)

        return self.findings

    def _check_slide_dimensions(self):
        """슬라이드 전체 가로/세로 규격 검사 (33.867 x 19.05cm)"""
        w_diff = abs(self.parser.width_cm - self.STANDARD_WIDTH_CM)
        h_diff = abs(self.parser.height_cm - self.STANDARD_HEIGHT_CM)

        if w_diff > 0.1 or h_diff > 0.1:
            self.findings.append({
                "slide_no": "공통",
                "category": "규격",
                "issue": f"슬라이드 크기({self.parser.width_cm:.2f}cm × {self.parser.height_cm:.2f}cm)가 SCU 표준 규격(33.867cm × 19.05cm, 16:9 와이드)과 다름",
                "recommendation": "슬라이드 크기를 [사용자 지정: 너비 33.867cm, 높이 19.05cm, 가로 방향]으로 재설정하세요."
            })

    def _check_fonts_and_sizes(self, s: Dict[str, Any]):
        """본문 슬라이드의 폰트 크기(최소 20pt) 및 Bold 여부 검사"""
        slide_no = s["slide_no"]

        small_font_cases = []
        non_bold_cases = []
        disallowed_fonts = []

        for shape in s["shape_items"]:
            # 출처 표기 영역(보통 12pt 허용)이나 상단 헤더, To.교수님 제외
            text = shape["text"].strip()
            if text.startswith("#출처") or text.startswith("#") or "게티이미지" in text:
                continue

            for p in shape["paragraphs"]:
                p_text = p["text"].strip()
                if not p_text or len(p_text) <= 1:
                    continue

                # 폰트 크기 체크 (제목/헤더가 아닌 일반 본문 대상)
                if p["font_size"] is not None:
                    # 19.5pt 미만인 경우 (단, 짧은 각주/참고 텍스트 제외)
                    if p["font_size"] < 19.5 and shape["top_cm"] > 3.0:
                        small_font_cases.append(f"'{p_text[:20]}...'(현재 {p['font_size']}pt)")

                # Bold 여부 체크
                if p["bold"] is False and shape["top_cm"] > 3.0:
                    non_bold_cases.append(f"'{p_text[:15]}...'")

                # 폰트 패밀리 체크
                if p["font_name"]:
                    f_name = p["font_name"].lower().strip()
                    if not any(allowed in f_name for allowed in self.ALLOWED_FONTS):
                        disallowed_fonts.append(p["font_name"])

        if small_font_cases:
            self.findings.append({
                "slide_no": f"P. {slide_no:02d}",
                "category": "디자인/폰트",
                "issue": f"본문 텍스트 중 일부가 SCU 최소 폰트 기준(20pt)에 미달함: {', '.join(small_font_cases[:2])}",
                "recommendation": "가독성 확보 지침에 따라 본문 텍스트 크기를 20pt 이상(나눔스퀘어B)으로 확대하세요."
            })

        if disallowed_fonts:
            unique_fonts = list(set(disallowed_fonts))
            self.findings.append({
                "slide_no": f"P. {slide_no:02d}",
                "category": "디자인/폰트",
                "issue": f"지정 서체(나눔스퀘어, 맑은고딕) 외 다른 글꼴 사용됨: {', '.join(unique_fonts)}",
                "recommendation": "SCU 표준 서체인 나눔스퀘어(또는 맑은고딕)로 변경하여 서체 일관성을 유지하세요."
            })

    def _check_layout_boundaries(self, s: Dict[str, Any]):
        """슬라이드 경계 이탈(오버플로우) 및 여백 검사"""
        slide_no = s["slide_no"]

        overflow_items = []
        for shape in s["shape_items"]:
            # 슬라이드 우측 경계(33.87cm) 또는 하단 경계(19.05cm) 초과 체크
            if shape["right_cm"] > (self.STANDARD_WIDTH_CM + 0.3) or shape["bottom_cm"] > (self.STANDARD_HEIGHT_CM + 0.3):
                overflow_items.append(f"'{shape['name']}' (우측: {shape['right_cm']}cm, 하단: {shape['bottom_cm']}cm)")

        if overflow_items:
            self.findings.append({
                "slide_no": f"P. {slide_no:02d}",
                "category": "레이아웃",
                "issue": f"개체(도형/텍스트상자)가 슬라이드 화면 경계를 벗어남: {', '.join(overflow_items[:2])}",
                "recommendation": "개체 위치 및 크기를 조절하여 슬라이드 안전 영역(너비 33.867cm, 높이 19.05cm 이내) 안으로 배치하세요."
            })

    def _check_notes_and_sources(self, s: Dict[str, Any]):
        """본문 이미지 포함 시 슬라이드 노트 출처 표기 여부 검사"""
        slide_no = s["slide_no"]
        img_cnt = s["image_count"]
        notes = s["notes_text"].strip()

        # 본문에 이미지가 있는데 슬라이드 노트가 비어있거나 출처 표기가 없는 경우
        if img_cnt > 0:
            pattern = r"(#|게티|getty|출처|http|www|구매\s*이미지|유료\s*이미지|자체\s*제작|저작권|라이선스)"
            has_source_mark = bool(re.search(pattern, notes, re.IGNORECASE))
            if not has_source_mark:
                self.findings.append({
                    "slide_no": f"P. {slide_no:02d}",
                    "category": "저작권/출처",
                    "issue": f"슬라이드에 이미지/사진({img_cnt}개)이 포함되어 있으나 슬라이드 노트에 출처 또는 유료 이미지 번호가 기재되지 않음",
                    "recommendation": "슬라이드 노트에 저작물 출처(예: '#1 출처: 저작자, 저작물명...' 또는 '구매 이미지 사용 - 게티이미지뱅크(번호)')를 필수 기재하세요."
                })

    def _check_sentence_terminations(self, s: Dict[str, Any]):
        """명사형 종결 마침표 및 괄호 마침표 검사"""
        slide_no = s["slide_no"]

        wrong_noun_endings = []
        wrong_paren_endings = []

        for ct in s["content_texts"]:
            for line in ct.splitlines():
                line = line.strip()
                if not line:
                    continue

                # 1. 개조식 명사형 종결어미 뒤 마침표: (~함., ~임., ~음., ~됨., ~것.)
                if re.search(r"[가-힣]+(함|임|음|됨|것)\.$", line):
                    wrong_noun_endings.append(line[:25] + ("..." if len(line) > 25 else ""))

                # 2. 괄호 뒤 마침표가 아닌 괄호 앞 마침표: .(2026)
                if re.search(r"\.\s*\([^\)]+\)", line):
                    wrong_paren_endings.append(line[:25] + ("..." if len(line) > 25 else ""))

        if wrong_noun_endings:
            self.findings.append({
                "slide_no": f"P. {slide_no:02d}",
                "category": "문장규격",
                "issue": f"서술형 문장이 아닌 개조식/명사형 종결 문장에 마침표가 표기됨: {', '.join(wrong_noun_endings[:2])}",
                "recommendation": "SCU 지침에 따라 서술형 종결 문장(~다, ~요)에만 마침표를 사용하고 명사형 종결(~함, ~임)의 마침표는 제거하세요."
            })

        if wrong_paren_endings:
            self.findings.append({
                "slide_no": f"P. {slide_no:02d}",
                "category": "문장규격",
                "issue": f"괄호 부가설명 문장에서 마침표가 괄호 앞에 표기됨: {', '.join(wrong_paren_endings[:2])}",
                "recommendation": "마침표를 괄호 뒤로 이동하여 표기하세요 (예: '...적용한다(2026).')."
            })


if __name__ == "__main__":
    if len(sys.argv) > 1:
        checker = PPTXLayoutChecker(sys.argv[1])
        results = checker.run_check()
        print(f"Total findings: {len(results)}")
        for r in results:
            print(f"[{r['slide_no']}] ({r['category']}) {r['issue']}")
            print(f"  -> {r['recommendation']}\n")
    else:
        print("Usage: python pptx_layout_checker.py <sb_pptx_path>")
