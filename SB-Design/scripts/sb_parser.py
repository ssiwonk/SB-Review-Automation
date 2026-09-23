"""
sb_parser.py
서울사이버대학교(SCU) 표준 스토리보드(pptx)를 정밀 분석하여
슬라이드 역할 분류, 페이지 코드, 헤더/본문 텍스트, 설계자 메모, 슬라이드 노트(출처)를 구조화하는 파서 모듈.
"""

import sys
import re
from pathlib import Path
from typing import List, Dict, Any, Optional

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

try:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
except ImportError:
    Presentation = None

CIRCLED = ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩", "⑪", "⑫", "⑬", "⑭", "⑮", "⑯", "⑰", "⑱", "⑲", "⑳"]
BLACK_CIRCLED = ["➊", "➋", "➌", "➍", "➎", "➏", "➐", "➑", "➒", "➓"]

# Wingdings 등 심볼 폰트의 글머리 기호 문자 -> 표준 유니코드 기호 매핑
WINGDINGS_BULLET_MAP = {
    'l': '●',  # Wingdings 108: 원형 불릿
    'n': '■',  # Wingdings 110: 사각형 불릿
    'u': '◆',  # Wingdings 117: 다이아몬드 불릿
    '§': '▪',  # Wingdings 167: 작은 사각형 불릿
    '–': '–',  # Wingdings 8211: 대시 불릿
    'q': '❏',
    'x': '✗',
    'y': '✓',
    'z': '✔',
}

def format_autonum(num_type: str, n: int) -> str:
    if num_type in ["circleNumDbPlain", "circleNumWdWhitePlain"]:
        if 1 <= n <= len(CIRCLED):
            return CIRCLED[n - 1]
        return f"({n})"
    elif num_type == "circleNumWdBlackPlain":
        if 1 <= n <= len(BLACK_CIRCLED):
            return BLACK_CIRCLED[n - 1]
        return f"({n})"
    elif num_type == "arabicPeriod":
        return f"{n}."
    elif num_type == "arabicParenR":
        return f"{n})"
    elif num_type == "arabicParenBoth":
        return f"({n})"
    elif num_type == "alphaLcParenR":
        ch = chr(ord('a') + (n - 1) % 26)
        return f"{ch})"
    elif num_type == "alphaUcParenR":
        ch = chr(ord('A') + (n - 1) % 26)
        return f"{ch})"
    elif num_type == "alphaLcPeriod":
        ch = chr(ord('a') + (n - 1) % 26)
        return f"{ch}."
    elif num_type == "alphaUcPeriod":
        ch = chr(ord('A') + (n - 1) % 26)
        return f"{ch}."
    return f"{n}."

def extract_paragraphs_with_bullets(text_frame) -> List[Dict[str, Any]]:
    """
    text_frame 내부의 paragraph들을 순회하며 OpenXML buAutoNum 및 buChar를 감지하여
    앞에 기호/번호를 복원한 텍스트 정보를 반환합니다.
    """
    results = []
    counters = {}

    for p in text_frame.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        
        xml = getattr(p._p, "xml", "")
        prefix = ""

        # buAutoNum 확인
        m_auto = re.search(r'<a:buAutoNum[^>]*type="([^"]+)"(?:\s+startAt="(\d+)")?', xml)
        if not m_auto:
            m_type = re.search(r'buAutoNum[^>]*type="([^"]+)"', xml)
            m_start = re.search(r'buAutoNum[^>]*startAt="(\d+)"', xml)
            if m_type:
                num_type = m_type.group(1)
                start_at = int(m_start.group(1)) if m_start else None
            else:
                num_type = None
                start_at = None
        else:
            num_type = m_auto.group(1)
            start_at = int(m_auto.group(2)) if m_auto.group(2) else None

        if num_type:
            lvl = p.level if hasattr(p, 'level') else 0
            if start_at is not None:
                counters[lvl] = start_at
            else:
                counters[lvl] = counters.get(lvl, 0) + 1
            prefix = format_autonum(num_type, counters[lvl]) + " "
        else:
            m_char = re.search(r'buChar[^>]*char="([^"]+)"', xml)
            if m_char:
                ch = m_char.group(1)
                m_font = re.search(r'buFont[^>]*typeface="([^"]+)"', xml)
                if m_font and "wingdings" in m_font.group(1).lower():
                    ch = WINGDINGS_BULLET_MAP.get(ch, "•")
                prefix = ch + " "

        full_text = prefix + t
        results.append({
            "paragraph": p,
            "text": full_text,
            "raw_text": t,
            "prefix": prefix
        })
    return results


class SCUSBParser:
    def __init__(self, pptx_path: str):
        if Presentation is None:
            raise ImportError("python-pptx 패키지가 필요합니다: pip install python-pptx")
        
        self.pptx_path = Path(pptx_path)
        self.prs = Presentation(str(self.pptx_path))
        self.width_cm = round(self.prs.slide_width.cm, 3)
        self.height_cm = round(self.prs.slide_height.cm, 3)
        self.slides_data: List[Dict[str, Any]] = []

    def parse(self, rendered_images: Optional[Dict[int, str]] = None) -> List[Dict[str, Any]]:
        """모든 슬라이드를 순회하며 구조화된 정보를 추출합니다."""
        self.slides_data = []

        for idx, slide in enumerate(self.prs.slides):
            slide_no = idx + 1
            layout_name = slide.slide_layout.name if slide.slide_layout else ""
            
            # 1. 슬라이드 역할(Role) 판별
            role = self._classify_role(slide_no, layout_name, slide)

            # 2. 슬라이드 노트(Notes) 추출 (출처 정보)
            notes_text = ""
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                notes_text = slide.notes_slide.notes_text_frame.text.strip()

            # 3. 셰이프별 정보 분리
            page_code = None
            unit_title = None
            sub_title = None
            content_texts = []
            designer_memos = []
            screen_descriptions = []
            tables_data = []
            image_count = 0
            shape_items = []

            for shape, rx, ry, rw, rh in self._iter_shapes(slide.shapes):
                # 이미지 개수 체크 (통상 아이콘/픽토그램 등 3.5cm 이하 소형 이미지는 출처 면제)
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    w_cm = round(rw / 914400 * 2.54, 2) if rw is not None else 0.0
                    h_cm = round(rh / 914400 * 2.54, 2) if rh is not None else 0.0
                    if w_cm > 3.5 or h_cm > 3.5:
                        image_count += 1

                # 표 데이터 추출 (셀 위 오버레이 도형 자동 감지 포함)
                if shape.has_table:
                    table_rows = self._extract_table_data(slide, shape)
                    tables_data.append(table_rows)

                # 텍스트 프레임 처리
                if shape.has_text_frame and shape.text_frame:
                    p_items = extract_paragraphs_with_bullets(shape.text_frame)
                    raw_text = "\n".join([item["text"] for item in p_items]).strip()
                    if not raw_text:
                        continue

                    left_cm = round(rx / 914400 * 2.54, 2) if rx is not None else 0.0
                    top_cm = round(ry / 914400 * 2.54, 2) if ry is not None else 0.0
                    width_cm = round(rw / 914400 * 2.54, 2) if rw is not None else 0.0
                    height_cm = round(rh / 914400 * 2.54, 2) if rh is not None else 0.0

                    # (1) 설계자 메모 감지 (To. 교수님 등)
                    if "To." in raw_text or "to." in raw_text or "교수님" in raw_text and ("확인" in raw_text or "수정" in raw_text):
                        designer_memos.append({
                            "text": raw_text,
                            "pos": (left_cm, top_cm, width_cm, height_cm)
                        })
                        continue

                    # (2) 페이지 코드 감지 (예: 01_02_01 등 패턴 또는 우측 상단/상단 소형 박스)
                    if self._is_page_code(raw_text):
                        page_code = raw_text

                    # (2-1) 화면설명 영역 감지 (우측 화면설명 영역: left_cm >= 28.0 and top_cm >= 1.5, 검수 대상에서 완전 제외)
                    elif left_cm >= 28.0 and top_cm >= 1.5:
                        screen_descriptions.append({
                            "text": raw_text,
                            "pos": (left_cm, top_cm, width_cm, height_cm)
                        })
                        continue

                    # (3) 상단 유닛 타이틀 / 부제목 감지 (헤더 영역: top < 3.5cm)
                    elif top_cm < 2.0 and height_cm <= 2.0 and not unit_title and len(raw_text) < 40:
                        unit_title = raw_text
                    elif 1.8 <= top_cm < 4.0 and not sub_title and len(raw_text) < 50:
                        sub_title = raw_text
                    else:
                        content_texts.append(raw_text)

                    # 폰트 검사용 셰이프 상세 정보 기록
                    shape_info = {
                        "name": shape.name,
                        "shape_type": str(shape.shape_type),
                        "left_cm": left_cm,
                        "top_cm": top_cm,
                        "width_cm": width_cm,
                        "height_cm": height_cm,
                        "right_cm": round(left_cm + width_cm, 2),
                        "bottom_cm": round(top_cm + height_cm, 2),
                        "text": raw_text,
                        "paragraphs": []
                    }

                    for p_item in p_items:
                        p = p_item["paragraph"]
                        p_info = {
                            "text": p_item["text"],
                            "font_name": p.font.name if p.font else None,
                            "font_size": p.font.size.pt if (p.font and p.font.size) else None,
                            "bold": p.font.bold if p.font else None,
                            "color": str(p.font.color.rgb) if (p.font and p.font.color and hasattr(p.font.color, 'rgb') and p.font.color.rgb) else None
                        }
                        shape_info["paragraphs"].append(p_info)

                    shape_items.append(shape_info)

            slide_entry = {
                "slide_no": slide_no,
                "role": role,
                "layout_name": layout_name,
                "page_code": page_code,
                "unit_title": unit_title,
                "sub_title": sub_title,
                "content_texts": content_texts,
                "designer_memos": designer_memos,
                "screen_descriptions": screen_descriptions,
                "tables_data": tables_data,
                "image_count": image_count,
                "notes_text": notes_text,
                "shape_items": shape_items,
                "rendered_image_path": rendered_images.get(slide_no) if rendered_images else None
            }

            self.slides_data.append(slide_entry)

        return self.slides_data

    def _extract_table_data(self, slide, table_shape) -> List[List[str]]:
        """표 데이터를 추출하며, 빈 셀 위에 얹혀진 오버레이 도형(예: 화살표, 텍스트박스)을 기하학적으로 감지합니다."""
        t = table_shape.table
        col_widths = [col.width for col in t.columns]
        table_rows = []

        for r_idx, row in enumerate(t.rows):
            row_vals = []
            for c_idx, cell in enumerate(row.cells):
                paras = extract_paragraphs_with_bullets(cell.text_frame)
                if paras:
                    val = " <br> ".join([p["text"].replace("\n", " ").strip() for p in paras]).strip()
                else:
                    val = ""
                # 본문 데이터 셀이 공란인 경우 해당 열/행 영역에 오버레이된 도형 탐색
                if not val and r_idx > 0:
                    c_left = table_shape.left + sum(col_widths[:c_idx])
                    c_w = col_widths[c_idx]
                    overlay_val = self._find_cell_overlay(slide, table_shape, c_left, c_w)
                    if overlay_val:
                        val = overlay_val
                row_vals.append(val)
            table_rows.append(row_vals)
        return table_rows

    def _find_cell_overlay(self, slide, table_shape, c_left, c_w) -> str:
        """표의 비어있는 셀 좌표와 기하학적으로 겹치는 별도 도형 또는 그룹의 텍스트를 탐색합니다."""
        for sh in slide.shapes:
            if sh == table_shape:
                continue
            sh_cx = sh.left + sh.width / 2
            sh_cy = sh.top + sh.height / 2
            # 가로 범위가 해당 열에 속하고, 세로 범위가 표 내부 영역에 속하는지 검사
            if (c_left - 180000) <= sh_cx <= (c_left + c_w + 180000):
                if (table_shape.top - 180000) <= sh_cy <= (table_shape.top + table_shape.height + 180000):
                    if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
                        texts = []
                        for sub in sorted(sh.shapes, key=lambda s: s.left):
                            if sub.has_text_frame and sub.text_frame.text.strip():
                                texts.append(sub.text_frame.text.strip())
                            elif any(k in sub.name.lower() for k in ["화살표", "그래픽", "arrow"]):
                                texts.append("→")
                        if texts:
                            return " ".join(texts)
                    elif sh.has_text_frame and sh.text_frame.text.strip():
                        return sh.text_frame.text.strip()
        return ""

    def _get_real_coords(self, shape, parent_transform=None):
        left = shape.left
        top = shape.top
        w = shape.width
        h = shape.height
        
        if parent_transform:
            off_x, off_y, ext_cx, ext_cy, chOff_x, chOff_y, chExt_cx, chExt_cy = parent_transform
            if chExt_cx and chExt_cy and left is not None and top is not None:
                real_x = off_x + (left - chOff_x) * ext_cx / chExt_cx
                real_y = off_y + (top - chOff_y) * ext_cy / chExt_cy
                real_w = w * ext_cx / chExt_cx if w is not None else 0
                real_h = h * ext_cy / chExt_cy if h is not None else 0
                return real_x, real_y, real_w, real_h
        return left, top, w, h

    def _iter_shapes(self, shapes, parent_transform=None):
        """그룹 도형을 포함하여 모든 하위 셰이프를 재귀적으로 평탄화(flatten)하고 실제 슬라이드 좌표를 계산하여 반환합니다."""
        for shape in shapes:
            if hasattr(shape, "shapes") and shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                grpSpPr = getattr(shape._element, "grpSpPr", None)
                xfrm = getattr(grpSpPr, "xfrm", None) if grpSpPr is not None else None
                if xfrm is not None and hasattr(xfrm, 'chOff') and hasattr(xfrm, 'chExt') and xfrm.chExt.cx and xfrm.chExt.cy:
                    if parent_transform:
                        cur_off_x, cur_off_y, cur_ext_cx, cur_ext_cy = self._get_real_coords(shape, parent_transform)
                    else:
                        cur_off_x, cur_off_y, cur_ext_cx, cur_ext_cy = shape.left, shape.top, shape.width, shape.height
                    
                    new_transform = (cur_off_x, cur_off_y, cur_ext_cx, cur_ext_cy, xfrm.chOff.x, xfrm.chOff.y, xfrm.chExt.cx, xfrm.chExt.cy)
                    for sub_item in self._iter_shapes(shape.shapes, new_transform):
                        yield sub_item
                else:
                    for sub_item in self._iter_shapes(shape.shapes, parent_transform):
                        yield sub_item
            else:
                rx, ry, rw, rh = self._get_real_coords(shape, parent_transform)
                yield shape, rx, ry, rw, rh

    def _classify_role(self, slide_no: int, layout_name: str, slide) -> str:
        """슬라이드의 역할을 분류합니다."""
        if slide_no == 1 or "표지" in layout_name:
            return "COVER"
        if slide_no == 2 or "목차" in layout_name and "1_" not in layout_name:
            return "INDEX"
        if slide_no == 3 or "체크리스트" in layout_name or "작성안내" in layout_name:
            return "CHECKLIST"
        if "제목 슬라이드" in layout_name or "간지" in layout_name:
            return "MEDIA_INTERLUDE"
        
        # 슬라이드 내 고정 URL이나 mp4 확장자만 있는 경우도 미디어 간지로 분류
        all_text = " ".join([item[0].text_frame.text for item in self._iter_shapes(slide.shapes) if item[0].has_text_frame])
        if "doRandomIntroMedia" in all_text or (len(all_text.strip().split()) <= 3 and ".mp4" in all_text):
            return "MEDIA_INTERLUDE"

        return "CONTENT"

    def _is_page_code(self, text: str) -> bool:
        """'01_02_01' 같은 형식의 페이지 식별 코드인지 검사합니다."""
        parts = text.strip().split("_")
        if len(parts) == 3 and all(p.isdigit() for p in parts):
            return True
        return False

    def get_structured_summary_for_llm(self) -> str:
        """LLM 에이전트(트랙 2)에게 제공할 압축 정제된 SB 마크다운 텍스트를 생성합니다."""
        lines = []
        for s in self.slides_data:
            lines.append(f"\n### [슬라이드 {s['slide_no']}] (구분: {s['role']}, 코드: {s['page_code'] or '없음'})")
            if s["unit_title"]:
                lines.append(f"- **유닛명/대제목**: {s['unit_title']}")
            if s["sub_title"]:
                lines.append(f"- **중/소제목**: {s['sub_title']}")

            if s["content_texts"]:
                lines.append("- **본문 내용**:")
                for ct in s["content_texts"]:
                    for l in ct.splitlines():
                        if l.strip():
                            lines.append(f"  • {l.strip()}")

            if s["tables_data"]:
                lines.append("- **표 데이터**:")
                for t in s["tables_data"]:
                    for row in t:
                        lines.append(f"  | {' | '.join(row)} |")

            if s["notes_text"]:
                lines.append(f"- **슬라이드 노트(출처)**: {s['notes_text']}")

            if s["designer_memos"]:
                lines.append("- **설계자 메모(To.교수님)**:")
                for m in s["designer_memos"]:
                    lines.append(f"  [메모]: {m['text']}")

        return "\n".join(lines)


if __name__ == "__main__":
    if len(sys.argv) > 1:
        parser = SCUSBParser(sys.argv[1])
        data = parser.parse()
        print(f"Parsed {len(data)} slides successfully.")
        summary = parser.get_structured_summary_for_llm()
        print(summary[:1200])
    else:
        print("Usage: python sb_parser.py <sb_pptx_path>")
