"""
doc_parser.py
다양한 형식의 원고(HWP, HWPX, DOCX, PDF, PPTX)에서 텍스트와 구조를 추출하는 범용 파서 모듈.
외부 오피스 프로그램 종속성 없이 순수 파이썬 라이브러리(olefile, zipfile, docx, pdfplumber, pptx)로 동작합니다.
"""

import os
import sys
import zlib
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List, Dict, Any, Optional

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

try:
    import olefile
except ImportError:
    olefile = None

try:
    import docx
except ImportError:
    docx = None

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    from pptx import Presentation
except ImportError:
    Presentation = None


def parse_hwp(file_path: str) -> str:
    """바이너리 HWP 5.0 파일에서 본문 텍스트를 추출합니다."""
    if olefile is None:
        raise ImportError("olefile 패키지가 필요합니다: pip install olefile")

    if not olefile.isOleFile(file_path):
        # 만약 확장자만 hwp이고 실제로는 hwpx(zip)인 경우 대비
        if zipfile.is_zipfile(file_path):
            return parse_hwpx(file_path)
        raise ValueError(f"올바른 HWP 파일 형식이 아닙니다: {file_path}")

    ole = olefile.OleFileIO(file_path)
    file_dirs = ole.listdir()

    # Section 스트림 목록 검색
    sections = []
    for entry in file_dirs:
        if entry[0] == "BodyText" and entry[1].startswith("Section"):
            sections.append(entry)

    # 섹션 번호 순으로 정렬
    sections.sort(key=lambda x: int(x[1].replace("Section", "")) if x[1].replace("Section", "").isdigit() else 0)

    extracted_text = []
    for section in sections:
        stream = ole.openstream(section)
        data = stream.read()
        
        # HWP 스트림은 일반적으로 zlib 압축되어 있음 (Header의 압축 플래그와 무관하게 대부분 zlib -15 모드)
        try:
            decompressed = zlib.decompress(data, -15)
        except Exception:
            try:
                decompressed = zlib.decompress(data)
            except Exception:
                decompressed = data

        # UTF-16LE 텍스트 디코딩 및 제어 문자 필터링
        text_chunk = []
        i = 0
        n = len(decompressed)
        while i < n - 1:
            code = decompressed[i] | (decompressed[i + 1] << 8)
            # 일반 문자 및 개행(10, 13, 9)
            if code in (10, 13, 9) or (0x20 <= code <= 0xD7FF) or (0xE000 <= code <= 0xFFFD):
                text_chunk.append(chr(code))
            elif code == 0:
                pass
            i += 2

        raw_str = "".join(text_chunk)
        # 제어문자 잔여물 정제
        cleaned = "\n".join([line.strip() for line in raw_str.splitlines() if line.strip()])
        if cleaned:
            extracted_text.append(cleaned)

    ole.close()
    return "\n\n".join(extracted_text)


def parse_hwpx(file_path: str) -> str:
    """XML 기반 HWPX 파일에서 본문 텍스트를 추출합니다."""
    with zipfile.ZipFile(file_path, "r") as z:
        section_names = [name for name in z.namelist() if name.startswith("Contents/section") and name.endswith(".xml")]
        section_names.sort()

        extracted_text = []
        for name in section_names:
            xml_data = z.read(name)
            root = ET.fromstring(xml_data)
            # 모든 텍스트 노드 추출 (hp:t 태그 등)
            texts = []
            for elem in root.iter():
                if elem.tag.endswith("}t") or elem.tag == "t":
                    if elem.text:
                        texts.append(elem.text)
                elif elem.tag.endswith("}p") or elem.tag == "p":
                    texts.append("\n")
            
            section_str = "".join(texts)
            cleaned = "\n".join([l.strip() for l in section_str.splitlines() if l.strip()])
            if cleaned:
                extracted_text.append(cleaned)

        return "\n\n".join(extracted_text)


def parse_docx(file_path: str) -> str:
    """DOCX 파일에서 단락 및 표 텍스트를 추출합니다."""
    if docx is None:
        raise ImportError("python-docx 패키지가 필요합니다: pip install python-docx")

    doc = docx.Document(file_path)
    output = []

    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt:
            output.append(txt)

    for table in doc.tables:
        table_lines = []
        for row in table.rows:
            row_txt = [c.text.strip().replace("\n", " ") for c in row.cells]
            table_lines.append(" | ".join(row_txt))
        if table_lines:
            output.append("[표 데이터]\n" + "\n".join(table_lines))

    return "\n\n".join(output)


def parse_pdf(file_path: str) -> str:
    """PDF 파일에서 텍스트 및 표를 추출합니다."""
    if pdfplumber is None:
        raise ImportError("pdfplumber 패키지가 필요합니다: pip install pdfplumber")

    output = []
    with pdfplumber.open(file_path) as pdf:
        for idx, page in enumerate(pdf.pages):
            p_text = page.extract_text()
            if p_text and p_text.strip():
                output.append(f"--- [PDF Page {idx + 1}] ---\n{p_text.strip()}")

    return "\n\n".join(output)


def parse_pptx(file_path: str) -> str:
    """PPTX 원고 파일에서 슬라이드별 텍스트 및 노트를 추출합니다."""
    if Presentation is None:
        raise ImportError("python-pptx 패키지가 필요합니다: pip install python-pptx")

    prs = Presentation(file_path)
    output = []

    for idx, slide in enumerate(prs.slides):
        slide_lines = [f"=== [슬라이드 {idx + 1}] ==="]
        
        # 슬라이드 노트
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
            slide_lines.append(f"[노트]: {slide.notes_slide.notes_text_frame.text.strip()}")

        # 셰이프 텍스트
        for shape in slide.shapes:
            if shape.has_text_frame:
                txt = shape.text_frame.text.strip()
                if txt:
                    slide_lines.append(txt)
            elif shape.has_table:
                for row in shape.table.rows:
                    r_text = [c.text.strip().replace("\n", " ") for c in row.cells]
                    slide_lines.append(" | ".join(r_text))

        if len(slide_lines) > 1:
            output.append("\n".join(slide_lines))

    return "\n\n".join(output)


def extract_manuscript_text(file_path: str) -> Dict[str, Any]:
    """
    확장자를 감지하여 해당 원고에서 텍스트를 추출하는 통합 함수.
    반환값:
        {
            "file_name": 파일명,
            "extension": 확장자,
            "text": 추출된 전체 텍스트,
            "char_count": 글자 수,
            "status": "success" | "error",
            "error_msg": 에러 메시지 (실패 시)
        }
    """
    path = Path(file_path)
    ext = path.suffix.lower()

    result = {
        "file_path": str(path.resolve()),
        "file_name": path.name,
        "extension": ext,
        "text": "",
        "char_count": 0,
        "status": "success",
        "error_msg": None
    }

    try:
        if ext == ".hwp":
            result["text"] = parse_hwp(str(path))
        elif ext == ".hwpx":
            result["text"] = parse_hwpx(str(path))
        elif ext in (".docx", ".doc"):
            result["text"] = parse_docx(str(path))
        elif ext == ".pdf":
            result["text"] = parse_pdf(str(path))
        elif ext in (".pptx", ".ppt"):
            result["text"] = parse_pptx(str(path))
        else:
            raise ValueError(f"지원하지 않는 원고 파일 형식입니다: {ext}")

        result["char_count"] = len(result["text"])

    except Exception as e:
        result["status"] = "error"
        result["error_msg"] = str(e)

    return result


if __name__ == "__main__":
    if len(sys.argv) > 1:
        target = sys.argv[1]
        print(f"Parsing: {target}")
        res = extract_manuscript_text(target)
        if res["status"] == "success":
            print(f"Success! Characters: {res['char_count']}")
            print("--- Preview (first 500 chars) ---")
            print(res["text"][:500])
        else:
            print(f"Error: {res['error_msg']}")
    else:
        print("Usage: python doc_parser.py <file_path>")
