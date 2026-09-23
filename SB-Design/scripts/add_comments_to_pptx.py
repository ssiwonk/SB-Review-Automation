"""
add_comments_to_pptx.py
검수 결과 보고서(마크다운)를 파싱하여, 파워포인트(PPTX)의 해당 슬라이드에
정식 '메모(Comment)'로 지적사항 및 수정요청을 자동 삽입하는 모듈.
"""

import os
import sys
import re
import argparse
from pathlib import Path
from collections import defaultdict
from typing import Dict, List, Any, Optional

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def parse_slide_numbers(slide_str: str) -> List[int]:
    """
    마크다운 표의 슬라이드 번호 문자열을 파싱하여 정수 리스트로 반환합니다.
    예: '**P. 01**' -> [1]
        '**P. 05 ~ P. 06**' -> [5, 6]
        '**P. 40-41, P. 44-45, P. 48-49**' -> [40, 41, 44, 45, 48, 49]
        'P. 06, P. 19, P. 25 (총 7개 슬라이드)' -> [6, 19, 25]
    """
    clean = re.sub(r'\(총\s*\d+개[^\)]*\)', '', slide_str)
    clean = clean.replace('*', '').replace('P.', '').replace('P', '').strip()
    slides = set()
    parts = [p.strip() for p in clean.split(',') if p.strip()]

    for part in parts:
        if '~' in part:
            sub = [s.strip() for s in part.split('~')]
            if len(sub) == 2 and sub[0].isdigit() and sub[1].isdigit():
                for n in range(int(sub[0]), int(sub[1]) + 1):
                    slides.add(n)
        elif '-' in part:
            sub = [s.strip() for s in part.split('-')]
            if len(sub) == 2 and sub[0].isdigit() and sub[1].isdigit():
                for n in range(int(sub[0]), int(sub[1]) + 1):
                    slides.add(n)
        else:
            m = re.search(r'\d+', part)
            if m:
                slides.add(int(m.group(0)))

    return sorted(list(slides))


def parse_report_findings(report_path: str) -> Dict[int, List[Dict[str, str]]]:
    """
    마크다운 보고서 파일에서 내용 검수, 레이아웃 검수, 저작권 점검 결과를 슬라이드 번호별로 추출합니다.
    """
    p = Path(report_path)
    if not p.exists():
        raise FileNotFoundError(f"보고서 파일을 찾을 수 없습니다: {p}")

    with open(p, "r", encoding="utf-8") as f:
        text = f.read()

    findings_by_slide = defaultdict(list)

    # 1. 표(Table) 파싱 (내용 검수 및 레이아웃 검수)
    lines = text.splitlines()
    for line in lines:
        line_s = line.strip()
        if line_s.startswith('|') and line_s.endswith('|'):
            cols = [c.strip() for c in line_s.split('|')[1:-1]]
            if len(cols) >= 4 and not cols[0].startswith('---') and not cols[0].startswith(':') and '슬라이드' not in cols[0] and '검수 영역' not in cols[0]:
                slide_col, cat_col, issue_col, req_col = cols[0], cols[1], cols[2], cols[3]
                issue_clean = issue_col.replace('<br>', '\n').replace('**', '').strip()
                req_clean = req_col.replace('<br>', '\n').replace('`', '').replace('**', '').strip()
                sl_nums = parse_slide_numbers(slide_col)
                for sn in sl_nums:
                    findings_by_slide[sn].append({
                        "category": cat_col,
                        "issue": issue_clean,
                        "request": req_clean
                    })

    # 2. 저작권 및 출처 점검 블록 파싱
    m_cr = re.search(r'> \*\*저작권, 출처 미기재\*\*\s*:\s*([^\n]+)', text)
    if m_cr:
        cr_slides_str = m_cr.group(1)
        cr_slides = parse_slide_numbers(cr_slides_str)
        for sn in cr_slides:
            findings_by_slide[sn].append({
                "category": "저작권 및 출처 점검",
                "issue": "본문에 이미지/사진(도형, 일러스트 등)이 포함되어 있으나 슬라이드 노트에 저작물 출처(#1, #출처/N) 또는 게티이미지 라이선스 번호가 누락되었습니다.",
                "request": "영상 촬영 전 슬라이드 노트에 명확한 출처 번호 및 게티이미지뱅크 고유 라이선스 번호를 기재해주세요."
            })

    # 3. 학습정리 AI 추천 완성 문안이 있다면 해당 슬라이드 항목에 자동 부가
    for m in re.finditer(r'```markdown\s*(■\s*\[(?:P\.\s*(\d+))?[^\]]*\].*?)```', text, re.DOTALL):
        block_text = m.group(1).strip()
        target_p = m.group(2)
        if target_p:
            sn = int(target_p)
            if sn in findings_by_slide:
                for item in findings_by_slide[sn]:
                    if "학습정리" in item["category"] or "학습정리" in item["issue"]:
                        item["request"] += f"\n\n[AI 추천 완성 문안]\n{block_text}"
        else:
            # P. 번호가 직접 지정되지 않은 경우(예: 단일 카드 등) 학습정리 항목이 있는 슬라이드에 부가
            for sn, items in findings_by_slide.items():
                for item in items:
                    if "학습정리" in item["category"] or "학습정리" in item["issue"]:
                        if "[AI 추천 완성 문안]" not in item["request"]:
                            item["request"] += f"\n\n[AI 추천 완성 문안]\n{block_text}"

    # 4. 설계자 확인요청 사항 블록 파싱 (To. 교수님 말풍선 등)
    # 보고서 내 '설계자 확인요청 사항' 섹션에서 슬라이드 번호와 요청 내용을 추출하여 '확인요청 사항 있음' 메모로 주입
    memo_section = re.search(r'설계자 확인요청 사항[^\n]*\n(.*?)(?=\n---|\n##|\Z)', text, re.DOTALL)
    if memo_section:
        sec_text = memo_section.group(1)
        for m in re.finditer(r'-\s*\*\*([^*]+)\*\*\s*:\s*([^\n]+)', sec_text):
            slide_str = m.group(1).strip()
            memo_content = m.group(2).strip().strip('`').strip('*').strip()
            sl_nums = parse_slide_numbers(slide_str)
            for sn in sl_nums:
                findings_by_slide[sn].append({
                    "category": "확인요청 사항 있음",
                    "issue": memo_content,
                    "request": "외주 설계자의 확인요청 사항입니다. 슬라이드 내용 및 연출 방향을 확인해 주세요."
                })

    return dict(findings_by_slide)


def format_slide_comment(findings: List[Dict[str, str]]) -> str:
    """
    슬라이드 1장에 들어갈 검수 의견들을 하나의 깔끔한 메모 본문으로 포맷팅합니다.
    """
    if len(findings) == 1:
        f = findings[0]
        if f["category"] == "확인요청 사항 있음":
            return (
                f"[SCU 품질검수 결과]\n"
                f"■ 구분: 확인요청 사항 있음\n"
                f"• 확인요청 내용: {f['issue']}"
            )
        return (
            f"[SCU 품질검수 결과]\n"
            f"■ 구분: {f['category']}\n"
            f"• 지적사항: {f['issue']}\n"
            f"• 수정요청: {f['request']}"
        )
    else:
        sections = [f"[SCU 품질검수 결과 (총 {len(findings)}건)]"]
        for idx, f in enumerate(findings, 1):
            if f["category"] == "확인요청 사항 있음":
                sections.append(
                    f"\n[{idx}] 구분: 확인요청 사항 있음\n"
                    f"• 확인요청 내용: {f['issue']}"
                )
            else:
                sections.append(
                    f"\n[{idx}] 구분: {f['category']}\n"
                    f"• 지적사항: {f['issue']}\n"
                    f"• 수정요청: {f['request']}"
                )
        return "\n".join(sections)


def add_comments_to_pptx(
    pptx_path: str,
    report_path: str,
    output_pptx_path: Optional[str] = None
) -> str:
    """
    보고서 결과를 바탕으로 PPTX의 해당 슬라이드에 파워포인트 정식 '메모(Comment)'를 삽입하여 새 파일로 저장합니다.
    """
    src_p = Path(pptx_path).resolve()
    if not src_p.exists():
        raise FileNotFoundError(f"원본 PPTX 파일을 찾을 수 없습니다: {src_p}")

    if not output_pptx_path:
        out_p = src_p.parent / f"{src_p.stem}_검수메모반영.pptx"
    else:
        out_p = Path(output_pptx_path).resolve()

    out_p.parent.mkdir(parents=True, exist_ok=True)

    findings_by_slide = parse_report_findings(report_path)
    total_findings_count = sum(len(v) for v in findings_by_slide.values())
    print(f"📋 보고서 파싱 완료: 총 {len(findings_by_slide)}개 슬라이드, {total_findings_count}건의 검수 의견 추출됨.")

    try:
        import win32com.client
        import pythoncom
    except ImportError:
        raise ImportError("PowerPoint COM 자동화를 위해 pywin32 패키지가 필요합니다: pip install pywin32")

    pythoncom.CoInitialize()
    ppt = None
    pres = None

    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        # 백그라운드에서 프레젠테이션 열기
        pres = ppt.Presentations.Open(str(src_p), WithWindow=False)
        slide_count = pres.Slides.Count
        print(f"📑 파워포인트 파일 오픈 완료 (총 {slide_count}개 슬라이드)")

        applied_slides = 0
        for slide_no, findings in sorted(findings_by_slide.items()):
            if 1 <= slide_no <= slide_count:
                slide = pres.Slides(slide_no)
                comment_text = format_slide_comment(findings)
                
                # 슬라이드 우측 상단(Left=850, Top=50)에 메모 추가
                slide.Comments.Add(
                    Left=850,
                    Top=50,
                    Author="SCU 품질검수 AI",
                    AuthorInitials="AI",
                    Text=comment_text
                )
                applied_slides += 1

        print(f"💬 총 {applied_slides}개 슬라이드에 검수 메모 삽입 완료!")

        # 새 파일로 저장
        pres.SaveAs(str(out_p))
        print(f"💾 메모가 반영된 PPTX 저장 완료: {out_p}")

    finally:
        if pres:
            try:
                pres.Close()
            except Exception:
                pass
        if ppt:
            try:
                ppt.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    return str(out_p)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="PPTX 슬라이드에 검수 결과 메모 자동 삽입")
    parser.add_argument("--pptx", "-p", required=True, help="입력 PPTX 파일 경로")
    parser.add_argument("--report", "-r", required=True, help="검수 결과 마크다운 보고서 경로")
    parser.add_argument("--output", "-o", help="출력 PPTX 파일 경로 (생략 시 _검수메모반영.pptx로 저장)")

    args = parser.parse_args()
    res = add_comments_to_pptx(args.pptx, args.report, args.output)
    print(f"Done: {res}")
