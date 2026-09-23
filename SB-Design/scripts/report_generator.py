"""
report_generator.py
Track 1(파이썬 정량 검사)과 Track 2(LLM 정성 검수) 결과를 통합하여
사용자 지정 4대 표준 항목 체계로 최종 마크다운 보고서를 생성하는 모듈.

[보고서 표준 구조]
1. 💡 총평 (각 항목의 지적사항 갯수 합계 표 & 종합 코멘트)
2. 📝 내용 검수 결과 (원고 대조 내용 누락·왜곡, 교수설계, 정보과밀 분할 등)
3. 📐 레이아웃 검수 결과 (규격, 폰트 20pt Bold, 화면 이탈, 도해화 제안 등)
4. ⚖️ 저작권 및 출처 점검 결과 (슬라이드 노트 출처/라이선스 번호 미기재 현황)
"""

import os
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any, Optional

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def classify_finding(item: Dict[str, Any], is_track1: bool = False) -> str:
    """
    지적사항을 3대 영역('copyright', 'layout', 'content') 중 하나로 분류합니다.
    """
    cat = item.get("category", "")

    # 1. 저작권 / 출처
    if any(k in cat for k in ["출처", "저작권", "게티이미지", "라이선스"]):
        return "copyright"

    # 2. 내용 / 교수설계 관련 키워드
    content_keywords = [
        "내용", "교수설계", "학습목표", "학습개요", "학습정리", "퀴즈", "평가",
        "원고", "불일치", "누락", "왜곡", "오탈자", "더미문자", "목차", "문장", "종결"
    ]
    if any(k in cat for k in content_keywords):
        return "content"

    # 3. 레이아웃 / 서식 관련 키워드
    layout_keywords = [
        "레이아웃", "서식", "폰트", "규격", "오버플로우", "화면이탈",
        "시각화", "도해화", "디자인", "배치", "제목규격", "인터랙션"
    ]
    if any(k in cat for k in layout_keywords):
        return "layout"

    # 기본값: Track 1은 layout, Track 2는 content
    return "layout" if is_track1 else "content"


def sort_key(item: Dict[str, Any]) -> int:
    """슬라이드 번호 숫자 기준 정렬 키"""
    s_no = str(item.get("slide_no", ""))
    digits = "".join(filter(str.isdigit, s_no))
    return int(digits) if digits else 999


def slide_str_sort_key(s_no: str) -> int:
    """슬라이드 문자열('P. 06' 등) 정렬 키"""
    digits = "".join(filter(str.isdigit, str(s_no)))
    return int(digits) if digits else 999


def generate_markdown_report(
    manuscript_name: str,
    sb_name: str,
    total_slides: int,
    track1_findings: List[Dict[str, Any]],
    track2_findings: Optional[List[Dict[str, Any]]] = None,
    output_dir: str = "output",
    course_name: Optional[str] = None,
    learning_summary_markdown: Optional[str] = None
) -> str:
    """
    검수 결과를 취합하여 4개 표준 항목 구조의 마크다운 보고서를 생성합니다.
    반환값: 생성된 마크다운 파일의 절대 경로
    """
    track2_findings = track2_findings or []

    content_findings: List[Dict[str, Any]] = []
    layout_findings: List[Dict[str, Any]] = []
    copyright_findings: List[Dict[str, Any]] = []
    source_missing_slides: List[str] = []

    # Track 1 지적사항 분류
    for item in track1_findings:
        c_type = classify_finding(item, is_track1=True)
        s_no = item.get("slide_no", "-")
        issue_text = item.get("issue", "")
        if c_type == "copyright":
            if s_no not in source_missing_slides:
                source_missing_slides.append(s_no)
            # 일반적인 슬라이드 노트 출처/라이선스 미기재는 콜아웃 블록에만 나열하고 표에는 추가하지 않음
            is_generic_missing = any(k in issue_text for k in ["노트에 출처", "출처 누락", "출처 미기재", "기재되지 않음", "유료 이미지 번호"])
            if not is_generic_missing:
                copyright_findings.append(item)
        elif c_type == "layout":
            layout_findings.append(item)
        else:
            content_findings.append(item)

    # Track 2 지적사항 분류
    for item in track2_findings:
        c_type = classify_finding(item, is_track1=False)
        s_no = item.get("slide_no", "-")
        issue_text = item.get("issue", "")
        if c_type == "copyright":
            if s_no not in source_missing_slides:
                source_missing_slides.append(s_no)
            is_generic_missing = any(k in issue_text for k in ["노트에 출처", "출처 누락", "출처 미기재", "기재되지 않음", "유료 이미지 번호"])
            if not is_generic_missing:
                copyright_findings.append(item)
        elif c_type == "layout":
            layout_findings.append(item)
        else:
            content_findings.append(item)


    # 정렬
    content_findings.sort(key=sort_key)
    layout_findings.sort(key=sort_key)
    copyright_findings.sort(key=sort_key)
    source_missing_slides.sort(key=slide_str_sort_key)

    # 건수 집계
    content_count = len(content_findings)
    layout_count = len(layout_findings)
    copyright_count = len(source_missing_slides) + len(copyright_findings)
    total_finding_count = content_count + layout_count + copyright_count

    # 타임스탬프 및 파일명 구성
    now = datetime.now()
    now_str = now.strftime("%Y-%m-%d %H:%M:%S")
    file_timestamp = now.strftime("%y%m%d_%H%M")

    base_title = Path(sb_name).stem.replace(".pptx", "")
    report_filename = f"SB_검수결과_{base_title}_{file_timestamp}.md"

    out_path = Path(output_dir)
    out_path.mkdir(parents=True, exist_ok=True)
    report_file_path = out_path / report_filename

    # 마크다운 본문 구성
    md_lines = [
        f"# [SB 검수 결과 보고서] {base_title}",
        "",
        f"- **검수 일시**: {now_str}",
        f"- **교수 원고**: `{manuscript_name}`",
        f"- **검수 대상 SB**: `{sb_name}` (총 {total_slides}개 슬라이드)",
    ]

    if course_name:
        md_lines.append(f"- **적용 과목 특화 기준**: `SB-Design/contents/{course_name}.md` (과목 고유 기준 반영)")

    md_lines.extend([
        "",
        "---",
        "",
        "## 1. 💡 총평",
        "",
        "### 📊 지적사항 갯수 합계",
        "",
        "| 검수 영역 | 지적 건수 | 주요 검수 내용 |",
        "| :--- | :---: | :--- |",
        f"| **2. 내용 검수 결과** | **{content_count}건** | 원고 대조 내용 누락·왜곡, 교수설계 논리, 정보과밀 분할 등 |",
        f"| **3. 레이아웃 검수 결과** | **{layout_count}건** | 슬라이드 규격(16:9), 본문 폰트(20pt Bold), 화면 이탈, 도해화 제안 등 |",
        f"| **4. 저작권 및 출처 점검 결과** | **{copyright_count}건** | 슬라이드 노트 이미지 출처(`#1`, 게티이미지 번호 등) 미기재 |",
        f"| **합계** | **{total_finding_count}건** | **전체 지적 및 개선 권고사항 총합** |",
        "",
        "### 📝 종합 평가 및 권고사항",
        f"1. **내용 및 교수설계**: 총 {content_count}건의 보완 사항이 도출되었습니다. 원고의 핵심 개념 반영 여부와 학습목표/퀴즈 피드백의 명확성을 점검하고, 정보과밀 슬라이드는 2장 분할 또는 레이아웃 재배치를 진행하세요.",
        f"2. **서식 및 레이아웃**: 총 {layout_count}건의 규격/서식 미비 사항이 확인되었습니다. 본문 폰트 20pt Bold 적용 및 화면 경계 이탈(오버플로우) 개체를 우선 정돈하세요.",
        f"3. **저작권 및 출처**: 총 {copyright_count}건의 이미지 출처 미표기가 확인되었습니다. 영상 촬영 전 슬라이드 노트에 명확한 출처 번호 및 라이선스 식별 번호를 기재해야 합니다.",
        "",
        "> [!TIP]",
        "> 지적사항을 수정한 후 최신 SB 파일을 다시 검수하여 모든 지적사항이 완결되었는지 재검증하는 것을 권장합니다.",
        "",
        "---",
        "",
        "## 2. 📝 내용 검수 결과",
        "",
        "| 슬라이드 번호 | 구분 | 지적사항 (무슨 문제인가요?) | 수정요청 (이렇게 바로 고쳐주세요!) |",
        "| :---: | :---: | :--- | :--- |"
    ])

    if not content_findings:
        md_lines.append("| - | - | 지적사항 없음 (원고 내용 충실히 반영 및 교수설계 기준 통과) | - |")
    else:
        for f in content_findings:
            s_no = f.get("slide_no", "-")
            cat = f.get("category", "내용")
            issue = str(f.get("issue", "")).replace("\n", "<br>")
            rec = str(f.get("recommendation", "")).replace("\n", "<br>")
            md_lines.append(f"| **{s_no}** | {cat} | {issue} | {rec} |")

    # 학습정리 AI 추천 문안이 있는 경우 내용 섹션 하단에 삽입
    if learning_summary_markdown:
        md_lines.extend([
            "",
            "---",
            "",
            "### 💡 [선택 권고 / 참고] 학습정리 서술형 AI 추천 완성 문안",
            learning_summary_markdown.strip()
        ])

    md_lines.extend([
        "",
        "---",
        "",
        "## 3. 📐 레이아웃 검수 결과",
        "",
        "| 슬라이드 번호 | 구분 | 지적사항 (무슨 문제인가요?) | 수정요청 (이렇게 바로 고쳐주세요!) |",
        "| :---: | :---: | :--- | :--- |"
    ])

    if not layout_findings:
        md_lines.append("| - | - | 지적사항 없음 (슬라이드 규격, 폰트 20pt Bold 및 레이아웃 기준 통과) | - |")
    else:
        for f in layout_findings:
            s_no = f.get("slide_no", "-")
            cat = f.get("category", "레이아웃")
            issue = str(f.get("issue", "")).replace("\n", "<br>")
            rec = str(f.get("recommendation", "")).replace("\n", "<br>")
            md_lines.append(f"| **{s_no}** | {cat} | {issue} | {rec} |")

    md_lines.extend([
        "",
        "---",
        "",
        "## 4. ⚖️ 저작권 및 출처 점검 결과",
        ""
    ])

    if source_missing_slides:
        slides_str = ", ".join(source_missing_slides)
        md_lines.extend([
            "> [!IMPORTANT]",
            f"> **저작권, 출처 미기재** : {slides_str} (총 {len(source_missing_slides)}개 슬라이드)",
            "> *위 슬라이드들은 본문에 이미지/사진이 포함되어 있으나 슬라이드 노트에 저작물 출처(`#1`, `#출처/N`) 또는 유료 라이선스 번호가 누락되었습니다.*",
            "",
            "- **조치 권고사항**: 본문에 사용된 모든 이미지에 대해 외주 개발사 및 설계자는 촬영 전 슬라이드 노트에 명확한 출처 번호(`#1` 등) 또는 게티이미지 라이선스 번호를 기재해야 합니다."
        ])
    else:
        md_lines.extend([
            "> [!NOTE]",
            "> 본문 이미지의 슬라이드 노트 출처 및 라이선스 표기가 모든 슬라이드에서 정상 확인되었습니다."
        ])

    # 기타 출처 지적사항 표
    if copyright_findings:
        md_lines.extend([
            "",
            "| 슬라이드 번호 | 구분 | 지적사항 (무슨 문제인가요?) | 수정요청 (이렇게 바로 고쳐주세요!) |",
            "| :---: | :---: | :--- | :--- |"
        ])
        for f in copyright_findings:
            s_no = f.get("slide_no", "-")
            cat = f.get("category", "저작권/출처")
            issue = str(f.get("issue", "")).replace("\n", "<br>")
            rec = str(f.get("recommendation", "")).replace("\n", "<br>")
            md_lines.append(f"| **{s_no}** | {cat} | {issue} | {rec} |")

    report_content = "\n".join(md_lines)
    with open(report_file_path, "w", encoding="utf-8") as f:
        f.write(report_content)

    return str(report_file_path.resolve())


if __name__ == "__main__":
    test_track1 = [
        {"slide_no": "P. 06", "category": "저작권/출처", "issue": "슬라이드 노트 출처 누락", "recommendation": "슬라이드 노트에 #1 출처 기재"},
        {"slide_no": "P. 12", "category": "레이아웃", "issue": "화면 경계 이탈", "recommendation": "화면 내 배치"},
        {"slide_no": "P. 20", "category": "폰트", "issue": "폰트 16pt", "recommendation": "20pt Bold 변경"}
    ]
    test_track2 = [
        {"slide_no": "P. 05", "category": "내용누락", "issue": "원고 내용 일부 누락", "recommendation": "내용 추가 보완"}
    ]
    path = generate_markdown_report("원고.docx", "SB_샘플.pptx", 50, test_track1, test_track2)
    print(f"Generated test report: {path}")
