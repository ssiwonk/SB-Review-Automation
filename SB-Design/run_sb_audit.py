"""
run_sb_audit.py
서울사이버대학교(SCU) SB 자동 검수 파이프라인 통합 실행 스크립트.

사용법:
  1. 자동 탐색 모드 (input/ 폴더 내 원고와 SB 파일 자동 매칭):
     python run_sb_audit.py

  2. 파일 명시 지정 모드:
     python run_sb_audit.py --manuscript "input/원고.docx" --sb "input/SB_01_v1.pptx"
"""

import os
import sys
import json
import argparse
from pathlib import Path
from typing import Optional, Tuple, Dict, Any

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

# scripts 모듈 경로 추가
CURRENT_DIR = Path(__file__).parent.resolve()
SCRIPTS_DIR = CURRENT_DIR / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from doc_parser import extract_manuscript_text
from sb_parser import SCUSBParser
from pptx_layout_checker import PPTXLayoutChecker
from report_generator import generate_markdown_report
from slide_renderer import SlideRenderer


def find_input_files(input_dir: Path) -> Tuple[Optional[Path], Optional[Path]]:
    """input 디렉토리에서 원고 파일과 SB 파일을 자동으로 감지합니다."""
    supported_doc_exts = [".hwp", ".hwpx", ".docx", ".doc", ".pdf", ".pptx"]
    
    files = [f for f in input_dir.iterdir() if f.is_file() and not f.name.startswith("~$")]

    manuscript_file = None
    sb_file = None

    # 1. 파일명 기반 1차 탐색
    for f in files:
        name_lower = f.name.lower()
        # SB 파일 감지
        if ("sb" in name_lower or "스토리보드" in name_lower) and f.suffix.lower() == ".pptx":
            if not sb_file:
                sb_file = f
        # 원고 파일 감지
        elif any(k in name_lower for k in ["원고", "교안", "본문", "강의록"]):
            if not manuscript_file and f.suffix.lower() in supported_doc_exts:
                manuscript_file = f

    # 2. 파일명에 키워드가 없는 경우 확장자 및 휴리스틱 기반 감지
    if not sb_file:
        pptx_candidates = [f for f in files if f.suffix.lower() == ".pptx" and f != manuscript_file]
        if pptx_candidates:
            sb_file = pptx_candidates[0]

    if not manuscript_file:
        doc_candidates = [f for f in files if f.suffix.lower() in supported_doc_exts and f != sb_file]
        if doc_candidates:
            manuscript_file = doc_candidates[0]

    return manuscript_file, sb_file


def find_course_specific_rule(
    sb_file_name: str,
    manuscript_name: str,
    contents_dir: Path
) -> Tuple[Optional[Path], Optional[str]]:
    """
    contents/ 폴더에서 현재 검수 중인 과목의 특화 규칙 마크다운 파일을 탐색하고 내용을 로드합니다.
    """
    if not contents_dir.exists():
        return None, None

    rule_files = [f for f in contents_dir.glob("*.md") if not f.name.startswith("_") and f.name.lower() != "readme.md"]
    if not rule_files:
        return None, None

    target_names = [sb_file_name.lower(), manuscript_name.lower()]

    # 1. 파일명 매칭 (예: 글로벌AI커머스.md가 sb_file_name에 포함되어 있는지)
    for rf in rule_files:
        c_name = rf.stem.lower()
        if any(c_name in tn for tn in target_names):
            with open(rf, "r", encoding="utf-8") as f:
                return rf, f.read()

    return None, None


def run_audit(
    manuscript_path: str,
    sb_path: str,
    output_dir: str = "output",
    auto_cleanup: bool = False
) -> Dict[str, Any]:
    """검수 파이프라인 전체를 구동합니다."""

    print("=" * 60)
    print("🚀 SCU 'SB-Design' 스토리보드 자동 검수 시스템 시작")
    print(f"📄 원고 파일: {manuscript_path}")
    print(f"📊 SB  파일: {sb_path}")
    print("=" * 60)

    # 1. 원고 파싱
    print("\n[Step 1] 교수 원고 텍스트 및 구조 추출 중...")
    doc_res = extract_manuscript_text(manuscript_path)
    if doc_res["status"] != "success":
        print(f"❌ 원고 파싱 오류: {doc_res['error_msg']}")
        return {"status": "error", "error": doc_res["error_msg"]}
    print(f"✅ 원고 추출 완료 (글자 수: {doc_res['char_count']:,}자)")

    # 2. 슬라이드 고해상도 이미지 렌더링 (2-Track 하이브리드 비전)
    rendered_dir = CURRENT_DIR / "rendered_slides"
    print("\n[Step 2] 🖼️ 2-Track 하이브리드: 슬라이드 고해상도 이미지 렌더링 중...")
    renderer = SlideRenderer()
    rendered_map = renderer.render_presentation(sb_path, str(rendered_dir))
    if rendered_map:
        print(f"✅ 슬라이드 {len(rendered_map)}장 이미지 렌더링 완료 ({rendered_dir.name}/)")
    else:
        print("ℹ️ 이미지 렌더링 생략 (PPTX 파싱 모드로 계속 진행)")

    # 3. SB 정밀 파싱 (메타데이터 + 표 오버레이 매핑)
    print("\n[Step 3] SB 슬라이드 구조 및 메타데이터 파싱 중...")
    sb_parser = SCUSBParser(sb_path)
    sb_slides = sb_parser.parse(rendered_images=rendered_map)
    print(f"✅ SB 파싱 완료 (총 {len(sb_slides)}개 슬라이드)")

    # 4. Track 1: 파이썬 정량 레이아웃 검사
    print("\n[Step 4] Track 1 정량 레이아웃/디자인 검사 실행 중...")
    layout_checker = PPTXLayoutChecker(sb_path)
    track1_findings = layout_checker.run_check()
    print(f"✅ Track 1 검사 완료: 지적사항 {len(track1_findings)}건 도출")
    for f in track1_findings[:3]:
        print(f"   - [{f['slide_no']}] ({f['category']}) {f['issue'][:40]}...")
    if len(track1_findings) > 3:
        print(f"   - ... 외 {len(track1_findings) - 3}건")

    # 4. contents/ 폴더에서 과목별 특화 검수 기준 탐색
    contents_dir = CURRENT_DIR / "contents"
    course_rule_file, course_rule_text = find_course_specific_rule(
        sb_file_name=Path(sb_path).name,
        manuscript_name=doc_res["file_name"],
        contents_dir=contents_dir
    )
    if course_rule_file:
        print(f"\n[Step 4] 🎯 과목별 특화 검수 기준 적용: [{course_rule_file.stem}]")
    else:
        print("\n[Step 4] ℹ️ 별도 과목 특화 기준 없음 (전사 공통 기준 적용)")

    # 5. 에이전트 정성 검수(Track 2)를 위한 페이로드 생성
    payload = {
        "manuscript_file": doc_res["file_name"],
        "manuscript_text": doc_res["text"],
        "sb_file": Path(sb_path).name,
        "total_slides": len(sb_slides),
        "sb_summary_for_llm": sb_parser.get_structured_summary_for_llm(),
        "track1_findings": track1_findings,
        "course_name": course_rule_file.stem if course_rule_file else None,
        "course_specific_rules": course_rule_text,
        "rendered_images": {str(k): v for k, v in rendered_map.items()}
    }

    # payload 임시 저장 (에이전트 또는 후속 분석용)
    payload_path = CURRENT_DIR / "audit_payload.json"
    with open(payload_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    # 마크다운 보고서 생성 (Track 1 우선 반영)
    print("\n[Step 4] 최종 검수 결과 마크다운 보고서 생성 중...")
    report_path = generate_markdown_report(
        manuscript_name=doc_res["file_name"],
        sb_name=Path(sb_path).name,
        total_slides=len(sb_slides),
        track1_findings=track1_findings,
        track2_findings=[], # AI 에이전트와 통합 시 병합됨
        output_dir=output_dir,
        course_name=course_rule_file.stem if course_rule_file else None
    )
    print(f"✅ 검수 보고서 저장 완료: {report_path}")
    print("=" * 60)

    if auto_cleanup:
        cleanup_temp_files(sb_name=Path(sb_path).name, keep_report_path=report_path, output_dir=output_dir)

    return {
        "status": "success",
        "report_path": report_path,
        "payload_path": str(payload_path),
        "track1_count": len(track1_findings),
        "total_slides": len(sb_slides)
    }


def cleanup_temp_files(
    sb_name: str,
    keep_report_path: Optional[str] = None,
    output_dir: str = "output"
) -> None:
    """
    검수 완료 후 분석 과정에서 생성된 중간 임시 파일(payload JSON, 슬라이드 렌더링 이미지 등)만
    깔끔하게 삭제합니다. 이전의 최종 검수 결과 보고서(output/SB_검수결과_*.md)는 이력 보존을 위해 절대로 삭제하지 않습니다.
    """
    # 1. audit_payload.json 임시 데이터 파일 삭제
    payload_path = CURRENT_DIR / "audit_payload.json"
    if payload_path.exists():
        try:
            payload_path.unlink()
            print("🧹 [Cleanup] 임시 분석 파일(audit_payload.json) 삭제 완료")
        except Exception as e:
            print(f"⚠️ [Cleanup] 임시 파일 삭제 실패: {e}")

    # 2. rendered_slides 임시 슬라이드 이미지 폴더 삭제
    for r_name in ["rendered_slides", ".rendered_slides"]:
        r_dir = CURRENT_DIR / r_name
        if r_dir.exists():
            try:
                import shutil
                shutil.rmtree(r_dir)
                print(f"🧹 [Cleanup] 임시 슬라이드 렌더링 이미지({r_name}) 삭제 완료")
            except Exception as e:
                print(f"⚠️ [Cleanup] 렌더링 폴더 삭제 실패 ({r_name}): {e}")





def main():
    parser = argparse.ArgumentParser(description="SCU SB 자동 검수 파이프라인")
    parser.add_argument("--manuscript", "-m", help="원고 파일 경로 (hwp, hwpx, docx, pdf, pptx)")
    parser.add_argument("--sb", "-s", help="검수 대상 SB 파일 경로 (pptx)")
    parser.add_argument("--output", "-o", default="output", help="출력 폴더 경로 (기본값: output)")
    parser.add_argument("--cleanup", action="store_true", default=True, help="검수 완료 후 임시 분석 파일 및 이전 초안 보고서 자동 삭제 (기본값: True)")
    parser.add_argument("--no-cleanup", dest="cleanup", action="store_false", help="임시 파일 및 이전 초안 보존")


    args = parser.parse_args()

    project_root = CURRENT_DIR.parent
    input_dir = project_root / "input"

    m_path = args.manuscript
    s_path = args.sb

    if not m_path or not s_path:
        found_m, found_s = find_input_files(input_dir)
        if not m_path and found_m:
            m_path = str(found_m)
        if not s_path and found_s:
            s_path = str(found_s)

    if not m_path or not s_path:
        print("❌ 원고 파일 또는 SB 파일을 찾을 수 없습니다.")
        print(f"   input/ 폴더({input_dir})에 파일을 넣거나, 명령어로 지정해주세요:")
        print("   python run_sb_audit.py --manuscript <원고경로> --sb <SB경로>")
        sys.exit(1)

    res = run_audit(m_path, s_path, output_dir=args.output, auto_cleanup=args.cleanup)

    if res["status"] != "success":
        sys.exit(1)


if __name__ == "__main__":
    main()
