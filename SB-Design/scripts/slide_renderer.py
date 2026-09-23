"""
slide_renderer.py
PowerPoint COM 자동화를 활용하여 PPTX 슬라이드를 고해상도 PNG(1920x1080) 이미지로 일괄 렌더링하는 모듈.
"""

import os
import sys
import re
import time
from pathlib import Path
from typing import Dict, Optional, List

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


class SlideRenderer:
    def __init__(self, output_dir: Optional[str] = None):
        self.output_dir = Path(output_dir) if output_dir else None

    def render_presentation(
        self,
        pptx_path: str,
        output_dir: str,
        width: int = 1920,
        height: int = 1080
    ) -> Dict[int, str]:
        """
        PPTX 파일의 전체 슬라이드를 PNG 이미지로 일괄 렌더링하고,
        슬라이드 번호(1부터 시작) -> 이미지 파일 경로 매핑 딕셔너리를 반환합니다.
        """
        pptx_p = Path(pptx_path).resolve()
        if not pptx_p.exists():
            raise FileNotFoundError(f"PPTX 파일을 찾을 수 없습니다: {pptx_p}")

        out_p = Path(output_dir).resolve()
        out_p.mkdir(parents=True, exist_ok=True)

        try:
            import win32com.client
            import pythoncom
        except ImportError:
            print("⚠️ win32com 모듈을 찾을 수 없어 슬라이드 이미지 렌더링을 건너뜁니다.")
            return {}

        pythoncom.CoInitialize()
        ppt = None
        pres = None
        rendered_map: Dict[int, str] = {}

        try:
            t0 = time.time()
            ppt = win32com.client.Dispatch("PowerPoint.Application")
            pres = ppt.Presentations.Open(str(pptx_p), WithWindow=False)
            total_slides = pres.Slides.Count
            print(f"🖼️ [SlideRenderer] 슬라이드 {total_slides}장 이미지 렌더링 시작...")

            # PowerPoint COM의 배치 Export 기능 사용
            pres.Export(str(out_p), "PNG", width, height)
            elapsed = time.time() - t0
            print(f"✅ [SlideRenderer] 렌더링 완료 ({elapsed:.1f}초 소요)")

            # 생성된 파일 매핑 및 표준 파일명(slide_001.png 등)으로 정돈
            seen_files = set()
            for f in out_p.iterdir():
                if f.is_file() and f.suffix.lower() in [".png", ".jpg", ".jpeg"]:
                    m = re.search(r"(\d+)", f.stem)
                    if m:
                        idx = int(m.group(1))
                        standard_name = out_p / f"slide_{idx:03d}.png"
                        if f.resolve() != standard_name.resolve():
                            if standard_name.exists():
                                standard_name.unlink()
                            f.rename(standard_name)
                        rendered_map[idx] = str(standard_name)

        except Exception as e:
            print(f"⚠️ [SlideRenderer] 렌더링 중 오류 발생: {e}")
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

        return rendered_map


if __name__ == "__main__":
    if len(sys.argv) > 1:
        renderer = SlideRenderer()
        out_dir = "temp_render_test"
        res = renderer.render_presentation(sys.argv[1], out_dir)
        print(f"Rendered {len(res)} slides in {out_dir}")
    else:
        print("Usage: python slide_renderer.py <pptx_path>")
