"""엑셀 -> PDF 변환 엔진.

우선순위:
1) 로컬에 설치된 Microsoft Excel (COM 자동화) - 서식/인쇄영역을 그대로 반영, 가장 정확함.
2) LibreOffice(soffice.exe) - Excel이 없는 PC를 위한 대체 엔진.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
import time
from dataclasses import dataclass
from typing import Callable, Optional

SUPPORTED_EXTS = (".xlsx", ".xlsm", ".xls", ".xlsb", ".csv")

# xlsx 확장자별 COM 상수
XL_TYPE_PDF = 0
XL_QUALITY_STANDARD = 0


class ConversionError(Exception):
    pass


class ConversionCancelled(Exception):
    pass


@dataclass
class ConvertOptions:
    fit_to_page: bool = True          # 한 페이지 너비에 맞춤
    overwrite: bool = True            # 동일 이름 PDF 덮어쓰기
    output_mode: str = "same"         # "same" | "custom"
    output_dir: Optional[str] = None  # output_mode == "custom" 일 때 사용


def find_libreoffice() -> Optional[str]:
    """soffice.exe 경로를 찾는다. 없으면 None."""
    exe = shutil.which("soffice")
    if exe:
        return exe
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def resolve_output_path(src_path: str, options: ConvertOptions, taken: Optional[set] = None) -> str:
    """원본에 대응하는 PDF 경로.

    taken: 이번 배치에서 이미 쓰기로 정한 경로들. 다른 폴더의 같은 이름 파일을 한
    지정 폴더로 모으면 "덮어쓰기" 옵션과 상관없이 서로 덮어써버리므로 번호를 붙인다.
    """
    base = os.path.splitext(os.path.basename(src_path))[0]
    if options.output_mode == "custom" and options.output_dir:
        out_dir = options.output_dir
    else:
        out_dir = os.path.dirname(src_path)
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, base + ".pdf")
    taken = taken if taken is not None else set()

    def blocked(p: str) -> bool:
        return os.path.normcase(p) in taken or (not options.overwrite and os.path.exists(p))

    n = 1
    while blocked(out_path):
        out_path = os.path.join(out_dir, f"{base}({n}).pdf")
        n += 1
    taken.add(os.path.normcase(out_path))
    return out_path


class ExcelEngine:
    """하나의 배치 변환 동안 Excel 프로세스를 한 번만 띄워 재사용한다."""

    def __init__(self):
        self._app = None
        self._pythoncom = None
        self._win32com_client = None

    def _ensure_started(self):
        if self._app is not None:
            return
        if self._pythoncom is None:  # COM 초기화는 스레드당 한 번 (Excel 재시작 시에는 생략)
            import pythoncom
            import win32com.client as win32

            self._pythoncom = pythoncom
            self._win32com_client = win32
            pythoncom.CoInitialize()
        app = self._win32com_client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.AskToUpdateLinks = False
        app.EnableEvents = False
        self._app = app

    def convert(self, src_path: str, out_path: str, options: ConvertOptions):
        self._ensure_started()
        app = self._app
        wb = None
        try:
            wb = app.Workbooks.Open(
                os.path.abspath(src_path),
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
            )
            if options.fit_to_page:
                for ws in wb.Worksheets:
                    try:
                        ws.PageSetup.Zoom = False
                        ws.PageSetup.FitToPagesWide = 1
                        ws.PageSetup.FitToPagesTall = False
                    except Exception:
                        pass  # 차트 시트 등 PageSetup 미지원 시트는 건너뜀

            wb.ExportAsFixedFormat(
                Type=XL_TYPE_PDF,
                Filename=os.path.abspath(out_path),
                Quality=XL_QUALITY_STANDARD,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )
        except Exception as exc:  # noqa: BLE001
            if not self._is_alive():
                # Excel 프로세스가 죽었다 (크래시/강제종료). 다음 파일에서 새로 띄우도록
                # 버린다 — 안 그러면 남은 파일이 전부 같은 오류로 줄줄이 실패한다.
                self._app = None
                raise ConversionError(f"Excel이 응답하지 않아 다시 시작합니다: {exc}") from exc
            raise ConversionError(str(exc)) from exc
        finally:
            if wb is not None:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass

    def _is_alive(self) -> bool:
        try:
            _ = self._app.Version
            return True
        except Exception:  # noqa: BLE001
            return False

    def shutdown(self):
        if self._app is not None:
            try:
                self._app.Quit()
            except Exception:
                pass
            self._app = None
        if self._pythoncom is not None:
            try:
                self._pythoncom.CoUninitialize()
            except Exception:
                pass
            self._pythoncom = None


def convert_with_libreoffice(src_path: str, out_path: str, soffice_path: str, profile_dir: str):
    """profile_dir: 이 배치 전용 LibreOffice 프로필.

    기본 프로필을 쓰면 사용자가 LibreOffice를 열어둔 상태에서 변환이 아무 메시지 없이
    실패한다 (프로필은 프로세스 하나만 쓸 수 있음). 그래서 별도 프로필로 실행한다.
    """
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    # LibreOffice는 결과를 항상 "<원본이름>.pdf"로 만든다. 출력 폴더에 바로 만들면
    # 같은 이름의 앞선 결과(예: 다른 폴더의 보고서.csv → 보고서.pdf)를 덮어써 버리므로
    # 임시 폴더에 만든 뒤 정해둔 경로로 옮긴다.
    stage_dir = tempfile.mkdtemp(prefix="exceltopdf_out_")
    try:
        result = subprocess.run(
            [
                soffice_path,
                "--headless",
                "--norestore",
                "-env:UserInstallation=file:///" + profile_dir.replace("\\", "/"),
                "--convert-to",
                "pdf",
                "--outdir",
                stage_dir,
                src_path,
            ],
            capture_output=True,
            text=True,
            timeout=180,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        produced = os.path.join(stage_dir, os.path.splitext(os.path.basename(src_path))[0] + ".pdf")
        if not os.path.exists(produced):
            detail = (result.stderr or "").strip()
            raise ConversionError(detail or "LibreOffice가 PDF를 생성하지 못했습니다 (손상되었거나 암호가 걸린 파일일 수 있음).")
        shutil.move(produced, out_path)
    finally:
        shutil.rmtree(stage_dir, ignore_errors=True)


class BatchConverter:
    """파일 목록을 순차 변환하며 콜백으로 진행 상황을 알린다."""

    def __init__(self, options: ConvertOptions, engine_name: str = "excel"):
        self.options = options
        self.engine_name = engine_name
        self._cancel = False
        self._excel_engine: Optional[ExcelEngine] = None

    def cancel(self):
        self._cancel = True

    def run(
        self,
        files: list[str],
        on_progress: Callable[[str, str, str, Optional[str]], None],
        # on_progress(file_path, status, message, out_path)
    ):
        soffice_path = None
        lo_profile = None
        if self.engine_name == "excel":
            self._excel_engine = ExcelEngine()
        else:
            soffice_path = find_libreoffice()
            if not soffice_path:
                raise ConversionError("LibreOffice(soffice.exe)를 찾을 수 없습니다.")
            lo_profile = tempfile.mkdtemp(prefix="exceltopdf_lo_")

        taken: set = set()
        try:
            for src in files:
                if self._cancel:
                    on_progress(src, "취소됨", "", None)
                    continue

                on_progress(src, "변환중", "", None)
                out_path = resolve_output_path(src, self.options, taken)
                start = time.time()
                try:
                    if self.engine_name == "excel":
                        self._excel_engine.convert(src, out_path, self.options)
                    else:
                        convert_with_libreoffice(src, out_path, soffice_path, lo_profile)
                    elapsed = time.time() - start
                    on_progress(src, "완료", f"{os.path.basename(out_path)} ({elapsed:.1f}초)", out_path)
                except Exception as exc:  # noqa: BLE001
                    on_progress(src, "실패", str(exc), None)
        finally:
            if self._excel_engine is not None:
                self._excel_engine.shutdown()
            if lo_profile:
                shutil.rmtree(lo_profile, ignore_errors=True)
