#!/usr/bin/env python3
"""
여러 Excel 파일(.xls, .xlsx)을 한 번에 각각 PDF로 변환하는 스크립트.

기본 동작:
- 입력 경로의 Excel 파일을 탐색
- 파일마다 동일한 이름의 PDF를 출력 디렉터리에 생성
- Windows 환경에서는 Microsoft Excel COM 자동화 사용

사용 예시:
    python excel_to_pdf_batch.py "C:/input_excels" "C:/output_pdfs"
    python excel_to_pdf_batch.py "C:/input_excels" "C:/output_pdfs" --recursive
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


XL_TYPE_PDF = 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="여러 Excel 파일을 개별 PDF 파일로 일괄 변환합니다."
    )
    parser.add_argument("input_dir", type=Path, help="Excel 파일들이 있는 폴더")
    parser.add_argument("output_dir", type=Path, help="PDF 저장 폴더")
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="하위 폴더까지 재귀적으로 검색",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="동일한 이름의 PDF가 있으면 덮어쓰기",
    )
    return parser.parse_args()


def find_excel_files(input_dir: Path, recursive: bool) -> list[Path]:
    pattern = "**/*" if recursive else "*"
    files = [
        p
        for p in input_dir.glob(pattern)
        if p.is_file() and p.suffix.lower() in {".xlsx", ".xls"}
    ]
    return sorted(files)


def convert_with_excel_com(excel_files: list[Path], output_dir: Path, overwrite: bool) -> None:
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "pywin32가 필요합니다. 먼저 `pip install pywin32`를 실행하세요."
        ) from exc

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        for src in excel_files:
            dst = output_dir / f"{src.stem}.pdf"

            if dst.exists() and not overwrite:
                print(f"[SKIP] 이미 존재함: {dst}")
                continue

            print(f"[CONVERT] {src} -> {dst}")
            workbook = excel.Workbooks.Open(str(src.resolve()))
            try:
                workbook.ExportAsFixedFormat(XL_TYPE_PDF, str(dst.resolve()))
            finally:
                workbook.Close(SaveChanges=False)

    finally:
        excel.Quit()


def main() -> int:
    args = parse_args()

    input_dir: Path = args.input_dir
    output_dir: Path = args.output_dir

    if not input_dir.exists() or not input_dir.is_dir():
        print(f"입력 폴더가 올바르지 않습니다: {input_dir}", file=sys.stderr)
        return 1

    output_dir.mkdir(parents=True, exist_ok=True)

    excel_files = find_excel_files(input_dir, args.recursive)
    if not excel_files:
        print("변환할 Excel 파일(.xls, .xlsx)을 찾지 못했습니다.")
        return 0

    try:
        convert_with_excel_com(excel_files, output_dir, args.overwrite)
    except Exception as exc:
        print(f"변환 중 오류 발생: {exc}", file=sys.stderr)
        return 2

    print(f"완료: 총 {len(excel_files)}개 파일 처리")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
