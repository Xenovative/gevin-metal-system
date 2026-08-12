"""
Export a completed invoice workbook to PDF.

Excel is the source of truth: we convert the .xlsx print layout directly
(LibreOffice headless preferred; Excel COM on Windows as fallback).
"""
from __future__ import annotations

import logging
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from config import OUTPUT_DIR

logger = logging.getLogger(__name__)


def _find_soffice() -> str | None:
    """Return path to LibreOffice soffice binary, or None."""
    env = os.environ.get("LIBREOFFICE_PATH") or os.environ.get("SOFFICE_PATH")
    if env and Path(env).exists():
        return env
    for name in ("soffice", "libreoffice", "soffice.exe"):
        found = shutil.which(name)
        if found:
            return found
    # Common Windows install locations
    for candidate in (
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
    ):
        if candidate.exists():
            return str(candidate)
    return None


def _export_via_libreoffice(excel_path: Path, pdf_path: Path) -> str:
    soffice = _find_soffice()
    if not soffice:
        raise FileNotFoundError(
            "LibreOffice (soffice) not found. Install LibreOffice or set LIBREOFFICE_PATH."
        )

    out_dir = pdf_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    # LibreOffice writes <stem>.pdf into --outdir; use a temp dir then move
    with tempfile.TemporaryDirectory(prefix="gevin_pdf_") as tmp:
        tmp_dir = Path(tmp)
        cmd = [
            soffice,
            "--headless",
            "--nologo",
            "--nolockcheck",
            "--nodefault",
            "--nofirststartwizard",
            "--convert-to",
            "pdf",
            "--outdir",
            str(tmp_dir),
            str(excel_path.resolve()),
        ]
        logger.info("Excel→PDF via LibreOffice: %s", " ".join(cmd))
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
        if proc.returncode != 0:
            logger.error(
                "LibreOffice convert failed rc=%s stdout=%s stderr=%s",
                proc.returncode,
                (proc.stdout or "")[:500],
                (proc.stderr or "")[:500],
            )
            raise RuntimeError(
                f"LibreOffice PDF convert failed (rc={proc.returncode}): "
                f"{(proc.stderr or proc.stdout or '').strip()[:300]}"
            )

        produced = tmp_dir / f"{excel_path.stem}.pdf"
        if not produced.exists():
            # Some LO versions alter the name slightly
            pdfs = list(tmp_dir.glob("*.pdf"))
            if not pdfs:
                raise FileNotFoundError(
                    f"LibreOffice did not produce a PDF in {tmp_dir}"
                )
            produced = pdfs[0]

        if pdf_path.exists():
            pdf_path.unlink()
        shutil.move(str(produced), str(pdf_path))

    size = pdf_path.stat().st_size
    logger.info("Excel→PDF LibreOffice OK: %s (%s bytes)", pdf_path, size)
    return str(pdf_path.resolve())


def _export_via_excel_com(excel_path: Path, pdf_path: Path) -> str:
    """Windows Excel COM ExportAsFixedFormat (requires Excel installed)."""
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 not installed; cannot use Excel COM PDF export"
        ) from exc

    excel = None
    wb = None
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        abs_xlsx = str(excel_path.resolve())
        abs_pdf = str(pdf_path.resolve())
        logger.info("Excel→PDF via Excel COM: %s → %s", abs_xlsx, abs_pdf)
        wb = excel.Workbooks.Open(abs_xlsx, ReadOnly=True)
        # 0 = xlTypePDF
        wb.ExportAsFixedFormat(0, abs_pdf)
    except Exception:
        logger.exception("Excel COM PDF export failed")
        raise
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
            try:
                del excel
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    if not pdf_path.exists():
        raise FileNotFoundError(f"Excel COM did not create PDF: {pdf_path}")
    size = pdf_path.stat().st_size
    logger.info("Excel→PDF Excel COM OK: %s (%s bytes)", pdf_path, size)
    return str(pdf_path.resolve())


def export_excel_to_pdf(excel_path: str | Path) -> str:
    """
    Convert a filled invoice workbook to PDF beside it (same stem).

    Tries LibreOffice first, then Excel COM on Windows.
    """
    excel_path = Path(excel_path)
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel not found: {excel_path}")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    pdf_path = excel_path.with_suffix(".pdf")
    # Prefer writing into OUTPUT_DIR if excel lives elsewhere
    if excel_path.parent.resolve() != OUTPUT_DIR.resolve():
        pdf_path = OUTPUT_DIR / f"{excel_path.stem}.pdf"

    errors: list[str] = []

    try:
        return _export_via_libreoffice(excel_path, pdf_path)
    except Exception as exc:
        errors.append(f"LibreOffice: {exc}")
        logger.warning("LibreOffice PDF export unavailable/failed: %s", exc)

    if os.name == "nt":
        try:
            return _export_via_excel_com(excel_path, pdf_path)
        except Exception as exc:
            errors.append(f"Excel COM: {exc}")
            logger.warning("Excel COM PDF export unavailable/failed: %s", exc)

    msg = (
        "Cannot export Excel to PDF. Install LibreOffice (recommended) "
        "or Microsoft Excel on Windows. Details: " + " | ".join(errors)
    )
    logger.error(msg)
    raise RuntimeError(msg)
