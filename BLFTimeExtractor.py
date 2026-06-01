# Short name suggestion: BLFTimeExtractor.py
"""
BLF Time Extractor
- Recursively finds .blf files in a selected folder, reads measurement start and end times via Vector vSignalyzer COM API,
  and writes a summary Excel report with filename, path, start and end timestamps.
- Keeps a single vSignalyzer COM instance for efficiency, robust error handling, and safe Excel output.
"""

from pathlib import Path
import os
import logging
import openpyxl
import win32com.client
from tkinter import Tk, filedialog, messagebox
from typing import Iterable, List, Optional, Tuple

# Configure basic logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")


def find_blf_files(folder: Path) -> List[Path]:
    """Recursively find .blf files under `folder` and return sorted list of Paths."""
    if not folder.exists() or not folder.is_dir():
        return []
    # Use rglob for simplicity and performance on modern filesystems
    return sorted(p for p in folder.rglob("*.blf") if p.is_file())


def get_blf_times_from_app(app, file_path: Path) -> Tuple[Optional[str], Optional[str]]:
    """
    Given an active vSignalyzer COM application instance and a file path, open the measurement,
    extract formatted start and end times, then close the measurement. Returns (start_str, end_str)
    or (None, None) on failure.
    """
    try:
        blf = app.OpenMeasurement(str(file_path))
        # MeasurementStartTime / MeasurementEndTime are COM Date objects with Format method
        start_time = getattr(blf, "MeasurementStartTime", None)
        end_time = getattr(blf, "MeasurementEndTime", None)

        start_str = None
        end_str = None
        if start_time is not None:
            try:
                start_str = start_time.Format("%Y-%m-%d %H:%M:%S")
            except Exception:
                start_str = str(start_time)
        if end_time is not None:
            try:
                end_str = end_time.Format("%Y-%m-%d %H:%M:%S")
            except Exception:
                end_str = str(end_time)

        try:
            blf.Close()
        except Exception:
            # best-effort close; continue even if close fails
            logging.debug("Failed to close BLF measurement object for %s", file_path)

        return start_str, end_str
    except Exception as exc:
        logging.warning("Failed to read BLF file %s: %s", file_path, exc)
        return None, None


def generate_excel_report(rows: Iterable[Tuple[int, str, str, Optional[str], Optional[str]]], output_path: Path) -> Path:
    """Create an Excel workbook and save the provided rows. Returns path to saved file."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BLF Metadata"
    ws.append(["Sr No", "File Name", "File Path", "Measurement Start", "Acquisition End"])
    for row in rows:
        ws.append(list(row))
    # Ensure output directory exists
    output_path.parent.mkdir(parents=True, exist_ok=True)
    # If file exists, append timestamp to avoid accidental overwrite
    final_path = output_path
    if final_path.exists():
        ts = Path.cwd().name  # fallback if timestamp creation fails
        try:
            from datetime import datetime
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        except Exception:
            pass
        final_path = output_path.with_name(f"{output_path.stem}_{ts}{output_path.suffix}")
    wb.save(final_path)
    return final_path


def process_blf_folder(source_folder: Path, dest_folder: Path) -> Tuple[int, int, Optional[Path]]:
    """
    Find BLF files under source_folder, extract times and save an Excel report in dest_folder.
    Returns (found_count, processed_count, saved_report_path_or_None).
    """
    blf_files = find_blf_files(source_folder)
    if not blf_files:
        logging.info("No .blf files found under %s", source_folder)
        return 0, 0, None

    rows = []
    processed = 0
    try:
        # Create a single vSignalyzer COM instance to improve performance
        app = win32com.client.Dispatch("vSignalyzer.Application")
    except Exception as exc:
        logging.error("Failed to start vSignalyzer COM application: %s", exc)
        # Return with zero processed so caller can report or retry
        return len(blf_files), 0, None

    try:
        for idx, blf_path in enumerate(blf_files, start=1):
            start_str, end_str = get_blf_times_from_app(app, blf_path)
            rows.append((idx, blf_path.name, str(blf_path), start_str, end_str))
            processed += 1
    finally:
        # Attempt to quit the COM application gracefully
        try:
            app.Quit()
        except Exception:
            pass

    # Prepare output path and save Excel report
    output_file = dest_folder / "BLF_Metadata_Report.xlsx"
    try:
        saved = generate_excel_report(rows, output_file)
        logging.info("Report saved at: %s", saved)
        return len(blf_files), processed, saved
    except Exception as exc:
        logging.error("Failed to save Excel report: %s", exc)
        return len(blf_files), processed, None


def choose_folder(title: str) -> Optional[Path]:
    """Show a folder selection dialog and return the chosen Path or None."""
    root = Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title=title)
    try:
        root.destroy()
    except Exception:
        pass
    return Path(folder) if folder else None


def main() -> None:
    src = choose_folder("Select Folder Containing .blf Files")
    if not src:
        logging.info("Source folder selection cancelled.")
        return

    dst = choose_folder("Select Folder to Save Report")
    if not dst:
        logging.info("Output folder selection cancelled.")
        return

    found, processed, report_path = process_blf_folder(src, dst)
    if found == 0:
        messagebox.showinfo("BLF Time Extractor", "No .blf files found.")
    elif report_path is None:
        messagebox.showwarning("BLF Time Extractor", f"Processed {processed}/{found} files but failed to save report.")
    else:
        messagebox.showinfo("BLF Time Extractor", f"Report saved at:\n{report_path}")


if __name__ == "__main__":
    main()