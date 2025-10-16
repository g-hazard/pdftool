#!/usr/bin/env python3
"""Merge multiple PDF files into a single document.

This script is designed to be invoked from a Windows context menu entry.
When run without an explicit output path it opens a save dialog so the user
can choose where to write the merged PDF.
"""

from __future__ import annotations

import re
import sys
import tempfile
import datetime
from pathlib import Path
from typing import Iterable, List, Optional, Sequence, Tuple

try:
    from winotify import Notification, audio
    WINOTIFY_AVAILABLE = True
except ImportError:
    WINOTIFY_AVAILABLE = False


def _load_pdf_backend() -> Tuple[Optional[object], Optional[object], Optional[str]]:
    """Return PdfReader, PdfWriter and module name if available."""
    try:
        from pypdf import PdfReader, PdfWriter  # type: ignore

        return PdfReader, PdfWriter, "pypdf"
    except ImportError:
        try:
            from PyPDF2 import PdfReader, PdfWriter  # type: ignore

            return PdfReader, PdfWriter, "PyPDF2"
        except ImportError:
            return None, None, None


def _parse_args(argv: Sequence[str]) -> Tuple[List[Path], Optional[Path]]:
    """Parse command line arguments without relying on argparse."""
    pdfs: List[Path] = []
    output: Optional[Path] = None
    it = iter(argv)

    for token in it:
        if token in ("-o", "--output"):
            try:
                output_token = next(it)
            except StopIteration as exc:  # pragma: no cover - defensive
                raise ValueError("Missing value for --output option") from exc
            output = Path(output_token).expanduser()
        else:
            # Handle Windows context menu passing multiple files as space-delimited string
            path = Path(token).expanduser()
            if not path.exists() and " " in token and token.lower().count('.pdf') > 1:
                # Token contains multiple .pdf references - likely multiple files
                # Split on .pdf boundaries to handle paths with spaces in filenames
                parts = re.split(r'(\.pdf)\s+', token, flags=re.IGNORECASE)
                potential_paths = []
                i = 0
                while i < len(parts):
                    if i + 1 < len(parts) and parts[i + 1].lower() == '.pdf':
                        potential_paths.append(parts[i] + parts[i + 1])
                        i += 2
                    elif parts[i].lower().endswith('.pdf'):
                        # Last part that already has .pdf
                        potential_paths.append(parts[i])
                        i += 1
                    else:
                        i += 1
                
                if len(potential_paths) > 1 and all(Path(p).exists() for p in potential_paths):
                    # All parts are valid paths, add them separately
                    for p in potential_paths:
                        pdfs.append(Path(p).expanduser())
                    continue
            pdfs.append(path)

    return pdfs, output


def _deduplicate(paths: Iterable[Path]) -> List[Path]:
    """Return input paths preserving order while removing duplicates."""
    resolved: List[Path] = []
    seen = set()
    for path in paths:
        try:
            key = path.resolve(strict=False)
        except Exception:  # pragma: no cover - unexpected, keep original path
            key = path
        if key not in seen:
            seen.add(key)
            resolved.append(path)
    return resolved


def _prepare_tk() -> Tuple[Optional["tkinter.Tk"], Optional[object], Optional[object]]:
    """Initialise Tkinter if available, return root, filedialog, messagebox."""
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
    except Exception:
        return None, None, None

    root = tk.Tk()
    root.withdraw()
    return root, filedialog, messagebox


def _choose_output_path(
    filedialog_module: object,
    first_pdf: Path,
) -> Optional[Path]:
    """Open a save dialog to pick the output path."""
    if filedialog_module is None:
        return None

    filedialog = filedialog_module  # avoid mypy error for attribute access
    default_name = f"{first_pdf.stem}_merged.pdf"
    initial_dir = str(first_pdf.parent)
    # pylint: disable=E1101  # attribute defined on tkinter filedialog
    destination = filedialog.asksaveasfilename(
        title="Merge PDF - Choose destination",
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        initialdir=initial_dir,
        initialfile=default_name,
    )
    if not destination:
        return None
    return Path(destination)


def _show_message(
    messagebox_module: Optional[object],
    kind: str,
    text: str,
) -> None:
    """Display a modern Windows 11 toast notification."""
    if WINOTIFY_AVAILABLE:
        try:
            toast = Notification(
                app_id="PDF Merge Tool",
                title="PDF Merge" if kind != "error" else "PDF Merge - Error",
                msg=text,
                duration="short" if kind == "info" else "long",
            )
            
            # Set icon based on message type
            if kind == "error":
                toast.set_audio(audio.Default, loop=False)
            elif kind == "info":
                toast.set_audio(audio.SMS, loop=False)
            
            toast.show()
            return
        except Exception:
            pass  # Fall back to Tkinter if toast fails
    
    # Fallback to Tkinter or console
    if messagebox_module is None:
        if kind == "error":
            print(f"Error: {text}", file=sys.stderr)
        else:
            print(text)
        return

    messagebox = messagebox_module  # avoid mypy attribute warning
    title = "Merge PDF"
    if kind == "error":
        messagebox.showerror(title, text)
    elif kind == "warning":
        messagebox.showwarning(title, text)
    else:
        messagebox.showinfo(title, text)


def merge_pdfs(pdfs: Sequence[Path], destination: Path) -> None:
    """Merge provided PDF files into `destination`."""
    PdfReader, PdfWriter, backend_name = _load_pdf_backend()
    if not PdfReader or not PdfWriter:
        raise RuntimeError(
            "Could not import a PDF backend. Install one with:\n"
            "  py -m pip install pypdf\n"
            "or\n"
            "  py -m pip install PyPDF2"
        )

    writer = PdfWriter()
    for pdf in pdfs:
        try:
            reader = PdfReader(str(pdf))
        except Exception as exc:
            raise RuntimeError(f"Failed to read '{pdf}': {exc}") from exc

        try:
            for page in reader.pages:
                writer.add_page(page)
        except Exception as exc:  # pragma: no cover - backend specific
            raise RuntimeError(f"Failed to merge '{pdf}': {exc}") from exc

    destination.parent.mkdir(parents=True, exist_ok=True)
    try:
        with destination.open("wb") as handle:
            writer.write(handle)
    except Exception as exc:
        raise RuntimeError(f"Failed to write '{destination}': {exc}") from exc

    if backend_name:
        flush = getattr(writer, "flush", None)
        if callable(flush):
            try:
                flush()
            except Exception:  # pragma: no cover - optional cleanup
                pass


def main(argv: Sequence[str]) -> int:
    _log(f"argv -> {list(argv)}")
    pdfs, explicit_output = _parse_args(argv)
    pdfs = _deduplicate(pdfs)
    _log(f"deduplicated -> {[str(p) for p in pdfs]}")

    root, filedialog_module, messagebox_module = _prepare_tk()
    try:
        if not pdfs:
            _show_message(
                messagebox_module,
                "error",
                "Please select at least 2 PDF files to merge",
            )
            return 1

        missing = [str(path) for path in pdfs if not path.exists()]
        if missing:
            formatted = "\n".join(missing)
            _show_message(
                messagebox_module,
                "error",
                f"Files not found:\n{formatted}",
            )
            return 1

        if len(pdfs) < 2:
            _show_message(
                messagebox_module,
                "warning",
                "Select 2 or more PDF files to merge",
            )
            return 1

        destination = explicit_output
        if destination is None:
            if filedialog_module is None:
                _show_message(
                    messagebox_module,
                    "error",
                    "Unable to show file save dialog",
                )
                return 1
            destination = _choose_output_path(filedialog_module, pdfs[0])
            if destination is None:
                # User cancelled the dialog
                return 0

        destination = destination.expanduser()
        if destination.suffix.lower() != ".pdf":
            destination = destination.with_suffix(".pdf")

        # Note: Windows Save As dialog already handles overwrite confirmation
        # so we don't need to ask again for a better UX
        merge_pdfs(pdfs, destination)

        _log(f"merged -> {len(pdfs)} files into {destination}")
        _show_message(
            messagebox_module,
            "info",
            f"Successfully merged {len(pdfs)} PDF files\n{destination.name}",
        )
        return 0
    except RuntimeError as exc:
        _show_message(messagebox_module, "error", str(exc))
        return 1
    finally:
        _log("main finished")
        if root is not None:
            try:
                root.destroy()
            except Exception:
                pass


LOG_FILE = Path(tempfile.gettempdir()) / "merge_pdfs.log"


def _log(message: str) -> None:
    """Append diagnostic messages to a temp file."""
    timestamp = datetime.datetime.now().isoformat(timespec="seconds")
    try:
        with LOG_FILE.open("a", encoding="utf-8") as handle:
            handle.write(f"[{timestamp}] {message}\n")
    except Exception:
        # Logging must never block merging; swallow errors silently.
        pass


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
