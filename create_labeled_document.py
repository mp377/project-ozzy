"""
create_labeled_document.py

Creates a Microsoft Office document (Word, Excel or PowerPoint) and applies a
Microsoft Purview sensitivity label to it as part of the creation process.

WHY THIS APPROACH WORKS
-----------------------
Purview cannot label a file that is open in an Office application: the app holds
an exclusive write lock and the labeling client is refused access. By creating
the file and labeling it BEFORE any application opens it, there is no lock to
contend with. The document exists labeled from the moment it is created.

REQUIREMENTS
------------
    pip install python-docx openpyxl python-pptx

    Microsoft Purview Information Protection client installed, and an
    authenticated session established beforehand:

        Set-Authentication

    apply_label_batch.py must be present (same folder by default).

USAGE
-----
    # Word document
    python create_labeled_document.py "C:\\path\\Report.docx" "Highly Restricted"

    # Excel workbook
    python create_labeled_document.py "C:\\path\\Model.xlsx" "Restricted"

    # PowerPoint deck
    python create_labeled_document.py "C:\\path\\Deck.pptx" "Limited"

    # With a heading and body content
    python create_labeled_document.py "C:\\path\\Report.docx" "Restricted" \
        --title "Q3 Portfolio Review" --content "Prepared by the Risk team."

    # Create, label, then open it
    python create_labeled_document.py "C:\\path\\Report.docx" "Limited" --open

    # Verify the label survives a save cycle (see --verify note below)
    python create_labeled_document.py "C:\\path\\Report.docx" "Limited" --verify

    # Machine-readable output, for calling from another program
    python create_labeled_document.py "C:\\path\\Report.docx" "Limited" --json

EXIT CODES
----------
    0   Document created and labeled successfully
    1   Document created but labeling failed
    2   Document creation failed
    3   Invalid arguments or missing dependency
"""

import argparse
import json
import os
import subprocess
import sys
import time

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

# Path to the labeling engine. Defaults to the same directory as this script.
DEFAULT_LABELER = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "apply_label_batch.py"
)

# Labeling can take up to ~2 minutes on documents with content.
LABEL_TIMEOUT_SECONDS = 300

SUPPORTED_EXTENSIONS = {
    ".docx": "word",
    ".xlsx": "excel",
    ".pptx": "powerpoint",
}

VALID_LABELS = [
    "Public",
    "Limited",
    "Restricted",
    "Highly Restricted",
]


# --------------------------------------------------------------------------
# Document creation
# --------------------------------------------------------------------------

def create_word(path, title=None, content=None):
    """Create a .docx file. Returns the path."""
    try:
        from docx import Document
    except ImportError:
        raise RuntimeError(
            "python-docx is not installed. Run: pip install python-docx"
        )

    doc = Document()
    if title:
        doc.add_heading(title, level=0)
    if content:
        for paragraph in content.split("\n"):
            doc.add_paragraph(paragraph)
    doc.save(path)
    return path


def create_excel(path, title=None, content=None):
    """Create a .xlsx file. Returns the path."""
    try:
        from openpyxl import Workbook
    except ImportError:
        raise RuntimeError(
            "openpyxl is not installed. Run: pip install openpyxl"
        )

    wb = Workbook()
    ws = wb.active
    if title:
        ws.title = title[:31]          # Excel caps sheet names at 31 chars
        ws["A1"] = title
        ws["A1"].font = ws["A1"].font.copy(bold=True, size=14)
    if content:
        for i, line in enumerate(content.split("\n"), start=3):
            ws.cell(row=i, column=1, value=line)
    wb.save(path)
    return path


def create_powerpoint(path, title=None, content=None):
    """Create a .pptx file. Returns the path."""
    try:
        from pptx import Presentation
    except ImportError:
        raise RuntimeError(
            "python-pptx is not installed. Run: pip install python-pptx"
        )

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])   # Title and Content
    slide.shapes.title.text = title or "Untitled"
    if content:
        slide.placeholders[1].text = content
    prs.save(path)
    return path


CREATORS = {
    "word": create_word,
    "excel": create_excel,
    "powerpoint": create_powerpoint,
}


def create_document(path, title=None, content=None):
    """
    Dispatch to the right creator based on file extension.
    Creates parent directories if they do not exist.
    """
    ext = os.path.splitext(path)[1].lower()
    kind = SUPPORTED_EXTENSIONS.get(ext)

    if kind is None:
        raise ValueError(
            f"Unsupported extension '{ext}'. "
            f"Supported: {', '.join(sorted(SUPPORTED_EXTENSIONS))}"
        )

    parent = os.path.dirname(os.path.abspath(path))
    if parent and not os.path.isdir(parent):
        os.makedirs(parent, exist_ok=True)

    return CREATORS[kind](path, title, content)


# --------------------------------------------------------------------------
# Labeling
# --------------------------------------------------------------------------

def apply_label(path, label, labeler_path):
    """
    Invoke the labeling engine on a closed file.
    Returns (success: bool, output: str).
    """
    if not os.path.isfile(labeler_path):
        return False, f"Labeling script not found at: {labeler_path}"

    try:
        result = subprocess.run(
            [sys.executable, labeler_path, path, label],
            capture_output=True,
            text=True,
            timeout=LABEL_TIMEOUT_SECONDS,
        )
    except subprocess.TimeoutExpired:
        return False, f"Labeling timed out after {LABEL_TIMEOUT_SECONDS}s"
    except Exception as e:
        return False, f"Labeling failed to run: {e}"

    output = (result.stdout or "") + (result.stderr or "")

    # The engine exits non-zero when any file fails.
    if result.returncode != 0:
        return False, output.strip()

    return True, output.strip()


def get_label_status(path):
    """
    Query the applied label via PowerShell.
    Returns the label name, or None if unlabeled / unavailable.
    """
    ps = (
        f"$s = Get-FileStatus -Path '{path}'; "
        "if ($s -and $s.IsLabeled) { $s.MainLabelName } else { '' }"
    )
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
            capture_output=True,
            text=True,
            timeout=120,
        )
        name = result.stdout.strip()
        return name if name else None
    except Exception:
        return None


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description=(
            "Create a Microsoft Office document and apply a Purview "
            "sensitivity label to it during creation."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "path",
        help="Full path of the document to create (.docx, .xlsx or .pptx)",
    )
    parser.add_argument(
        "label",
        help=f"Sensitivity label. One of: {', '.join(VALID_LABELS)}",
    )
    parser.add_argument("--title", help="Heading or sheet title for the document")
    parser.add_argument("--content", help="Body text. Use \\n for line breaks.")
    parser.add_argument(
        "--labeler",
        default=DEFAULT_LABELER,
        help="Path to apply_label_batch.py (defaults to the same folder)",
    )
    parser.add_argument(
        "--open",
        action="store_true",
        help="Open the document in its Office application after labeling",
    )
    parser.add_argument(
        "--verify",
        action="store_true",
        help="Query the label after applying and report what is actually set",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite the file if it already exists",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Emit a JSON result object instead of human-readable output",
    )

    args = parser.parse_args()

    start = time.perf_counter()

    result = {
        "path": args.path,
        "requestedLabel": args.label,
        "created": False,
        "labeled": False,
        "verifiedLabel": None,
        "error": None,
        "elapsedSeconds": None,
    }

    def emit_and_exit(code):
        if result["elapsedSeconds"] is None:
            result["elapsedSeconds"] = round(time.perf_counter() - start, 2)
        if args.json:
            print(json.dumps(result, indent=2))
        sys.exit(code)

    def say(msg):
        if not args.json:
            print(msg)

    # --- Validate label ---------------------------------------------------
    if args.label not in VALID_LABELS:
        result["error"] = (
            f"Invalid label '{args.label}'. Valid: {', '.join(VALID_LABELS)}"
        )
        say(f"[ERROR] {result['error']}")
        emit_and_exit(3)

    # --- Guard against overwriting ---------------------------------------
    if os.path.exists(args.path) and not args.overwrite:
        result["error"] = "File already exists. Use --overwrite to replace it."
        say(f"[ERROR] {result['error']}")
        emit_and_exit(3)

    # --- Create -----------------------------------------------------------
    say(f"[1/2] Creating document: {args.path}")
    try:
        create_document(args.path, args.title, args.content)
        result["created"] = True
        say("      Created.")
    except Exception as e:
        result["error"] = str(e)
        say(f"[ERROR] Could not create document: {e}")
        emit_and_exit(2)

    # --- Label ------------------------------------------------------------
    say(f"[2/2] Applying label '{args.label}' — this can take up to 2 minutes...")
    ok, output = apply_label(args.path, args.label, args.labeler)
    result["labeled"] = ok

    if not ok:
        result["error"] = output
        say(f"[ERROR] Labeling failed:\n{output}")
        say("\nThe document was created but is UNLABELED.")
        say("Check that Set-Authentication has been run in this session.")
        emit_and_exit(1)

    say(f"      Label applied: {args.label}")

    # --- Verify -----------------------------------------------------------
    if args.verify:
        say("      Verifying...")
        actual = get_label_status(args.path)
        result["verifiedLabel"] = actual
        if actual:
            match = "matches" if actual == args.label else "DOES NOT MATCH"
            say(f"      Verified label on disk: {actual}  ({match} requested)")
        else:
            say("      [WARN] Verification could not read a label from the file.")

    result["elapsedSeconds"] = round(time.perf_counter() - start, 2)

    # --- Open -------------------------------------------------------------
    if args.open:
        say("      Opening in Office...")
        try:
            os.startfile(os.path.abspath(args.path))
        except AttributeError:
            say("      [WARN] --open is only supported on Windows.")
        except Exception as e:
            say(f"      [WARN] Could not open the document: {e}")

    say(f"\nDone in {result['elapsedSeconds']}s — {args.path} is labeled '{args.label}'.")
    emit_and_exit(0)


if __name__ == "__main__":
    main()
