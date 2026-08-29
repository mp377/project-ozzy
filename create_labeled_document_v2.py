"""
create_labeled_document_v2.py

Interactive document creation with Purview sensitivity labeling, backed by a
persistent ("warm") PowerShell session.

WHY THE WARM SESSION MATTERS
----------------------------
Measured on this environment:

    Cold PowerShell process ....... ~92 seconds
    Warm session, first call ...... ~2 seconds
    Warm session, after 10 min .... ~0.8 seconds

The cost is engine bootstrap, not the labeling call itself. Every time a new
PowerShell process starts, the Purview client re-authenticates and rebuilds the
MIP engine. Version 1 of this script paid that cost on every run.

This version starts ONE PowerShell session, warms it, and keeps it alive for as
long as an Office application is running. Every document after the first is
labeled in roughly a second.

REQUIREMENTS
------------
    pip install python-docx openpyxl python-pptx

    Purview Information Protection client installed, and authenticated once:
        Set-Authentication

USAGE
-----
    # Interactive - prompts for everything
    python create_labeled_document_v2.py

    # Skip prompts by supplying values
    python create_labeled_document_v2.py --name Report --type docx --label "Restricted"

    # Keep the session alive and create documents one after another
    python create_labeled_document_v2.py --session

    # Session mode that exits automatically when all Office apps close
    python create_labeled_document_v2.py --session --until-office-closes

EXIT CODES
----------
    0   Success
    1   Labeling failed
    2   Document creation failed
    3   Invalid input or missing dependency
    4   Could not start the PowerShell session
"""

import argparse
import os
import queue
import subprocess
import sys
import threading
import time
import uuid

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

DEFAULT_SAVE_DIR = r"C:\Users\Miguel\LabelTest"

# Tenant-specific. Retrieve with: Get-Label | Select-Object DisplayName, Guid
LABELS = {
    "Public":            "994c2e99-9b3f-459d-8e3e-36de4c4f4f91",
    "Limited":           "7430563a-23b6-4772-b5b1-9317d899d4b3",
    "Restricted":        "7cd5dfcb-9dc7-4f57-985b-cabf186e9678",
    "Highly Restricted": "82c90564-5d3b-485b-933a-93e69a6e60ba",
}

DOC_TYPES = {
    "docx": ("Word document",       "word"),
    "xlsx": ("Excel workbook",      "excel"),
    "pptx": ("PowerPoint deck",     "powerpoint"),
}

# Office executables, used to detect whether an Office app is open
OFFICE_PROCESSES = {
    "WINWORD":  "Word",
    "EXCEL":    "Excel",
    "POWERPNT": "PowerPoint",
}

COMMAND_TIMEOUT = 300      # seconds to wait for a single PowerShell command
WARMUP_TIMEOUT  = 300      # seconds to wait for initial engine bootstrap


# --------------------------------------------------------------------------
# Persistent PowerShell session
# --------------------------------------------------------------------------

class WarmSession:
    """
    Holds a single long-lived PowerShell process with the Purview module
    loaded and the MIP engine bootstrapped.

    Commands are written to stdin. After each command a unique sentinel is
    echoed, so we know when the output for that command is complete.
    """

    def __init__(self, verbose=True):
        self.proc = None
        self.verbose = verbose
        self._out_q = queue.Queue()
        self._reader = None
        self._started = False

    # -- internals --------------------------------------------------------

    def _log(self, msg):
        if self.verbose:
            print(msg)

    def _drain_stdout(self):
        """Background thread: push every stdout line onto a queue."""
        for line in iter(self.proc.stdout.readline, ""):
            self._out_q.put(line.rstrip("\r\n"))
        self.proc.stdout.close()

    def _read_until(self, sentinel, timeout):
        """Collect lines until the sentinel appears. Returns the lines before it."""
        lines = []
        deadline = time.time() + timeout
        while True:
            remaining = deadline - time.time()
            if remaining <= 0:
                raise TimeoutError(f"No response within {timeout}s")
            try:
                line = self._out_q.get(timeout=min(remaining, 1.0))
            except queue.Empty:
                if self.proc.poll() is not None:
                    raise RuntimeError("PowerShell session ended unexpectedly")
                continue
            if sentinel in line:
                return lines
            lines.append(line)

    # -- lifecycle --------------------------------------------------------

    def start(self):
        """Spawn PowerShell, import the module, and warm the engine."""
        self._log("Starting PowerShell session...")

        try:
            self.proc = subprocess.Popen(
                ["powershell", "-NoProfile", "-NoLogo", "-NonInteractive", "-Command", "-"],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
            )
        except FileNotFoundError:
            raise RuntimeError("PowerShell not found. This script requires Windows.")

        self._reader = threading.Thread(target=self._drain_stdout, daemon=True)
        self._reader.start()
        self._started = True

        self.run("Import-Module PurviewInformationProtection -ErrorAction SilentlyContinue",
                 timeout=60)

        self._log("Warming the labeling engine (first call takes up to 2 minutes)...")
        t0 = time.perf_counter()
        self._warm_engine()
        elapsed = time.perf_counter() - t0
        self._log(f"Engine ready in {elapsed:.1f}s. Subsequent labels will be fast.\n")

    def _warm_engine(self):
        """
        Force a full engine bootstrap by labeling a throwaway file.

        A status query alone may short-circuit before the MIP engine loads.
        Running a real Set-FileLabel guarantees the whole path is warm, so the
        first document the user creates does not pay the ~90 second cost.
        """
        import tempfile
        warm_path = os.path.join(tempfile.gettempdir(), f"_warmup_{uuid.uuid4().hex}.docx")

        try:
            create_word(warm_path, title="warmup")
        except Exception as e:
            # If we cannot build a temp file, fall back to a status query.
            self._log(f"  (temp file unavailable: {e} — using status query instead)")
            self.run("Get-FileStatus -Path 'C:\\_warmup_.docx' -ErrorAction SilentlyContinue",
                     timeout=WARMUP_TIMEOUT)
            return

        try:
            label_id = LABELS["Public"]
            safe = warm_path.replace("'", "''")
            self.run(
                f"try {{ Set-FileLabel -Path '{safe}' -LabelId '{label_id}' "
                f"-ErrorAction Stop | Out-Null }} catch {{ }}",
                timeout=WARMUP_TIMEOUT,
            )
        finally:
            try:
                os.remove(warm_path)
            except OSError:
                pass

    def run(self, command, timeout=COMMAND_TIMEOUT):
        """Execute a PowerShell command in the warm session. Returns output lines."""
        if not self._started or self.proc.poll() is not None:
            raise RuntimeError("Session is not running")

        sentinel = f"__DONE_{uuid.uuid4().hex}__"
        payload = f"{command}\nWrite-Output '{sentinel}'\n"

        self.proc.stdin.write(payload)
        self.proc.stdin.flush()

        return self._read_until(sentinel, timeout)

    def stop(self):
        """Close the session cleanly."""
        if self.proc and self.proc.poll() is None:
            try:
                self.proc.stdin.write("exit\n")
                self.proc.stdin.flush()
                self.proc.wait(timeout=10)
            except Exception:
                self.proc.kill()
        self._started = False

    # -- Purview operations ----------------------------------------------

    def label_file(self, path, label_name):
        """Apply a label. Returns (success, message)."""
        label_id = LABELS.get(label_name)
        if not label_id:
            return False, f"Unknown label: {label_name}"

        safe_path = path.replace("'", "''")
        cmd = (
            f"try {{ "
            f"Set-FileLabel -Path '{safe_path}' -LabelId '{label_id}' -ErrorAction Stop "
            f"| Out-Null; Write-Output 'OK' "
            f"}} catch {{ Write-Output ('FAIL: ' + $_.Exception.Message) }}"
        )

        try:
            out = self.run(cmd)
        except Exception as e:
            return False, str(e)

        text = " ".join(out).strip()
        if "OK" in text and "FAIL" not in text:
            return True, "Label applied"
        return False, text or "No response from Set-FileLabel"

    def get_label(self, path):
        """Return the label name currently on the file, or None."""
        safe_path = path.replace("'", "''")
        cmd = (
            f"$s = Get-FileStatus -Path '{safe_path}' -ErrorAction SilentlyContinue; "
            f"if ($s -and $s.IsLabeled) {{ Write-Output $s.MainLabelName }} "
            f"else {{ Write-Output '' }}"
        )
        try:
            out = self.run(cmd, timeout=120)
        except Exception:
            return None
        text = " ".join(l for l in out if l.strip()).strip()
        return text or None


# --------------------------------------------------------------------------
# Office process detection
# --------------------------------------------------------------------------

def running_office_apps():
    """Return the set of Office app names currently running."""
    try:
        result = subprocess.run(
            ["tasklist", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=15,
        )
    except Exception:
        return set()

    found = set()
    for line in result.stdout.splitlines():
        name = line.split(",")[0].strip('"').upper()
        base = name[:-4] if name.endswith(".EXE") else name
        if base in OFFICE_PROCESSES:
            found.add(OFFICE_PROCESSES[base])
    return found


# --------------------------------------------------------------------------
# Document creation
# --------------------------------------------------------------------------

def create_word(path, title=None, content=None):
    try:
        from docx import Document
    except ImportError:
        raise RuntimeError("python-docx not installed. Run: pip install python-docx")
    doc = Document()
    if title:
        doc.add_heading(title, level=0)
    if content:
        for line in content.split("\n"):
            doc.add_paragraph(line)
    doc.save(path)


def create_excel(path, title=None, content=None):
    try:
        from openpyxl import Workbook
    except ImportError:
        raise RuntimeError("openpyxl not installed. Run: pip install openpyxl")
    wb = Workbook()
    ws = wb.active
    if title:
        ws.title = title[:31]
        ws["A1"] = title
    if content:
        for i, line in enumerate(content.split("\n"), start=3):
            ws.cell(row=i, column=1, value=line)
    wb.save(path)


def create_powerpoint(path, title=None, content=None):
    try:
        from pptx import Presentation
    except ImportError:
        raise RuntimeError("python-pptx not installed. Run: pip install python-pptx")
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = title or "Untitled"
    if content:
        slide.placeholders[1].text = content
    prs.save(path)


CREATORS = {
    "word": create_word,
    "excel": create_excel,
    "powerpoint": create_powerpoint,
}


def create_document(path, title=None, content=None):
    ext = os.path.splitext(path)[1].lower().lstrip(".")
    if ext not in DOC_TYPES:
        raise ValueError(f"Unsupported type '{ext}'. Use: {', '.join(DOC_TYPES)}")
    parent = os.path.dirname(os.path.abspath(path))
    if parent and not os.path.isdir(parent):
        os.makedirs(parent, exist_ok=True)
    CREATORS[DOC_TYPES[ext][1]](path, title, content)


# --------------------------------------------------------------------------
# Prompts
# --------------------------------------------------------------------------

def prompt_choice(question, options, default_index=0):
    """Numbered menu. Returns the chosen option."""
    print(f"\n{question}")
    for i, opt in enumerate(options, 1):
        marker = " (default)" if i - 1 == default_index else ""
        print(f"  {i}. {opt}{marker}")
    while True:
        raw = input(f"Choice [1-{len(options)}]: ").strip()
        if not raw:
            return options[default_index]
        if raw.isdigit() and 1 <= int(raw) <= len(options):
            return options[int(raw) - 1]
        print("  Not a valid choice.")


def prompt_text(question, default=None, required=True):
    suffix = f" [{default}]" if default else ""
    while True:
        raw = input(f"{question}{suffix}: ").strip()
        if raw:
            return raw
        if default is not None:
            return default
        if not required:
            return ""
        print("  A value is required.")


def prompt_document_spec(defaults=None):
    """
    Collect filename, location and label from the user.
    Returns (path, label, title).
    """
    defaults = defaults or {}

    print("\n" + "=" * 58)
    print(" New labeled document")
    print("=" * 58)

    # --- File name ---
    name = defaults.get("name") or prompt_text("File name (without extension)")
    name = os.path.splitext(name)[0]

    # --- Type ---
    if defaults.get("type"):
        ext = defaults["type"].lstrip(".").lower()
    else:
        labels = [f"{k}  —  {v[0]}" for k, v in DOC_TYPES.items()]
        chosen = prompt_choice("Document type:", labels, default_index=0)
        ext = chosen.split()[0]

    # --- Location ---
    location = defaults.get("location") or prompt_text(
        "Save to folder", default=DEFAULT_SAVE_DIR
    )
    location = os.path.expandvars(os.path.expanduser(location))

    # --- Label ---
    if defaults.get("label"):
        label = defaults["label"]
        if label not in LABELS:
            raise ValueError(f"Unknown label '{label}'. Valid: {', '.join(LABELS)}")
    else:
        label = prompt_choice("Sensitivity label:", list(LABELS.keys()),
                              default_index=1)

    # --- Optional title ---
    title = defaults.get("title")
    if title is None:
        title = prompt_text("Document heading (optional)", default="", required=False)

    path = os.path.normpath(os.path.join(location, f"{name}.{ext}"))

    print("\n" + "-" * 58)
    print(f"  File   : {path}")
    print(f"  Label  : {label}")
    if title:
        print(f"  Heading: {title}")
    print("-" * 58)

    return path, label, title or None


# --------------------------------------------------------------------------
# Workflow
# --------------------------------------------------------------------------

def make_one(session, path, label, title, open_after=False, verify=True):
    """Create and label a single document using the warm session."""

    if os.path.exists(path):
        answer = input(f"\n{path} already exists. Overwrite? [y/N]: ").strip().lower()
        if answer != "y":
            print("Skipped.")
            return False

    print("\nCreating document...")
    try:
        create_document(path, title)
    except Exception as e:
        print(f"[ERROR] Could not create the document: {e}")
        return False
    print("  Created.")

    print(f"Applying label '{label}'...")
    t0 = time.perf_counter()
    ok, msg = session.label_file(os.path.abspath(path), label)
    elapsed = time.perf_counter() - t0

    if not ok:
        print(f"[ERROR] Labeling failed: {msg}")
        print("        The document exists but is unlabeled.")
        return False

    print(f"  Labeled in {elapsed:.1f}s.")

    if verify:
        actual = session.get_label(os.path.abspath(path))
        if actual:
            state = "matches" if actual == label else "DOES NOT MATCH requested"
            print(f"  Verified on disk: {actual}  ({state})")
        else:
            print("  [WARN] Could not read a label back from the file.")

    if open_after:
        try:
            os.startfile(os.path.abspath(path))
            print("  Opened in Office.")
        except AttributeError:
            print("  [WARN] --open only works on Windows.")
        except Exception as e:
            print(f"  [WARN] Could not open it: {e}")

    return True


def session_loop(session, until_office_closes):
    """Create documents repeatedly while the session stays warm."""
    created = 0

    while True:
        if until_office_closes:
            apps = running_office_apps()
            if not apps:
                print("\nNo Office applications running. Closing session.")
                break
            print(f"\n[Office running: {', '.join(sorted(apps))}]")

        try:
            path, label, title = prompt_document_spec()
        except ValueError as e:
            print(f"[ERROR] {e}")
            continue
        except (KeyboardInterrupt, EOFError):
            print("\nCancelled.")
            break

        if make_one(session, path, label, title, open_after=True):
            created += 1

        try:
            again = input("\nCreate another document? [Y/n]: ").strip().lower()
        except (KeyboardInterrupt, EOFError):
            break
        if again == "n":
            break

    print(f"\n{created} document(s) created and labeled this session.")


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Create Office documents with Purview labels, using a warm session.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--name", help="File name without extension")
    parser.add_argument("--type", choices=list(DOC_TYPES), help="docx, xlsx or pptx")
    parser.add_argument("--location", help="Folder to save into")
    parser.add_argument("--label", help=f"One of: {', '.join(LABELS)}")
    parser.add_argument("--title", help="Heading inside the document")
    parser.add_argument("--session", action="store_true",
                        help="Stay open and create multiple documents")
    parser.add_argument("--until-office-closes", action="store_true",
                        help="In session mode, exit when no Office app is running")
    parser.add_argument("--open", action="store_true",
                        help="Open the document after labeling")
    parser.add_argument("--no-verify", action="store_true",
                        help="Skip reading the label back after applying")
    parser.add_argument("--quiet", action="store_true",
                        help="Suppress session startup chatter")

    args = parser.parse_args()

    print("=" * 58)
    print(" Labeled Document Creator  —  v2 (warm session)")
    print("=" * 58)

    if args.until_office_closes:
        apps = running_office_apps()
        if apps:
            print(f"Office running: {', '.join(sorted(apps))}")
        else:
            print("Note: no Office application is currently running.")

    session = WarmSession(verbose=not args.quiet)

    try:
        session.start()
    except Exception as e:
        print(f"[ERROR] Could not start the PowerShell session: {e}")
        print("        Check that the Purview client is installed and that")
        print("        Set-Authentication has been run.")
        sys.exit(4)

    try:
        if args.session:
            session_loop(session, args.until_office_closes)
        else:
            defaults = {
                "name": args.name,
                "type": args.type,
                "location": args.location,
                "label": args.label,
                "title": args.title,
            }
            defaults = {k: v for k, v in defaults.items() if v}

            try:
                path, label, title = prompt_document_spec(defaults)
            except ValueError as e:
                print(f"[ERROR] {e}")
                sys.exit(3)

            ok = make_one(session, path, label, title,
                          open_after=args.open,
                          verify=not args.no_verify)
            if not ok:
                sys.exit(1)

    except (KeyboardInterrupt, EOFError):
        print("\nInterrupted.")
    finally:
        session.stop()
        print("Session closed.")


if __name__ == "__main__":
    main()
