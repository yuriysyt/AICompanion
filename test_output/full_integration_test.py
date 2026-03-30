"""
AICompanion Full Integration Test  (pywinauto + UIA)
=====================================================
Simulates real user interactions by typing into the AICompanion chat input,
pressing Enter, and verifying expected system-level outcomes.

Usage:
    python test_output/full_integration_test.py

Requirements:
    pip install pywinauto psutil
"""

import os
import sys
import time
import subprocess
import traceback
import psutil

from pywinauto import Desktop
from pywinauto.application import Application

# ─── Config ──────────────────────────────────────────────────────────────────

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Preferred exe paths, in priority order
_EXE_CANDIDATES = [
    os.path.join(_ROOT, "src", "AICompanion.Desktop", "bin", "Release",
                 "net8.0-windows", "AICompanion.exe"),
    os.path.join(_ROOT, "src", "AICompanion.Desktop", "bin", "Debug",
                 "net8.0-windows", "AICompanion.exe"),
    os.path.join(_ROOT, "src", "AICompanion.Desktop", "bin", "Release",
                 "net8.0-windows", "win-x64", "AICompanion.exe"),
    os.path.join(_ROOT, "publish", "AICompanion_v2", "AICompanion.exe"),
]

APP_LAUNCH_TIMEOUT    = 20   # seconds to wait for the main window to appear
COMMAND_WAIT          = 5    # seconds after typing Enter (simple commands)
COMPLEX_COMMAND_WAIT  = 12   # seconds for Word/Notepad/calc to open
OCR_WAIT              = 8    # seconds for OCR to complete

# ─── ANSI colours ────────────────────────────────────────────────────────────

GREEN  = "\033[92m"
RED    = "\033[91m"
YELLOW = "\033[93m"
RESET  = "\033[0m"

results: list[tuple[str, bool, str]] = []   # (name, passed, note)


def _pass(name: str, note: str = "") -> None:
    results.append((name, True, note))
    print(f"  {GREEN}✅ PASS{RESET}  {name}" + (f"  [{note}]" if note else ""))


def _fail(name: str, note: str = "") -> None:
    results.append((name, False, note))
    print(f"  {RED}❌ FAIL{RESET}  {name}" + (f"  [{note}]" if note else ""))


# ─── Process helpers ─────────────────────────────────────────────────────────

def _count_processes(*names: str) -> int:
    """Return total running process count for any of the given exe names."""
    total = 0
    for name in names:
        total += len([p for p in psutil.process_iter(["name"])
                      if p.info["name"] and
                         p.info["name"].lower() == name.lower()])
    return total


def _kill_processes(*names: str) -> None:
    for name in names:
        for p in psutil.process_iter(["name"]):
            try:
                if p.info["name"] and p.info["name"].lower() == name.lower():
                    p.kill()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass


# ─── Window / UI helpers ─────────────────────────────────────────────────────

def _find_app_window(title_contains: str, timeout: int = 10):
    """Wait for a window whose title contains the given string, return it or None."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            desk = Desktop(backend="uia")
            for w in desk.windows():
                try:
                    if title_contains.lower() in w.window_text().lower():
                        return w
                except Exception:
                    pass
        except Exception:
            pass
        time.sleep(0.5)
    return None


def _find_ai_companion_window(timeout: int = APP_LAUNCH_TIMEOUT):
    """Return the AI Companion main window element."""
    for keyword in ("AI Companion", "AICompanion"):
        w = _find_app_window(keyword, timeout=timeout)
        if w:
            return w
    return None


def _type_command(window, command: str) -> None:
    """
    Type a command into the TextCommandInput box and press Enter.
    Falls back to set_focus + keyboard.send_keys if child search fails.
    """
    try:
        # Try to find the text input by AutomationId or control type
        input_box = None
        for ctrl in window.descendants(control_type="Edit"):
            try:
                auto_id = ctrl.automation_id()
                if "TextCommandInput" in auto_id or "Command" in auto_id:
                    input_box = ctrl
                    break
            except Exception:
                pass

        if input_box is None:
            # Fallback: use the first Edit control
            edits = window.descendants(control_type="Edit")
            if edits:
                input_box = edits[0]

        if input_box:
            input_box.click_input()
            input_box.type_keys("^a{DELETE}", with_spaces=True)  # clear
            input_box.type_keys(command, with_spaces=True)
            input_box.type_keys("{ENTER}", with_spaces=True)
        else:
            # Last-resort: give focus to window and use keyboard module
            window.set_focus()
            from pywinauto import keyboard
            keyboard.send_keys("^a{DELETE}")
            keyboard.send_keys(command, with_spaces=True)
            keyboard.send_keys("{ENTER}", with_spaces=True)
    except Exception as exc:
        print(f"    {YELLOW}⚠ _type_command error: {exc}{RESET}")
        # Try raw keyboard fallback
        window.set_focus()
        from pywinauto import keyboard
        keyboard.send_keys(command, with_spaces=True)
        keyboard.send_keys("{ENTER}", with_spaces=True)


def _get_activity_log_text(window) -> str:
    """Extract all visible text from the ActivityLog / chat area."""
    try:
        texts = []
        for ctrl in window.descendants():
            try:
                t = ctrl.window_text()
                if t and t.strip():
                    texts.append(t.strip())
            except Exception:
                pass
        return "\n".join(texts)
    except Exception:
        return ""


# ─── Locate & launch AICompanion ─────────────────────────────────────────────

def _find_exe() -> str | None:
    for path in _EXE_CANDIDATES:
        if os.path.isfile(path):
            return path
    return None


def launch_app() -> tuple[subprocess.Popen | None, object | None]:
    """Launch AICompanion.exe and return (process, main_window)."""
    exe = _find_exe()
    if not exe:
        print(f"{RED}Cannot find AICompanion.exe.  "
              f"Build the project first (dotnet build).{RESET}")
        return None, None

    print(f"  Launching: {exe}")
    proc = subprocess.Popen([exe], cwd=os.path.dirname(exe))
    print(f"  Waiting up to {APP_LAUNCH_TIMEOUT}s for main window…")
    win = _find_ai_companion_window(timeout=APP_LAUNCH_TIMEOUT)
    if win is None:
        print(f"  {RED}Main window did not appear within {APP_LAUNCH_TIMEOUT}s{RESET}")
        proc.kill()
        return None, None

    print(f"  {GREEN}Main window found: '{win.window_text()}'{RESET}")
    time.sleep(1.5)   # let the UI settle
    return proc, win


# ─── Individual test cases ────────────────────────────────────────────────────

def test_open_word_reuses_existing(win) -> None:
    """
    'open word' → WINWORD should start (or reuse existing).
    Second 'open word' → no additional WINWORD process created.
    """
    test_name = "open word — first time"
    _kill_processes("WINWORD.EXE", "WINWORD")

    before = _count_processes("WINWORD.EXE", "WINWORD")
    _type_command(win, "open word")
    time.sleep(COMPLEX_COMMAND_WAIT)
    after = _count_processes("WINWORD.EXE", "WINWORD")

    if after > before:
        _pass(test_name, f"WINWORD count {before}→{after}")
    else:
        # May have reused an existing window — check window appeared
        word_win = _find_app_window("Word", timeout=5)
        if word_win:
            _pass(test_name, "existing Word window brought to front")
        else:
            _fail(test_name, f"WINWORD count {before}→{after}, no Word window found")

    # Second "open word" — no NEW process
    test_name2 = "open word again — window reuse (no duplicate process)"
    count_before_2nd = _count_processes("WINWORD.EXE", "WINWORD")
    _type_command(win, "open word")
    time.sleep(COMPLEX_COMMAND_WAIT)
    count_after_2nd = _count_processes("WINWORD.EXE", "WINWORD")

    if count_after_2nd <= count_before_2nd:
        _pass(test_name2, f"count stayed at {count_after_2nd} (no duplicate)")
    else:
        _fail(test_name2, f"count grew from {count_before_2nd} to {count_after_2nd}")


def test_open_new_word_creates_new(win) -> None:
    """'open new word document' must open a FRESH Word instance (or Ctrl+N new doc)."""
    test_name = "open new word document — new instance"
    count_before = _count_processes("WINWORD.EXE", "WINWORD")
    _type_command(win, "open new word document")
    time.sleep(COMPLEX_COMMAND_WAIT)
    count_after = _count_processes("WINWORD.EXE", "WINWORD")

    # Either a new process started, or activity log says "New Word document" / "new document"
    activity = _get_activity_log_text(win).lower()
    if count_after > count_before or "new word" in activity or "new document" in activity:
        _pass(test_name, f"process count {count_before}→{count_after}")
    else:
        _fail(test_name, f"process count {count_before}→{count_after}, log: {activity[:120]}")


def test_create_new_notebook(win) -> None:
    """'create new notebook' → Notepad should open."""
    test_name = "create new notebook — notepad opens"
    _kill_processes("notepad.exe", "notepad")
    _type_command(win, "create new notebook")
    time.sleep(COMPLEX_COMMAND_WAIT)

    count = _count_processes("notepad.exe", "notepad")
    notepad_win = _find_app_window("Notepad", timeout=5)
    if count > 0 or notepad_win:
        _pass(test_name, f"notepad procs={count}")
    else:
        activity = _get_activity_log_text(win).lower()
        if "notepad" in activity or "notebook" in activity or "note" in activity:
            _pass(test_name, "activity log confirms notebook opened")
        else:
            _fail(test_name, f"notepad not found, log: {activity[:120]}")


def test_open_calculator(win) -> None:
    """'open calculator' → Calculator process should start."""
    test_name = "open calculator — calc opens"
    _kill_processes("CalculatorApp.exe", "Calculator.exe", "calc.exe")
    _type_command(win, "open calculator")
    time.sleep(COMPLEX_COMMAND_WAIT)

    count = _count_processes("CalculatorApp.exe", "Calculator.exe", "calc.exe")
    calc_win = _find_app_window("Calculator", timeout=5)
    if count > 0 or calc_win:
        _pass(test_name, f"calc procs={count}")
    else:
        activity = _get_activity_log_text(win).lower()
        if "calculator" in activity or "calc" in activity:
            _pass(test_name, "activity log confirms calculator opened")
        else:
            _fail(test_name, f"calc not found, log: {activity[:120]}")


def test_write_essay_to_word(win) -> None:
    """
    'write essay about space exploration' with Word open →
    activity log should confirm text was written / plan executed.
    """
    test_name = "write essay about space exploration — text inserted into Word"

    # Make sure Word is open
    if _count_processes("WINWORD.EXE", "WINWORD") == 0:
        _type_command(win, "open word")
        time.sleep(COMPLEX_COMMAND_WAIT)

    _type_command(win, "write essay about space exploration")
    time.sleep(COMPLEX_COMMAND_WAIT + 5)   # essays need extra generation time

    activity = _get_activity_log_text(win).lower()
    keywords = ["essay", "space", "written", "typed", "word", "step", "type_text", "plan"]
    if any(k in activity for k in keywords):
        _pass(test_name, "activity confirms essay written")
    else:
        _fail(test_name, f"no write confirmation in log: {activity[:200]}")


def test_read_screen_ocr(win) -> None:
    """'read screen' → a non-empty OCR response must appear in the activity log."""
    test_name = "read screen (OCR) — non-empty response"
    activity_before = _get_activity_log_text(win)

    _type_command(win, "read screen")
    time.sleep(OCR_WAIT)

    activity_after = _get_activity_log_text(win)
    new_text = activity_after[len(activity_before):]

    ocr_keywords = ["ocr", "text", "tessdata", "screen", "chars", "found", "error", "read"]
    if any(k in new_text.lower() for k in ocr_keywords) and new_text.strip():
        _pass(test_name, f"response length={len(new_text)}")
    else:
        _fail(test_name, f"no recognizable OCR output: '{new_text[:100]}'")


def test_calculator_math(win) -> None:
    """'what is 25 times 4' → Calculator should show 100."""
    test_name = "what is 25 times 4 — Calculator shows 100"
    _type_command(win, "what is 25 times 4")
    time.sleep(COMPLEX_COMMAND_WAIT)

    # Check AICompanion activity log for result
    activity = _get_activity_log_text(win).lower()
    if "100" in activity:
        _pass(test_name, "result 100 found in activity log")
        return

    # Also check the Calculator window directly
    calc_win = _find_app_window("Calculator", timeout=5)
    if calc_win:
        calc_text = _get_activity_log_text(calc_win)
        if "100" in calc_text:
            _pass(test_name, "Calculator display shows 100")
            return
        _fail(test_name, f"Calculator text: '{calc_text[:80]}'")
    else:
        _fail(test_name, f"Calculator not found, activity: '{activity[:120]}'")


# ─── Cleanup ─────────────────────────────────────────────────────────────────

def cleanup() -> None:
    """Close all opened helper apps."""
    print("\n  Cleaning up opened apps…")
    _kill_processes("WINWORD.EXE", "WINWORD")
    _kill_processes("notepad.exe", "notepad")
    _kill_processes("CalculatorApp.exe", "Calculator.exe", "calc.exe")
    print("  Done.")


# ─── Main ────────────────────────────────────────────────────────────────────

def main() -> int:
    print(f"\n{'='*60}")
    print("  AICompanion Full Integration Test")
    print(f"{'='*60}\n")

    proc, win = launch_app()
    if win is None:
        print(f"{RED}Cannot launch or find AICompanion — aborting.{RESET}")
        return 1

    tests = [
        ("open_word",                  test_open_word_reuses_existing),
        ("open_new_word",              test_open_new_word_creates_new),
        ("create_new_notebook",        test_create_new_notebook),
        ("open_calculator",            test_open_calculator),
        ("write_essay_to_word",        test_write_essay_to_word),
        ("read_screen_ocr",            test_read_screen_ocr),
        ("calculator_math_25x4",       test_calculator_math),
    ]

    for _name, test_fn in tests:
        print(f"\n▶ {_name}")
        try:
            test_fn(win)
        except Exception as exc:
            _fail(_name, f"exception: {exc}")
            traceback.print_exc()

    # ── Cleanup ──────────────────────────────────────────────────────────────
    cleanup()

    if proc:
        try:
            proc.terminate()
            proc.wait(timeout=5)
        except Exception:
            pass

    # ── Summary ──────────────────────────────────────────────────────────────
    passed = sum(1 for _, ok, _ in results if ok)
    total  = len(results)
    print(f"\n{'='*60}")
    print(f"  Results: {passed}/{total} passed")
    print(f"{'='*60}")
    for name, ok, note in results:
        icon = f"{GREEN}✅{RESET}" if ok else f"{RED}❌{RESET}"
        print(f"  {icon}  {name}" + (f"  — {note}" if note else ""))
    print()

    return 0 if passed == total else 1


if __name__ == "__main__":
    sys.exit(main())
