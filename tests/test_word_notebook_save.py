import os, sys, time, subprocess, ctypes
from pathlib import Path
from datetime import datetime
from typing import Optional
import pytest
import pyautogui
import pywinauto
from pywinauto import Desktop, Application
from pywinauto.keyboard import send_keys
try:
    from pywinauto.findwindows import ElementNotFoundError
except ImportError:
    ElementNotFoundError = Exception

SCREENSHOT_DIR = Path(__file__).parent / "screenshots"
SCREENSHOT_DIR.mkdir(exist_ok=True)

APP_EXE = str(
    Path(__file__).parent.parent
    / "src" / "AICompanion.Desktop" / "bin" / "Release"
    / "net8.0-windows" / "win-x64" / "AICompanion.exe"
)
WORD_PATHS = [
    r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
]
DOCS = Path.home() / "Documents"


def screenshot(name):
    ts = datetime.now().strftime("%H%M%S")
    p = str(SCREENSHOT_DIR / (ts + "_" + name + ".png"))
    try:
        pyautogui.screenshot().save(p)
        print("  [SS] " + p)
    except Exception as e:
        print("  [SS-ERR] " + str(e))
    return p


def find_word():
    for p in WORD_PATHS:
        if os.path.exists(p):
            return p
    return None


def wait_win(title, timeout=15):
    """Return a WindowSpecification (supports child_window) or None."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            for w in Desktop(backend="uia").windows():
                try:
                    t = w.window_text()
                    if title.lower() in t.lower():
                        print("  [WIN] Found: " + t)
                        # Return a proper WindowSpecification via Application.connect
                        try:
                            app = Application(backend="uia").connect(handle=w.handle)
                            return app.top_window()
                        except Exception:
                            return w
                except Exception:
                    pass
        except Exception:
            pass
        time.sleep(0.5)
    print("  [WIN] NOT FOUND: " + title)
    return None


def win_click_center(win):
    """Click the center of a window using win32 RECT coordinates."""
    try:
        r = win.rectangle()
        cx = (r.left + r.right) // 2
        cy = (r.top + r.bottom) // 2 + 60
        pyautogui.click(cx, cy)
        time.sleep(0.3)
        print("  [CLK] (" + str(cx) + "," + str(cy) + ")")
    except Exception as e:
        print("  [CLK-ERR] " + str(e))


def kill_all(name):
    """Force-kill all processes with this exe name."""
    try:
        subprocess.call(["taskkill","/f","/im",name+".exe"],
                        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        time.sleep(0.5)
    except Exception:
        pass

def close_win(partial):
    try:
        for w in Desktop(backend="uia").windows():
            try:
                if partial.lower() in w.window_text().lower():
                    w.close()
                    time.sleep(0.3)
            except Exception:
                pass
    except Exception:
        pass


def dismiss():
    time.sleep(0.5)
    found = False
    try:
        for w in Desktop(backend="uia").windows():
            try:
                t = w.window_text().lower()
                keywords = ("save", "format", "replace", "keep", "compatible", "encode")
                if any(k in t for k in keywords):
                    print("  [DLG] " + w.window_text())
                    w.set_focus()
                    time.sleep(0.2)
                    send_keys("{ENTER}")
                    time.sleep(0.5)
                    found = True
            except Exception:
                pass
    except Exception:
        pass
    return found


def paste(text):
    """Set clipboard text and press Ctrl+V. Works on any keyboard layout."""
    import win32clipboard
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text)
    win32clipboard.CloseClipboard()
    time.sleep(0.15)
    send_keys("^v")
    time.sleep(0.3)


def send_cmd(win, cmd):
    """Type a command into AI Companion's text input and press Enter."""
    print("  [CMD] " + cmd)
    sent = False
    # --- Try 1: Application.connect to get WindowSpecification with child_window ---
    try:
        app = Application(backend="uia").connect(title_re=".*AI Companion.*")
        dlg = app.top_window()
        b = dlg.child_window(auto_id="TextCommandInput", control_type="Edit")
        b.set_focus()
        b.set_edit_text("")
        time.sleep(0.1)
        paste(cmd)
        time.sleep(0.15)
        send_keys("{ENTER}")
        time.sleep(0.5)
        sent = True
        print("  [CMD] sent via TextCommandInput")
    except Exception as e:
        print("  [CMD-WARN] TextCommandInput not found: " + str(e)[:80])
    if sent:
        return
    # --- Try 2: click quick-command button by text (e.g. "Notepad", "Word") ---
    try:
        app = Application(backend="uia").connect(title_re=".*AI Companion.*")
        dlg = app.top_window()
        # Try to find a button matching the first word of cmd
        first_word = cmd.split()[0] if cmd else cmd
        for word in cmd.split():
            try:
                btn = dlg.child_window(title_re="(?i).*" + word + ".*", control_type="Button")
                btn.click_input()
                sent = True
                print("  [CMD] clicked button: " + word)
                time.sleep(0.5)
                break
            except Exception:
                pass
    except Exception as e:
        print("  [CMD-WARN] Button click failed: " + str(e)[:80])
    if sent:
        return
    # --- Fallback: focus window, paste text, Enter ---
    try:
        win.set_focus()
    except Exception:
        pass
    time.sleep(0.3)
    paste(cmd)
    send_keys("{ENTER}")
    time.sleep(0.5)
    print("  [CMD] sent via fallback paste+Enter")


def click_body(win):
    win_click_center(win)


def saveas_dialog(path):
    time.sleep(0.8)
    try:
        for w in Desktop(backend="uia").windows():
            try:
                t = w.window_text().lower()
                if "save as" in t or "save a copy" in t:
                    print("  [SA] " + w.window_text())
                    screenshot("saveas")
                    try:
                        fn = w.child_window(auto_id="1148", control_type="Edit")
                        fn.set_focus()
                        fn.set_edit_text(str(path))
                    except Exception:
                        paste(str(path))
                    time.sleep(0.3)
                    send_keys("{ENTER}")
                    time.sleep(1.2)
                    return True
            except Exception:
                pass
    except Exception:
        pass
    paste(str(path))
    send_keys("{ENTER}")
    time.sleep(1.0)
    return False


def count_dlg():
    n = 0
    try:
        for w in Desktop(backend="uia").windows():
            try:
                t = w.window_text().lower()
                if any(k in t for k in ("save", "format", "replace", "keep", "encode")):
                    n += 1
            except Exception:
                pass
    except Exception:
        pass
    return n


def verify(path, sentinel):
    if Path(path).exists():
        d = Path(path).read_bytes()
        assert sentinel.encode() in d, \
            "Sentinel NOT in file! first200=" + str(d[:200])
        print("  [OK] " + Path(path).name + " (" + str(len(d)) + " bytes)")
        return
    stem = Path(path).stem.rsplit("_", 1)[0]
    found = sorted(DOCS.glob(stem + "_*"),
                   key=lambda f: f.stat().st_mtime, reverse=True)
    found = [f for f in found if (time.time() - f.stat().st_mtime) < 300]
    assert found, "FILE NOT SAVED! Expected: " + str(path)
    d = found[0].read_bytes()
    assert sentinel.encode() in d, \
        "Sentinel not in " + found[0].name + ": " + str(d[:200])
    print("  [OK-ALT] " + found[0].name)
    found[0].unlink(missing_ok=True)


class App:
    def __init__(self):
        self.proc = None
        self.win = None

    def start(self):
        if not os.path.exists(APP_EXE):
            print("  [SKIP] EXE not found: " + APP_EXE)
            return False
        print("  [LAUNCH] " + os.path.basename(APP_EXE))
        self.proc = subprocess.Popen(
            [APP_EXE], creationflags=subprocess.CREATE_NEW_PROCESS_GROUP
        )
        w = wait_win("AI Companion", 20)
        if w is None:
            w = wait_win("AICompanion", 5)
        if w is None:
            print("  [WARN] App window not found")
            return False
        self.win = w
        screenshot("app_launched")
        return True

    def stop(self):
        try:
            if self.proc:
                self.proc.terminate()
                self.proc.wait(timeout=5)
        except Exception:
            pass


# =========================================================
# TEST 1: NOTEPAD -- open / type / Ctrl+S / verify
# =========================================================
def test_notepad_create_write_save():
    print("\n" + "=" * 60)
    print("TEST 1: NOTEPAD -- create / write / save")
    print("=" * 60)

    fx = App()
    sentinel = "NOTEPAD_SAVE_2026"
    save_path = DOCS / ("test_notepad_" + datetime.now().strftime("%H%M%S") + ".txt")
    close_win("Notepad")
    time.sleep(0.4)

    try:
        ok = fx.start()
        if ok and fx.win:
            send_cmd(fx.win, "open notepad")
        else:
            subprocess.Popen(["notepad.exe"])

        np = wait_win("Notepad", 12)
        assert np is not None, "FAIL: Notepad did not open"
        screenshot("np_open")
        print("  PASS Step1: Notepad opened")

        win_click_center(np)
        time.sleep(0.2)

        print("  Typing: " + sentinel)
        paste(sentinel)
        time.sleep(0.4)
        screenshot("np_typed")
        print("  PASS Step2: text typed")

        try: np.set_focus()
        except Exception: pass
        time.sleep(0.2)
        print("  Ctrl+S ...")
        send_keys("^s")
        time.sleep(1.5)
        screenshot("np_save_dlg")

        paste(str(save_path))
        time.sleep(0.3)
        send_keys("{ENTER}")
        time.sleep(1.5)

        for _ in range(3):
            if dismiss():
                time.sleep(0.5)
            else:
                break
        screenshot("np_after_save")

        verify(save_path, sentinel)
        print("PASS TEST 1: Notepad create/write/save OK")

    finally:
        fx.stop()
        close_win("Notepad")
        try:
            save_path.unlink(missing_ok=True)
        except Exception:
            pass


# =========================================================
# TEST 2: WORD -- open RTF / type / Ctrl+S / verify
# =========================================================
def test_word_open_write_save():
    print("\n" + "=" * 60)
    print("TEST 2: WORD -- open / write / save")
    print("=" * 60)

    if not find_word():
        pytest.skip("Word not installed")

    fx = App()
    sentinel = "WORD_SAVE_2026"
    save_path = DOCS / ("test_word_" + datetime.now().strftime("%H%M%S") + ".rtf")
    rtf = (r"{\rtf1\ansi\deff0 {\fonttbl{\f0 Times New Roman;}} \f0\fs24 INIT}")
    save_path.write_text(rtf, encoding="ascii")
    print("  Created RTF: " + save_path.name)
    close_win("Word")
    time.sleep(0.5)

    try:
        ok = fx.start()
        if ok and fx.win:
            send_cmd(fx.win, "open " + save_path.name)
            time.sleep(1.0)
        if wait_win("Word", 4) is None:
            subprocess.Popen([find_word(), str(save_path)])

        ww = wait_win("Word", 15)
        assert ww is not None, "FAIL: Word did not open"
        screenshot("word_open")
        print("  PASS Step1: Word opened -- " + ww.window_text())

        time.sleep(1.5)
        send_keys("{ESC}")
        time.sleep(0.5)
        click_body(ww)
        screenshot("word_focused")

        send_keys("^{END}")
        time.sleep(0.2)
        send_keys("{ENTER}")
        time.sleep(0.1)
        print("  Typing: " + sentinel)
        paste(sentinel)
        time.sleep(0.6)
        screenshot("word_typed")
        print("  PASS Step2: text typed in Word")

        ww.set_focus()
        time.sleep(0.2)
        print("  Ctrl+S ...")
        send_keys("^s")
        time.sleep(2.5)
        screenshot("word_save_dlg")

        for _ in range(5):
            if dismiss():
                time.sleep(0.6)
                screenshot("word_dlg_" + str(_))
            else:
                break
        screenshot("word_after_save")

        print("  Closing Word ...")
        close_win("Word")
        time.sleep(2.0)
        for _ in range(3):
            if dismiss():
                time.sleep(0.5)
            else:
                break
        close_win("Word")
        time.sleep(1.0)

        verify(save_path, sentinel)
        print("PASS TEST 2: Word open/write/save OK")

    finally:
        fx.stop()
        close_win("Word")
        time.sleep(0.5)
        try:
            save_path.unlink(missing_ok=True)
        except Exception:
            pass


# =========================================================
# TEST 3: NOTEPAD via App -- anti-flood check
# =========================================================
def test_notepad_no_flood():
    print("\n" + "=" * 60)
    print("TEST 3: NOTEPAD via App -- no flooding")
    print("=" * 60)

    fx = App()
    sentinel = "NO_FLOOD_2026"
    save_path = DOCS / ("test_flood_" + datetime.now().strftime("%H%M%S") + ".txt")
    close_win("Notepad")
    time.sleep(0.4)

    try:
        if not fx.start():
            pytest.skip("AICompanion.exe not available")

        send_cmd(fx.win, "open notepad")
        np = wait_win("Notepad", 12)
        assert np is not None, "FAIL: Notepad not opened via App"
        screenshot("flood_open")
        print("  PASS: Notepad via App")

        win_click_center(np)
        time.sleep(0.2)
        paste(sentinel)
        time.sleep(0.4)
        screenshot("flood_typed")

        before = count_dlg()
        print("  Dialogs before: " + str(before))

        np.set_focus()
        send_keys("^s")
        time.sleep(1.5)
        screenshot("flood_save_dlg")

        paste(str(save_path))
        send_keys("{ENTER}")
        time.sleep(1.5)
        dismiss()
        time.sleep(0.5)
        screenshot("flood_after_save")

        after = count_dlg()
        print("  Dialogs after: " + str(after))
        assert after <= 1, \
            "FLOODING DETECTED: " + str(after) + " dialogs still open!"

        verify(save_path, sentinel)
        print("PASS TEST 3: no flooding")

    finally:
        fx.stop()
        close_win("Notepad")
        try:
            save_path.unlink(missing_ok=True)
        except Exception:
            pass


# =========================================================
# TEST 4: WORD new doc via App -- create / write / save
# =========================================================
def test_word_new_doc_via_app():
    print("\n" + "=" * 60)
    print("TEST 4: WORD new doc via App -- create/write/save")
    print("=" * 60)

    if not find_word():
        pytest.skip("Word not installed")

    fx = App()
    sentinel = "WORD_NEW_2026"
    save_path = DOCS / ("test_word_new_" + datetime.now().strftime("%H%M%S") + ".rtf")
    close_win("Word")
    time.sleep(0.5)

    try:
        if not fx.start():
            pytest.skip("AICompanion.exe not available")

        print("  Sending: create a new word document")
        send_cmd(fx.win, "create a new word document")
        screenshot("newdoc_cmd")

        ww = wait_win("Word", 20)
        assert ww is not None, \
            "FAIL: Word not opened via 'create a new word document'"
        screenshot("newdoc_open")
        print("  PASS Step1: Word open -- " + ww.window_text())

        time.sleep(2.0)
        send_keys("{ESC}")
        time.sleep(0.5)
        click_body(ww)
        screenshot("newdoc_focus")

        send_keys("^{END}")
        time.sleep(0.2)
        print("  Typing: " + sentinel)
        paste(sentinel)
        time.sleep(0.6)
        screenshot("newdoc_typed")
        print("  PASS Step2: text typed")

        ww.set_focus()
        time.sleep(0.2)
        print("  Ctrl+S ...")
        send_keys("^s")
        time.sleep(2.5)
        screenshot("newdoc_save_dlg")

        saveas_dialog(save_path)
        time.sleep(1.5)
        for _ in range(5):
            if dismiss():
                time.sleep(0.5)
            else:
                break
        screenshot("newdoc_after_save")

        print("  Closing Word ...")
        close_win("Word")
        time.sleep(2.0)
        for _ in range(3):
            if dismiss():
                time.sleep(0.5)
            else:
                break
        close_win("Word")
        time.sleep(1.0)

        verify(save_path, sentinel)
        print("PASS TEST 4: Word new doc via App OK")

    finally:
        fx.stop()
        close_win("Word")
        time.sleep(0.5)
        for f in DOCS.glob("test_word_new_*"):
            try:
                f.unlink()
            except Exception:
                pass


if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-v", "-s", "--tb=short"]))
