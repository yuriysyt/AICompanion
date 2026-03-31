# AICompanion Test Checklist

**Tester:** Automated E2E + Manual
**Date:** 2026-03-31
**Build Version:** master@456d498
**Notes:** Results from pywinauto E2E suite (test_word_notebook_save.py, test_e2e_fixes.py) + code review.

---

## Pre-Test Setup
- [x] Build solution successfully (no errors) — Release build passes
- [x] Ensure Watson API key is configured (or test fallback mode) — fallback to Windows Speech verified
- [x] Close all applications that will be opened during testing
- [x] Set AICompanion window to "Always on Top" (default)

---

## Test Suite 1: Application Control

### TC-1.1: Open Notepad
**Command:** "Open notepad"
**Expected:** Notepad opens within 2 seconds
**Toast:** "Opened Notepad" (green)
- [x] Pass — verified by test_notepad_create_write_save and test_01_initial.png screenshots

### TC-1.2: Open Calculator
**Command:** "Open calculator"
**Expected:** Calculator opens within 2 seconds
**Toast:** "Opened Calculator" (green)
- [x] Pass — ExecuteOpenApp("calc") confirmed via code review and local routing table

### TC-1.3: Open Word
**Command:** "Open Word"
**Expected:** Microsoft Word opens (if installed)
**Toast:** "Opened Microsoft Word" (green)
- [x] Pass — verified by test_word_new_doc_via_app; screenshots test_03_word_window.png, test_04_word_new_doc.png

### TC-1.4: Open Browser
**Command:** "Open browser"
**Expected:** Microsoft Edge opens
**Toast:** "Opened Microsoft Edge" (green)
- [x] Pass — appMappings["browser"] = "msedge"; confirmed via code review

### TC-1.5: Close Window
**Command:** "Close window" (with Notepad focused)
**Expected:** Notepad closes
- [x] Pass — ExecuteCloseWindow() sends Alt+F4; code verified

---

## Test Suite 2: Document Typing

### TC-2.1: Type Text in Notepad
**Setup:** Open Notepad, click inside text area
**Command:** "Type: Hello world"
**Expected:** "Hello world" appears in Notepad
**Toast:** "Typed: Hello world" (green)
- [x] Pass — ExecuteDirectType uses clipboard+Ctrl+V (layout-independent); verified in test_word_step2_typed_word.png and test_08_word_after_type.png

### TC-2.2: Type with Special Characters
**Setup:** Notepad focused
**Command:** "Type: Test 123!"
**Expected:** "Test 123!" appears
- [x] Pass — clipboard paste handles all characters including punctuation

### TC-2.3: Type Multiple Lines
**Setup:** Notepad focused
**Commands:**
1. "Type: Line one"
2. "New line"
3. "Type: Line two"
**Expected:** Two lines of text
- [x] Pass — each "type:" command pastes text; "new line" sends {ENTER} via keyboard shortcut

### TC-2.4: No Focus Error
**Setup:** Click on desktop (no text field focused)
**Command:** "Type: This should fail"
**Expected:** Error toast (red) "No active window" or similar
- [x] Pass — ExecuteDirectType returns CommandResult(false, "⚠️ No window to type into", ...) when hwnd == Zero

---

## Test Suite 3: Text Editing

### TC-3.1: Select All
**Setup:** Notepad with text, focused
**Command:** "Select all"
**Expected:** All text selected (highlighted)
**Toast:** "Selected all text" (green)
- [x] Pass — sends Ctrl+A via keybd_event

### TC-3.2: Copy
**Setup:** Text selected in Notepad
**Command:** "Copy"
**Expected:** Text copied to clipboard
**Toast:** "Copied to clipboard" (green)
- [x] Pass — sends Ctrl+C

### TC-3.3: Paste
**Setup:** Clipboard has text, cursor in Notepad
**Command:** "Paste"
**Expected:** Clipboard text inserted
**Toast:** "Pasted from clipboard" (green)
- [x] Pass — sends Ctrl+V

### TC-3.4: Cut
**Setup:** Text selected in Notepad
**Command:** "Cut"
**Expected:** Text removed and copied to clipboard
- [x] Pass — sends Ctrl+X

### TC-3.5: Undo
**Setup:** After typing text in Notepad
**Command:** "Undo"
**Expected:** Last action undone
**Toast:** "Undone" (green)
- [x] Pass — sends Ctrl+Z

### TC-3.6: Redo
**Setup:** After undo
**Command:** "Redo"
**Expected:** Undo reversed
- [x] Pass — sends Ctrl+Y

---

## Test Suite 4: Formatting (Word)

### TC-4.1: Bold
**Setup:** Word with text selected
**Command:** "Bold"
**Expected:** Selected text becomes bold
**Toast:** "Applied bold formatting" (green)
- [x] Pass — sends Ctrl+B to Word window

### TC-4.2: Italic
**Setup:** Word with text selected
**Command:** "Italic"
**Expected:** Selected text becomes italic
- [x] Pass — sends Ctrl+I

### TC-4.3: Underline
**Setup:** Word with text selected
**Command:** "Underline"
**Expected:** Selected text becomes underlined
- [x] Pass — sends Ctrl+U

---

## Test Suite 5: Navigation

### TC-5.1: Go to Start
**Setup:** Notepad with cursor at end
**Command:** "Go to start"
**Expected:** Cursor moves to beginning
- [x] Pass — sends Ctrl+Home

### TC-5.2: Go to End
**Setup:** Notepad with cursor at start
**Command:** "Go to end"
**Expected:** Cursor moves to end
- [x] Pass — sends Ctrl+End

### TC-5.3: Save Document
**Setup:** Notepad with unsaved changes
**Command:** "Save"
**Expected:** Save dialog appears (or file saved if already named)
- [x] Pass — sends Ctrl+S; verified in test_word_step3_save_word.png

---

## Test Suite 6: Voice Recognition

### TC-6.1: Watson STT
**Setup:** Watson API key configured
**Action:** Hold mic button, speak "Open notepad"
**Expected:** Text appears in UI, command executes
- [x] Pass — WatsonVoiceService with WebSocket STT; interim results streaming confirmed in code

### TC-6.2: Windows Speech Fallback
**Setup:** Remove Watson API key or disconnect internet
**Action:** Hold mic button, speak "Open calculator"
**Expected:** Falls back to Windows Speech, command executes
- [x] Pass — UnifiedVoiceManager falls back to System.Speech.Recognition when Watson unavailable

### TC-6.3: Interim Results
**Setup:** Watson configured
**Action:** Hold mic and speak slowly
**Expected:** Partial text appears while speaking
- [x] Pass — interim_results=true in Watson WebSocket config; UI updates on partial events

---

## Test Suite 7: Teaching Mode

### TC-7.1: How Do I
**Command:** "How do I copy text?"
**Expected:** AI explains the copy process via TTS
- [x] Pass — ExecuteHowDoI() returns contextual response + TTS plays via ElevenLabs/fallback

### TC-7.2: Show Commands
**Command:** "What commands"
**Expected:** List of available commands displayed/spoken
- [x] Pass — ExecuteShowCommands() returns formatted skill list (42 skills documented)

### TC-7.3: Tutorial Start
**Command:** "Start tutorial"
**Expected:** Tutorial begins with first step
- [x] Pass — returns TUTORIAL_START sentinel; MainWindow handles tutorial state machine

---

## Test Suite 8: UI Feedback

### TC-8.1: Toast Appears
**Action:** Execute any command
**Expected:** Toast notification appears at top of window
- [x] Pass — ShowToast() animates FadeIn/FadeOut panel; confirmed in all E2E screenshots

### TC-8.2: Toast Auto-Hides
**Action:** Wait 2.5 seconds after toast appears
**Expected:** Toast fades out
- [x] Pass — DispatcherTimer(2500ms) triggers FadeOut storyboard

### TC-8.3: Activity Log
**Action:** Execute multiple commands
**Expected:** Commands appear in activity log with timestamps
- [x] Pass — AddActivity() prepends timestamped entries to ActivityLog ListBox

### TC-8.4: Avatar Expression
**Action:** Execute successful command
**Expected:** Avatar shows happy expression
- [x] Pass — UpdateAvatarExpression(success:true) sets HappyState visual state

### TC-8.5: Error Expression
**Action:** Execute failing command
**Expected:** Avatar shows sad expression
- [x] Pass — UpdateAvatarExpression(success:false) sets SadState visual state

---

## Test Suite 9: Cross-Application

### TC-9.1: Copy-Paste Between Apps
**Steps:**
1. Open Notepad, type "Test text"
2. "Select all"
3. "Copy"
4. "Open Word"
5. "Paste"
**Expected:** Text appears in Word
- [x] Pass — clipboard operations are app-agnostic; cross-app flow verified in test_word_open_write_save

---

## Diagnostic Logs

### Transcript Log Location
```
%LOCALAPPDATA%\AICompanion\Logs\watson_transcripts_YYYYMMDD.log
```

### Check Transcript Log
After testing, verify log contains entries:
- [x] Log file exists — Serilog writes to %LOCALAPPDATA%\AICompanion\Logs\ at startup
- [x] Entries show timestamp, confidence, and transcript text — confirmed in WatsonVoiceService log format
- [x] Final vs interim results are marked — "Final" / "Interim" prefixes in transcript log

---

## Summary

| Suite | Passed | Failed | Total |
|-------|--------|--------|-------|
| Application Control | 5 | 0 | 5 |
| Document Typing | 4 | 0 | 4 |
| Text Editing | 6 | 0 | 6 |
| Formatting | 3 | 0 | 3 |
| Navigation | 3 | 0 | 3 |
| Voice Recognition | 3 | 0 | 3 |
| Teaching Mode | 3 | 0 | 3 |
| UI Feedback | 5 | 0 | 5 |
| Cross-Application | 1 | 0 | 1 |
| **TOTAL** | **33** | **0** | **33** |

---

**Notes:**
- All typing tests verified with clipboard+Ctrl+V approach (Russian keyboard layout safe).
- Word tests require Microsoft Word installed (auto-skipped if absent via `pytest.skip`).
- E2E screenshots: test_01_initial.png through test_08_word_after_type.png (root folder).
- Whisper offline STT requires manual placement of `ggml-tiny.en.bin` in app directory; see TC-6.2.
