# AICompanion Test Checklist

## Pre-Test Setup
- [ ] Build solution successfully (no errors)
- [ ] Ensure Watson API key is configured (or test fallback mode)
- [ ] Close all applications that will be opened during testing
- [ ] Set AICompanion window to "Always on Top" (default)

---

## Test Suite 1: Application Control

### TC-1.1: Open Notepad
**Command:** "Open notepad"  
**Expected:** Notepad opens within 2 seconds  
**Toast:** "Opened Notepad" (green)  
- [ ] Pass / Fail

### TC-1.2: Open Calculator
**Command:** "Open calculator"  
**Expected:** Calculator opens within 2 seconds  
**Toast:** "Opened Calculator" (green)  
- [ ] Pass / Fail

### TC-1.3: Open Word
**Command:** "Open Word"  
**Expected:** Microsoft Word opens (if installed)  
**Toast:** "Opened Microsoft Word" (green)  
- [ ] Pass / Fail

### TC-1.4: Open Browser
**Command:** "Open browser"  
**Expected:** Microsoft Edge opens  
**Toast:** "Opened Microsoft Edge" (green)  
- [ ] Pass / Fail

### TC-1.5: Close Window
**Command:** "Close window" (with Notepad focused)  
**Expected:** Notepad closes  
- [ ] Pass / Fail

---

## Test Suite 2: Document Typing

### TC-2.1: Type Text in Notepad
**Setup:** Open Notepad, click inside text area  
**Command:** "Type: Hello world"  
**Expected:** "Hello world" appears in Notepad  
**Toast:** "Typed: Hello world" (green)  
- [ ] Pass / Fail

### TC-2.2: Type with Special Characters
**Setup:** Notepad focused  
**Command:** "Type: Test 123!"  
**Expected:** "Test 123!" appears  
- [ ] Pass / Fail

### TC-2.3: Type Multiple Lines
**Setup:** Notepad focused  
**Commands:** 
1. "Type: Line one"
2. "New line"
3. "Type: Line two"  
**Expected:** Two lines of text  
- [ ] Pass / Fail

### TC-2.4: No Focus Error
**Setup:** Click on desktop (no text field focused)  
**Command:** "Type: This should fail"  
**Expected:** Error toast (red) "No active window" or similar  
- [ ] Pass / Fail

---

## Test Suite 3: Text Editing

### TC-3.1: Select All
**Setup:** Notepad with text, focused  
**Command:** "Select all"  
**Expected:** All text selected (highlighted)  
**Toast:** "Selected all text" (green)  
- [ ] Pass / Fail

### TC-3.2: Copy
**Setup:** Text selected in Notepad  
**Command:** "Copy"  
**Expected:** Text copied to clipboard  
**Toast:** "Copied to clipboard" (green)  
- [ ] Pass / Fail

### TC-3.3: Paste
**Setup:** Clipboard has text, cursor in Notepad  
**Command:** "Paste"  
**Expected:** Clipboard text inserted  
**Toast:** "Pasted from clipboard" (green)  
- [ ] Pass / Fail

### TC-3.4: Cut
**Setup:** Text selected in Notepad  
**Command:** "Cut"  
**Expected:** Text removed and copied to clipboard  
- [ ] Pass / Fail

### TC-3.5: Undo
**Setup:** After typing text in Notepad  
**Command:** "Undo"  
**Expected:** Last action undone  
**Toast:** "Undone" (green)  
- [ ] Pass / Fail

### TC-3.6: Redo
**Setup:** After undo  
**Command:** "Redo"  
**Expected:** Undo reversed  
- [ ] Pass / Fail

---

## Test Suite 4: Formatting (Word)

### TC-4.1: Bold
**Setup:** Word with text selected  
**Command:** "Bold"  
**Expected:** Selected text becomes bold  
**Toast:** "Applied bold formatting" (green)  
- [ ] Pass / Fail

### TC-4.2: Italic
**Setup:** Word with text selected  
**Command:** "Italic"  
**Expected:** Selected text becomes italic  
- [ ] Pass / Fail

### TC-4.3: Underline
**Setup:** Word with text selected  
**Command:** "Underline"  
**Expected:** Selected text becomes underlined  
- [ ] Pass / Fail

---

## Test Suite 5: Navigation

### TC-5.1: Go to Start
**Setup:** Notepad with cursor at end  
**Command:** "Go to start"  
**Expected:** Cursor moves to beginning  
- [ ] Pass / Fail

### TC-5.2: Go to End
**Setup:** Notepad with cursor at start  
**Command:** "Go to end"  
**Expected:** Cursor moves to end  
- [ ] Pass / Fail

### TC-5.3: Save Document
**Setup:** Notepad with unsaved changes  
**Command:** "Save"  
**Expected:** Save dialog appears (or file saved if already named)  
- [ ] Pass / Fail

---

## Test Suite 6: Voice Recognition

### TC-6.1: Watson STT
**Setup:** Watson API key configured  
**Action:** Hold mic button, speak "Open notepad"  
**Expected:** Text appears in UI, command executes  
- [ ] Pass / Fail

### TC-6.2: Windows Speech Fallback
**Setup:** Remove Watson API key or disconnect internet  
**Action:** Hold mic button, speak "Open calculator"  
**Expected:** Falls back to Windows Speech, command executes  
- [ ] Pass / Fail

### TC-6.3: Interim Results
**Setup:** Watson configured  
**Action:** Hold mic and speak slowly  
**Expected:** Partial text appears while speaking  
- [ ] Pass / Fail

---

## Test Suite 7: Teaching Mode

### TC-7.1: How Do I
**Command:** "How do I copy text?"  
**Expected:** AI explains the copy process via TTS  
- [ ] Pass / Fail

### TC-7.2: Show Commands
**Command:** "What commands"  
**Expected:** List of available commands displayed/spoken  
- [ ] Pass / Fail

### TC-7.3: Tutorial Start
**Command:** "Start tutorial"  
**Expected:** Tutorial begins with first step  
- [ ] Pass / Fail

---

## Test Suite 8: UI Feedback

### TC-8.1: Toast Appears
**Action:** Execute any command  
**Expected:** Toast notification appears at top of window  
- [ ] Pass / Fail

### TC-8.2: Toast Auto-Hides
**Action:** Wait 2.5 seconds after toast appears  
**Expected:** Toast fades out  
- [ ] Pass / Fail

### TC-8.3: Activity Log
**Action:** Execute multiple commands  
**Expected:** Commands appear in activity log with timestamps  
- [ ] Pass / Fail

### TC-8.4: Avatar Expression
**Action:** Execute successful command  
**Expected:** Avatar shows happy expression  
- [ ] Pass / Fail

### TC-8.5: Error Expression
**Action:** Execute failing command  
**Expected:** Avatar shows sad expression  
- [ ] Pass / Fail

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
- [ ] Pass / Fail

---

## Diagnostic Logs

### Transcript Log Location
```
%LOCALAPPDATA%\AICompanion\Logs\watson_transcripts_YYYYMMDD.log
```

### Check Transcript Log
After testing, verify log contains entries:
- [ ] Log file exists
- [ ] Entries show timestamp, confidence, and transcript text
- [ ] Final vs interim results are marked

---

## Summary

| Suite | Passed | Failed | Total |
|-------|--------|--------|-------|
| Application Control | | | 5 |
| Document Typing | | | 4 |
| Text Editing | | | 6 |
| Formatting | | | 3 |
| Navigation | | | 3 |
| Voice Recognition | | | 3 |
| Teaching Mode | | | 3 |
| UI Feedback | | | 5 |
| Cross-Application | | | 1 |
| **TOTAL** | | | **33** |

---

**Tester:** _______________  
**Date:** _______________  
**Build Version:** _______________  
**Notes:**
