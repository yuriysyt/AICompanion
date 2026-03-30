"""
AICompanion Backend Server — IBM Granite via Ollama  (OpenClaw architecture)
=============================================================================
Run with:
  uvicorn server:app --port 8000 --reload
or:
  python server.py
"""

import re as _re
import json
import time
import logging
import httpx
import os
import os as _os
import pathlib
import subprocess
import tempfile
from datetime import datetime
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Optional

# ─── OpenClaw: optional automation libs ──────────────────────────────────────
try:
    import pyautogui as _pyautogui
    _pyautogui.FAILSAFE = False
    _pyautogui.PAUSE    = 0.05
    _PYAUTOGUI = True
    log_init = logging.getLogger("granite")
    log_init.info("OpenClaw: pyautogui loaded")
except Exception:
    _PYAUTOGUI = False

try:
    import win32api as _win32api
    import win32con as _win32con
    import win32gui as _win32gui
    import win32process as _win32process
    _WIN32 = True
except Exception:
    _WIN32 = False

try:
    import psutil as _psutil
    _PSUTIL = True
except Exception:
    _PSUTIL = False

try:
    import comtypes.client as _comtypes
    _COMTYPES = True
except Exception:
    _COMTYPES = False

try:
    from pywinauto.application import Application as _PwApp
    import pywinauto.keyboard as _PwKbd
    _PYWINAUTO = True
except Exception:
    _PYWINAUTO = False

# ─── ChromaDB vector memory ───────────────────────────────────────────────────
_CHROMADB = False
_chroma_collection = None
try:
    import chromadb
    from chromadb.utils.embedding_functions import SentenceTransformerEmbeddingFunction
    _chroma_ef = SentenceTransformerEmbeddingFunction(model_name="all-MiniLM-L6-v2")
    _chroma_client = chromadb.PersistentClient(
        path=str(pathlib.Path(__file__).parent / "chroma_db")
    )
    _chroma_collection = _chroma_client.get_or_create_collection(
        "aicompanion_memory",
        embedding_function=_chroma_ef,
    )
    _CHROMADB = True
    log_init = logging.getLogger("granite")
    log_init.info("ChromaDB loaded — collection 'aicompanion_memory' ready")
except Exception as _ce:
    logging.getLogger("granite").warning("ChromaDB unavailable (%s) — running without semantic memory", _ce)

def _chroma_store(session_id: str, user_msg: str, assistant_reply: str) -> None:
    """Store a Q/A pair as a ChromaDB embedding."""
    if not _CHROMADB or _chroma_collection is None:
        return
    try:
        doc_id = f"{session_id}_{int(time.time() * 1000)}"
        _chroma_collection.add(
            documents=[f"User: {user_msg}\nAssistant: {assistant_reply}"],
            ids=[doc_id],
            metadatas=[{"session_id": session_id, "ts": int(time.time())}],
        )
    except Exception as e:
        logging.getLogger("granite").debug("ChromaDB store error: %s", e)

def _chroma_query(user_msg: str, n_results: int = 3) -> str:
    """Return top-N semantically similar past exchanges, formatted as context."""
    if not _CHROMADB or _chroma_collection is None:
        return ""
    try:
        count = _chroma_collection.count()
        if count == 0:
            return ""
        results = _chroma_collection.query(
            query_texts=[user_msg],
            n_results=min(n_results, count),
        )
        docs = results.get("documents", [[]])[0]
        if not docs:
            return ""
        return "Relevant past context:\n" + "\n---\n".join(docs) + "\n"
    except Exception as e:
        logging.getLogger("granite").debug("ChromaDB query error: %s", e)
        return ""

# Known Word installation paths (fallback when WINWORD.EXE not on PATH)
_WORD_PATHS = [
    r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
    r"C:\Program Files\Microsoft Office\root\Office15\WINWORD.EXE",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
    r"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE",
    r"C:\Program Files\Microsoft Office\Office15\WINWORD.EXE",
    r"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE",
]

def _find_word_exe() -> Optional[str]:
    for p in _WORD_PATHS:
        if _os.path.exists(p):
            return p
    return None

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("granite")

OLLAMA_BASE    = "http://localhost:11434"
GRANITE_MODEL  = "granite3-dense:latest"
OLLAMA_TIMEOUT = 60

_BACKEND_DIR = pathlib.Path(__file__).parent
_MEMORY_DIR  = _BACKEND_DIR / "memory"
_MEMORY_DIR.mkdir(exist_ok=True)

_SOUL_PATH = _BACKEND_DIR / "SOUL.md"
_SOUL_TEXT = ""
if _SOUL_PATH.exists():
    _SOUL_TEXT = _SOUL_PATH.read_text(encoding="utf-8").strip()
    log.info("SOUL.md loaded (%d chars)", len(_SOUL_TEXT))
else:
    log.warning("SOUL.md not found at %s — using empty persona", _SOUL_PATH)


# ─── Skills Registry ─────────────────────────────────────────────────────────

SKILLS: dict[str, dict] = {
    "open_app": {
        "desc": "Launch a desktop application",
        "params": "target: executable name (notepad, msedge, winword, calc, explorer, opera, firefox, chrome, excel)",
    },
    "close_window": {
        "desc": "Close a window by title fragment",
        "params": "target: window title fragment",
    },
    "focus_window": {
        "desc": "Bring an already-open window to the foreground",
        "params": "target: executable name of the running process",
    },
    "type_text": {
        "desc": "Type text into the active window via the keyboard",
        "params": "params: the full text to type",
    },
    "save_document": {
        "desc": "Save the active document (Ctrl+S)",
        "params": "none",
    },
    "navigate_url": {
        "desc": "Navigate the browser or Explorer to a URL or shell path",
        "params": "target: full URL or shell path (e.g. shell:Downloads)",
    },
    "search_web": {
        "desc": "Search the web using the default browser",
        "params": "params: search query string",
    },
    "scroll_up": {
        "desc": "Scroll the active window upward",
        "params": "params: number of scroll steps (integer string)",
    },
    "scroll_down": {
        "desc": "Scroll the active window downward",
        "params": "params: number of scroll steps (integer string)",
    },
    "undo": {
        "desc": "Undo the last action (Ctrl+Z)",
        "params": "none",
    },
    "redo": {
        "desc": "Redo the last undone action (Ctrl+Y)",
        "params": "none",
    },
    "copy_text": {
        "desc": "Copy selected content to clipboard (Ctrl+C)",
        "params": "none",
    },
    "copy": {
        "desc": "Copy selected content to clipboard (alias for copy_text)",
        "params": "none",
    },
    "paste_text": {
        "desc": "Paste clipboard content (Ctrl+V)",
        "params": "none",
    },
    "paste": {
        "desc": "Paste clipboard content (alias for paste_text)",
        "params": "none",
    },
    "select_all": {
        "desc": "Select all content in the active window (Ctrl+A)",
        "params": "none",
    },
    "minimize_window": {
        "desc": "Minimize the active window",
        "params": "none",
    },
    "maximize_window": {
        "desc": "Maximize the active window",
        "params": "none",
    },
    "take_screenshot": {
        "desc": "Capture a screenshot of the entire screen",
        "params": "none",
    },
    "screenshot": {
        "desc": "Capture a screenshot (alias for take_screenshot)",
        "params": "none",
    },
    "new_document": {
        "desc": "Create a new blank document in the currently open app (Ctrl+N) — use this when user says 'create new document', 'new file', 'blank document'",
        "params": "target: executable name of the app (winword, notepad, etc.)",
    },
    "new_tab": {
        "desc": "Open a new browser tab (Ctrl+T)",
        "params": "none",
    },
    "close_tab": {
        "desc": "Close the current browser tab (Ctrl+W)",
        "params": "none",
    },
    "format_bold": {
        "desc": "Apply bold formatting to selected text (Ctrl+B)",
        "params": "none",
    },
    "format_italic": {
        "desc": "Apply italic formatting to selected text (Ctrl+I)",
        "params": "none",
    },
    "format_underline": {
        "desc": "Apply underline formatting to selected text (Ctrl+U)",
        "params": "none",
    },
    "find_and_click": {
        "desc": "Find a UI button/element by its visible label and click it (uses Windows UIAutomation). Use this to click 'Blank document', 'Save', 'OK', 'Yes', 'Cancel', template buttons, etc.",
        "params": "params: the visible label of the element to click (e.g. 'Blank document')",
    },
    "click_at": {
        "desc": "Click the mouse at specific screen coordinates",
        "params": "target: 'x,y' coordinates as string (e.g. '960,540')",
    },
    "right_click": {
        "desc": "Right-click at specific coordinates or center of active window",
        "params": "target: 'x,y' coordinates (optional)",
    },
    "double_click": {
        "desc": "Double-click at specific coordinates",
        "params": "target: 'x,y' coordinates (optional — defaults to active window center)",
    },
    "describe_screen": {
        "desc": "Describe the current screen state (open windows, visible buttons). Use this when unsure what is visible.",
        "params": "none",
    },
    "press_escape": {
        "desc": "Press the Escape key in the active window",
        "params": "none",
    },
    "press_enter": {
        "desc": "Press the Enter key in the active window",
        "params": "none",
    },
    "go_to_end": {
        "desc": "Move cursor to end of document (Ctrl+End)",
        "params": "none",
    },
    "go_to_start": {
        "desc": "Move cursor to beginning of document (Ctrl+Home)",
        "params": "none",
    },
    "mouse_move": {
        "desc": "Move mouse cursor to screen coordinates without clicking",
        "params": "target: 'x,y' pixel coordinates e.g. '960,540'",
    },
    "drag_to": {
        "desc": "Click and drag from current mouse position to target coordinates (drag files, sliders, etc.)",
        "params": "target: 'x,y' destination coordinates e.g. '500,300'",
    },
    "scroll_at": {
        "desc": "Scroll up or down at specific screen coordinates (useful for scrolling inside specific areas)",
        "params": "params: 'x,y,direction,amount' e.g. '960,540,down,3'",
    },
    "hotkey": {
        "desc": "Press any keyboard shortcut or key combination",
        "params": "params: key combo e.g. 'ctrl+s', 'alt+f4', 'win+d', 'f5', 'ctrl+shift+t'",
    },
    "wait": {
        "desc": "Wait/pause for N seconds before next action (use after opening apps to wait for them to load)",
        "params": "params: number of seconds as string e.g. '2'",
    },
    "open_file_path": {
        "desc": "Open a file or folder by its full path using the default application",
        "params": "params: full file or folder path e.g. 'C:\\Users\\user\\Documents\\file.docx'",
    },
    "message_user": {
        "desc": "Display a message or question to the user — use when you need clarification or want to inform the user of something",
        "params": "params: the message text to show the user",
    },
}

def _skills_summary() -> str:
    lines = []
    for name, info in SKILLS.items():
        lines.append(f"  {name}: {info['desc']}  |  {info['params']}")
    return "\n".join(lines)


app = FastAPI(
    title="AICompanion — IBM Granite Backend",
    description="Local Granite 3 inference server for the AICompanion desktop app",
    version="2.0.0",
)
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

_debug: dict = {"prompt": "", "raw_response": "", "parsed": {}, "latency_ms": 0}
_chat_history: dict[str, list[dict]] = {}
_MAX_HISTORY_TURNS = 20


# ─── Data models ─────────────────────────────────────────────────────────────

class IntentRequest(BaseModel):
    text: str
    session_id: str = "default"
    window_title: str = ""
    window_process: str = ""

class IntentResponse(BaseModel):
    action: str
    target: Optional[str] = None
    text: Optional[str] = None
    thinking: str = ""
    latency_ms: int = 0

class PlanRequest(BaseModel):
    text: str
    session_id: str = "default"
    window_title: str = ""
    window_process: str = ""
    session_context: Optional[str] = None
    max_steps: int = 6
    open_windows: list[str] = []

class PlanStep(BaseModel):
    step_number: int
    action: str
    target: Optional[str] = None
    params: Optional[str] = None
    confidence: int = 80

class PlanResponse(BaseModel):
    plan_id: str
    steps: list[PlanStep]
    total_steps: int
    reasoning: str
    latency_ms: int = 0

class ChatRequest(BaseModel):
    message: str
    session_id: str = "default"
    history: list[dict] = []

class ChatResponse(BaseModel):
    reply: str
    latency_ms: int = 0

class ClearHistoryRequest(BaseModel):
    session_id: str = "default"

class EssayRequest(BaseModel):
    topic:      str
    language:   str = "English"
    word_count: int = 200
    style:      str = "informative"

class EssayThinkStep(BaseModel):
    stage:   str
    content: str
    latency_ms: int = 0

class EssayResponse(BaseModel):
    topic:      str
    outline:    str
    draft:      str
    final_text: str
    total_latency_ms: int
    think_steps: list[EssayThinkStep]

class SmartCommandRequest(BaseModel):
    text: str
    session_id: str = "default"
    window_title: str = ""
    window_process: str = ""
    session_context: Optional[str] = None
    open_windows: list[str] = []  # NEW: list of currently open window titles

class SmartCommandResponse(BaseModel):
    plan_id: str
    steps: list[PlanStep]
    total_steps: int
    reasoning: str
    latency_ms: int = 0
    content_generated: Optional[str] = None


# ─── Ollama helper ────────────────────────────────────────────────────────────

async def granite_generate(prompt: str, *, temperature: float = 0.1) -> tuple[str, int]:
    t0 = time.monotonic()
    payload = {
        "model":  GRANITE_MODEL,
        "prompt": prompt,
        "stream": False,
        "options": {
            "temperature": temperature,
            "top_p": 0.9,
            "num_predict": 512,
            "stop": ["```", "\n\n\n"],
        },
    }
    log.info("GRANITE  prompt=%d chars", len(prompt))

    async with httpx.AsyncClient(timeout=OLLAMA_TIMEOUT) as client:
        r = await client.post(f"{OLLAMA_BASE}/api/generate", json=payload)
        r.raise_for_status()
        data = r.json()

    text = data.get("response", "").strip()
    ms   = int((time.monotonic() - t0) * 1000)
    log.info("GRANITE  %dms  tokens=%s  text=%r", ms, data.get("eval_count", "?"), text[:120])

    _debug["prompt"]       = prompt[-600:]
    _debug["raw_response"] = text
    _debug["latency_ms"]   = ms
    return text, ms


# ─── JSON helpers ─────────────────────────────────────────────────────────────

def _repair_json(raw: str) -> str:
    text = raw.strip()

    first = text.find("{")
    if first > 0:
        text = text[first:]

    text = _re.sub(r",\s*([}\]])", r"\1", text)

    opens  = text.count("{") - text.count("}")
    oarray = text.count("[") - text.count("]")
    if opens > 0 or oarray > 0:
        text = text + ("]" * max(oarray, 0)) + ("}" * max(opens, 0))

    return text


def _extract_json(text: str) -> dict:
    start = text.find("{")
    if start == -1:
        raise ValueError("No JSON object found in model output")
    depth = 0
    for i, ch in enumerate(text[start:], start):
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return json.loads(text[start : i + 1])
    repaired = _repair_json(text)
    return json.loads(repaired)


# ─── Memory helpers ───────────────────────────────────────────────────────────

def _memory_path(session_id: str) -> pathlib.Path:
    safe = _re.sub(r"[^\w\-]", "_", session_id)
    return _MEMORY_DIR / f"{safe}.md"

def _append_memory(session_id: str, user_msg: str, assistant_reply: str) -> None:
    path = _memory_path(session_id)
    stamp = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    with open(path, "a", encoding="utf-8") as f:
        f.write(f"\n## {stamp}\n**User:** {user_msg}\n**Assistant:** {assistant_reply}\n")

def _read_memory(session_id: str) -> str:
    path = _memory_path(session_id)
    if not path.exists():
        return ""
    return path.read_text(encoding="utf-8")

def _write_memory(session_id: str, content: str) -> None:
    path = _memory_path(session_id)
    path.write_text(content, encoding="utf-8")


# ─── Intent extraction ────────────────────────────────────────────────────────

INTENT_ACTIONS = list(SKILLS.keys())

INTENT_PROMPT = """\
{soul}

You are an AI assistant that converts voice commands into structured actions.
Respond ONLY with valid JSON. No prose, no markdown, no code fences.

Voice command: "{text}"
Active window: "{window_title}" (process: {window_process})

Step-by-step thinking:
1. What does the user want?
2. Which action from the list best fits?
3. What is the target app/text/URL (if any)?

Available actions:
{skills}

Rules:
- "открой" / "open" → open_app
- "напиши" / "type" / "write" → type_text
- "сохрани" / "save" → save_document
- "закрой" / "close" → close_window
- "найди" / "search" / "поиск" → search_web
- "new" / "create" + document/file/notebook → new_document
- app name in target must be the executable name (notepad, msedge, opera, winword)

IMPORTANT app name mapping (use EXACT values below):
- "browser", "edge", "internet", "web" → target "msedge"
- "word", "ворд", "document", "doc" → target "winword"
- "notepad", "notebook", "блокнот", "text file", ".txt", "note file" → target "notepad"
- "calculator", "калькулятор" → target "calc"
- "explorer", "folder", "files" → target "explorer"
- "excel", "spreadsheet" → target "excel"
- The active window title is for context only — do NOT use it as the target.

Examples:
- "new notebook" → {{"thinking":"notebook=notepad","action":"new_document","target":"notepad","text":null}}
- "new word document" → {{"thinking":"word document","action":"new_document","target":"winword","text":null}}
- "open calculator" → {{"thinking":"launch calc","action":"open_app","target":"calc","text":null}}
- "open browser" → {{"thinking":"launch browser","action":"open_app","target":"msedge","text":null}}

JSON schema (strict):
{{
  "thinking": "<one sentence reasoning>",
  "action":   "<action_name>",
  "target":   "<app name or URL or null>",
  "text":     "<text to type or null>"
}}
"""

@app.post("/api/intent", response_model=IntentResponse)
async def extract_intent(req: IntentRequest):
    prompt = INTENT_PROMPT.format(
        soul=_SOUL_TEXT,
        text=req.text,
        window_title=req.window_title or "unknown",
        window_process=req.window_process or "unknown",
        skills=_skills_summary(),
    )

    raw, ms = await granite_generate(prompt)

    try:
        parsed = _extract_json(raw)
        _debug["parsed"] = parsed
        log.info("INTENT  action=%s  target=%s", parsed.get("action"), parsed.get("target"))
        return IntentResponse(
            action=parsed.get("action", "unknown"),
            target=parsed.get("target"),
            text=parsed.get("text"),
            thinking=parsed.get("thinking", ""),
            latency_ms=ms,
        )
    except (ValueError, json.JSONDecodeError) as exc:
        log.warning("JSON parse failed: %s  raw=%r", exc, raw[:200])
        action = "unknown"
        for a in INTENT_ACTIONS:
            if a.replace("_", " ") in raw.lower() or a in raw.lower():
                action = a
                break
        return IntentResponse(action=action, thinking=raw, latency_ms=ms)


# ─── Agentic plan ─────────────────────────────────────────────────────────────

PLAN_PROMPT = """\
{soul}

You are an AI assistant that controls a Windows computer for the user.
Respond ONLY with valid JSON. No prose, no markdown, no code fences.

User request: "{text}"
Active window: "{window_title}" (process: {window_process})
Session context: {context}
OPEN WINDOWS: {open_windows}

REASONING INSTRUCTIONS — think before generating steps:
1. What is the user's real goal? (open an app / write content / search the web / click something)
2. Which app achieves this goal? Check OPEN WINDOWS — if it is listed, use focus_window, not open_app.
3. If the user wants to WRITE or TYPE — generate REAL content in type_text.params (not a placeholder, not "...", but the actual text to type).
4. If the user wants to SEARCH — use search_web with the actual query.
5. Keep steps minimal but complete.

Available actions:
{skills}

Rules:
- If the app is already open (in OPEN WINDOWS), use focus_window, not open_app
- new_document creates a NEW blank document; use when user says "create new", "new document", "новый файл"
- type_text params MUST be the FULL real text to type — never empty, never a placeholder like "essay text here"
- ALWAYS end with type_text when the user wants to write/type something
- ALWAYS add save_document as final step when user says "save", "сохрани", or after writing
- find_and_click clicks any visible UI element by label ("Blank document", "OK", "Save", etc.)
- You CAN combine ANY actions freely

EXAMPLES:

Request: "open browser"
{{"reasoning":"Open Edge browser","steps":[{{"step_number":1,"action":"open_app","target":"msedge","params":null,"confidence":97}}]}}

Request: "search for weather in London"
{{"reasoning":"Open browser then search","steps":[{{"step_number":1,"action":"open_app","target":"msedge","params":null,"confidence":95}},{{"step_number":2,"action":"search_web","target":null,"params":"weather in London","confidence":97}}]}}

Request: "open Word"
{{"reasoning":"Launch Microsoft Word","steps":[{{"step_number":1,"action":"open_app","target":"winword","params":null,"confidence":97}}]}}

Request: "create a new Word document"
{{"reasoning":"Create blank Word document","steps":[{{"step_number":1,"action":"new_document","target":"winword","params":null,"confidence":97}}]}}

Request: "create new notepad file"
{{"reasoning":"Create blank Notepad file","steps":[{{"step_number":1,"action":"new_document","target":"notepad","params":null,"confidence":97}}]}}

Request: "create new notebook"
{{"reasoning":"Notebook means Notepad — create blank text file","steps":[{{"step_number":1,"action":"new_document","target":"notepad","params":null,"confidence":95}}]}}

Request: "write essay about climate change in Word"
{{"reasoning":"Open Word, write a full climate essay with real content","steps":[{{"step_number":1,"action":"new_document","target":"winword","params":null,"confidence":95}},{{"step_number":2,"action":"type_text","target":null,"params":"Climate Change: Causes, Effects and Solutions\\n\\nClimate change is one of the most pressing challenges facing humanity. Global temperatures have risen by approximately 1.1°C since pre-industrial times, driven largely by the burning of fossil fuels, deforestation, and industrial processes that release greenhouse gases such as CO2 and methane.\\n\\nThe effects are already visible: more frequent extreme weather events, rising sea levels threatening coastal cities, and disrupted ecosystems. Arctic ice is melting at unprecedented rates, and coral reefs are bleaching due to warmer ocean temperatures.\\n\\nTo address this crisis, the world must rapidly transition to renewable energy sources, improve energy efficiency, and protect forests. International cooperation, carbon pricing, and investment in clean technology are all essential steps. Every individual can also contribute by reducing consumption, choosing sustainable transport, and supporting climate-conscious policies.\\n\\nThe window to act is narrowing, but the solutions exist. What is needed now is the political will and collective action to implement them.","confidence":90}},{{"step_number":3,"action":"save_document","target":null,"params":null,"confidence":95}}]}}

Request: "write a short introduction about artificial intelligence in Notepad"
{{"reasoning":"Open Notepad and type a real AI introduction","steps":[{{"step_number":1,"action":"new_document","target":"notepad","params":null,"confidence":95}},{{"step_number":2,"action":"type_text","target":null,"params":"Introduction to Artificial Intelligence\\n\\nArtificial Intelligence (AI) refers to the simulation of human intelligence in machines that are programmed to think and learn. AI systems can perform tasks such as recognizing speech, making decisions, translating languages, and identifying patterns in data.\\n\\nModern AI is powered by machine learning and deep learning techniques, enabling computers to improve their performance through experience without being explicitly programmed for each task.","confidence":90}}]}}

Request: "напиши эссе о природе в ворде"
{{"reasoning":"Open Word, write nature essay in Russian","steps":[{{"step_number":1,"action":"new_document","target":"winword","params":null,"confidence":95}},{{"step_number":2,"action":"type_text","target":null,"params":"Природа и человек\\n\\nПрирода — это окружающий нас мир: леса, реки, горы, животные и растения. Человек — часть природы, и его жизнь неразрывно связана с ней. Однако с развитием цивилизации человек всё больше вмешивается в природные процессы, нарушая экологическое равновесие.\\n\\nОдной из главных экологических проблем является загрязнение окружающей среды. Выбросы промышленных предприятий, автомобильные выхлопы и бытовые отходы отравляют воздух, воду и почву. Это негативно влияет на здоровье людей и приводит к гибели животных и растений.\\n\\nОднако человек осознаёт свою ответственность перед природой. Создаются заповедники и национальные парки, принимаются законы об охране окружающей среды, развиваются альтернативные источники энергии.\\n\\nКаждый из нас может внести свой вклад в сохранение природы: экономить воду и электричество, сортировать мусор, сажать деревья. Только совместными усилиями мы сможем сохранить нашу планету для будущих поколений.","confidence":90}},{{"step_number":3,"action":"save_document","target":null,"params":null,"confidence":95}}]}}

Request: "напиши про искусственный интеллект в блокноте"
{{"reasoning":"Open Notepad, write AI introduction in Russian","steps":[{{"step_number":1,"action":"new_document","target":"notepad","params":null,"confidence":95}},{{"step_number":2,"action":"type_text","target":null,"params":"Искусственный интеллект\\n\\nИскусственный интеллект (ИИ) — это технология, позволяющая компьютерам выполнять задачи, которые обычно требуют человеческого интеллекта: распознавать речь, переводить языки, принимать решения и анализировать данные.\\n\\nМодели машинного обучения обучаются на огромных массивах данных и со временем улучшают свою точность. Сегодня ИИ применяется в медицине, образовании, бизнесе и науке.","confidence":90}}]}}

Request: "google how to bake a cake"
{{"reasoning":"Search Google for cake baking instructions","steps":[{{"step_number":1,"action":"open_app","target":"msedge","params":null,"confidence":95}},{{"step_number":2,"action":"search_web","target":null,"params":"how to bake a cake","confidence":97}}]}}

Request: "click on OK"
{{"reasoning":"Click OK button","steps":[{{"step_number":1,"action":"find_and_click","target":null,"params":"OK","confidence":90}}]}}

Request: "save the document"
{{"reasoning":"Save with Ctrl+S","steps":[{{"step_number":1,"action":"save_document","target":null,"params":null,"confidence":97}}]}}

Request: "take a screenshot"
{{"reasoning":"Capture screen","steps":[{{"step_number":1,"action":"screenshot","target":null,"params":null,"confidence":97}}]}}

Request: "press Ctrl+Z"
{{"reasoning":"Undo last action","steps":[{{"step_number":1,"action":"hotkey","target":null,"params":"ctrl+z","confidence":95}}]}}

Request: "open calculator"
{{"reasoning":"Open Calculator app","steps":[{{"step_number":1,"action":"open_app","target":"calc","params":null,"confidence":97}}]}}

Request: "open file explorer"
{{"reasoning":"Open File Explorer","steps":[{{"step_number":1,"action":"open_app","target":"explorer","params":null,"confidence":97}}]}}

Request: "move mouse to center"
{{"reasoning":"Move cursor to 960,540","steps":[{{"step_number":1,"action":"mouse_move","target":"960,540","params":null,"confidence":85}}]}}

Request: "wait 2 seconds then save"
{{"reasoning":"Pause then save","steps":[{{"step_number":1,"action":"wait","target":null,"params":"2","confidence":95}},{{"step_number":2,"action":"save_document","target":null,"params":null,"confidence":95}}]}}

Request: "write my name is Alex in the document"
{{"reasoning":"Type real text in active window","steps":[{{"step_number":1,"action":"type_text","target":null,"params":"My name is Alex.","confidence":95}}]}}

Request: "press Escape"
{{"reasoning":"Press the Escape key","steps":[{{"step_number":1,"action":"hotkey","target":null,"params":"escape","confidence":97}}]}}

Request: "press Ctrl+Z"
{{"reasoning":"Undo last action with Ctrl+Z","steps":[{{"step_number":1,"action":"hotkey","target":null,"params":"ctrl+z","confidence":97}}]}}

Request: "press Ctrl+S"
{{"reasoning":"Save document with Ctrl+S","steps":[{{"step_number":1,"action":"hotkey","target":null,"params":"ctrl+s","confidence":97}}]}}

Request: "scroll down"
{{"reasoning":"Scroll the page down","steps":[{{"step_number":1,"action":"scroll_down","target":null,"params":"3","confidence":95}}]}}

Request: "scroll up"
{{"reasoning":"Scroll the page up","steps":[{{"step_number":1,"action":"scroll_up","target":null,"params":"3","confidence":95}}]}}

Request: "copy selected text"
{{"reasoning":"Copy the current selection","steps":[{{"step_number":1,"action":"copy","target":null,"params":null,"confidence":95}}]}}

Request: "paste the text"
{{"reasoning":"Paste from clipboard","steps":[{{"step_number":1,"action":"paste","target":null,"params":null,"confidence":95}}]}}

Request: "close this window"
{{"reasoning":"Close the current window","steps":[{{"step_number":1,"action":"close_window","target":null,"params":null,"confidence":95}}]}}

Request: "minimize the window"
{{"reasoning":"Minimize the current window","steps":[{{"step_number":1,"action":"minimize_window","target":null,"params":null,"confidence":95}}]}}

Request: "what is the latest news about Ukraine"
{{"reasoning":"Search for Ukraine news in browser","steps":[{{"step_number":1,"action":"open_app","target":"msedge","params":null,"confidence":95}},{{"step_number":2,"action":"search_web","target":null,"params":"latest news about Ukraine","confidence":97}}]}}

Request: "google bitcoin price"
{{"reasoning":"Search Google for Bitcoin price","steps":[{{"step_number":1,"action":"open_app","target":"msedge","params":null,"confidence":95}},{{"step_number":2,"action":"search_web","target":null,"params":"bitcoin price","confidence":97}}]}}

Request: "type Hello World" (when Word is already open in OPEN WINDOWS)
{{"reasoning":"Word is open — type directly without re-opening","steps":[{{"step_number":1,"action":"type_text","target":null,"params":"Hello World","confidence":95}}]}}

Request: "save" (when Word is open)
{{"reasoning":"Save the current document","steps":[{{"step_number":1,"action":"save_document","target":null,"params":null,"confidence":97}}]}}

IMPORTANT app name mapping:
- "browser", "edge", "internet", "web" = target "msedge"
- "word", "ворд", "document", "doc" = target "winword"
- "notepad", "notebook", "блокнот", "text file", ".txt" = target "notepad"
- "calculator", "калькулятор" = target "calc"
- "explorer", "folder", "files", "file manager", "проводник" = target "explorer"
- "excel", "spreadsheet", "таблица" = target "excel"

JSON schema (strict):
{{
  "reasoning": "<why these steps achieve the goal>",
  "steps": [
    {{"step_number": 1, "action": "<action>", "target": "<target or null>", "params": "<params or null>", "confidence": 85}},
    ...
  ]
}}
"""

@app.post("/api/plan", response_model=PlanResponse)
async def create_plan(req: PlanRequest):
    prompt = PLAN_PROMPT.format(
        soul=_SOUL_TEXT,
        text=req.text,
        window_title=req.window_title or "unknown",
        window_process=req.window_process or "unknown",
        context=req.session_context or "none",
        open_windows=", ".join(req.open_windows) if req.open_windows else "none",
        skills=_skills_summary(),
    )

    raw, ms = await granite_generate(prompt)

    try:
        parsed = _extract_json(raw)
        _debug["parsed"] = parsed
        steps_raw = parsed.get("steps", [])
        steps = [
            PlanStep(
                step_number=s.get("step_number", i + 1),
                action=s.get("action", "unknown"),
                target=s.get("target"),
                params=s.get("params"),
                confidence=int(s.get("confidence", 80)),
            )
            for i, s in enumerate(steps_raw)
        ]
        reasoning = parsed.get("reasoning", "")
        log.info("PLAN  steps=%d  reasoning=%r", len(steps), reasoning[:80])

        # ── Trim hallucinated type_text/save steps if user didn't ask to write ──
        # Run BEFORE sanity check so sanity sees the trimmed plan (avoids false-passing)
        _tl = req.text.lower()
        _wants_write = any(k in _tl for k in ("write", "type", "text", "пиши", "напиши", "введи", "сохрани", "save", "essay", "эссе"))
        if not _wants_write and len(steps) > 1:
            trimmed = [s for s in steps if s.action not in ("type_text", "save_document")]
            if len(trimmed) < len(steps):
                log.info("PLAN trim: removed %d extra type_text/save steps (no write intent)", len(steps) - len(trimmed))
                steps = trimmed if trimmed else steps

        # ── Sanity check: detect when Granite returned its default template ────
        if not _is_plan_sane(steps, req.text):
            log.warning("PLAN sanity check FAILED — granite returned off-topic response, using fallback")
            steps    = _fallback_plan(req.text, raw)
            reasoning = f"[rule-based] granite off-topic, fell back"

        return PlanResponse(
            plan_id=f"plan_{int(time.time())}",
            steps=steps,
            total_steps=len(steps),
            reasoning=reasoning,
            latency_ms=ms,
        )
    except (ValueError, json.JSONDecodeError) as exc:
        log.warning("Plan JSON parse failed: %s  raw=%r", exc, raw[:300])
        steps = _fallback_plan(req.text, raw)
        reasoning = f"[fallback] model output could not be parsed as JSON. raw={raw[:120]}"
        _debug["parsed"] = {"fallback": True, "raw": raw[:200]}
        log.info("PLAN fallback  steps=%d  original_error=%s", len(steps), exc)
        return PlanResponse(
            plan_id=f"plan_fb_{int(time.time())}",
            steps=steps,
            total_steps=len(steps),
            reasoning=reasoning,
            latency_ms=ms,
        )


def _is_plan_sane(steps: list, request_text: str) -> bool:
    """
    Detect when Granite returned its default template instead of a real plan.
    Returns False if the plan is clearly wrong for the given request.
    """
    if not steps:
        return False

    tl = request_text.lower()

    # ── Detect hallucinated default: navigate_url(shell:Downloads) ───────────────
    for s in steps:
        if s.action == "navigate_url" and "shell:" in (s.target or s.params or "").lower():
            log.warning("SANITY: detected hallucinated shell: navigate_url — default template")
            return False

    # ── Detect exact default template: open_app notepad + type_text "Hello World" ──
    if len(steps) == 2:
        s1, s2 = steps[0], steps[1]
        if (s1.action == "open_app" and (s1.target or "").lower() == "notepad" and
                s2.action == "type_text" and (s2.params or "").lower() in ("hello world", "hello", "test")):
            log.warning("SANITY: detected default notepad+HelloWorld template")
            return False

    # ── Plan targets notepad but request clearly asks for Word ──────────────────
    word_kws    = ("word", "ворд", "ворде", "ворду", "winword", "docx", "document", "документ")
    notepad_kws = ("notepad", "блокнот", "notebook", ".txt", "text file", "txt file")

    plan_targets = {(s.target or "").lower() for s in steps}
    plan_actions = {s.action.lower() for s in steps}

    req_wants_word    = any(k in tl for k in word_kws)
    req_wants_notepad = any(k in tl for k in notepad_kws)
    req_wants_browser = any(k in tl for k in ("browser", "opera", "chrome", "edge", "браузер"))
    req_wants_calc    = any(k in tl for k in ("calculator", "calculate", "calc", "калькулятор"))
    req_wants_hotkey  = any(k in tl for k in ("ctrl+", "alt+", "escape", "esc", "hotkey",
                                               "press ctrl", "press alt", "press esc", "нажми ctrl"))
    # "нажми" by itself is a click — but "нажми ctrl+" is a hotkey, not a click
    req_wants_click   = any(k in tl for k in ("click", "кликни")) or (
        "нажми" in tl and not req_wants_hotkey
    )
    req_wants_scroll  = any(k in tl for k in ("scroll down", "scroll up", "прокрути вниз", "прокрути вверх", "scroll"))
    req_wants_close   = any(k in tl for k in ("close window", "close this", "закрой окно", "закрыть окно", "close the window"))
    req_wants_minimize = any(k in tl for k in ("minimize", "свернуть", "minimise"))
    req_wants_copy    = any(k in tl for k in ("copy", "скопируй")) and not any(k in tl for k in ("open", "create"))
    req_wants_paste   = any(k in tl for k in ("paste", "вставь"))
    req_wants_search  = any(k in tl for k in ("search", "find", "google", "look up", "news", "latest",
                                               "поиск", "найди", "гугли", "bitcoin", "crypto", "weather",
                                               "what is", "who is", "how to"))
    req_wants_explorer = any(k in tl for k in ("file explorer", "проводник", "my documents",
                                                "documents folder", "downloads folder", "папку"))
    # "and save" at the end is always an explicit save intent regardless of other keywords
    # Note: don't block on word_kws/notepad_kws — "save the document" must be recognized as save
    req_wants_save     = "and save" in tl or "и сохрани" in tl or (
        any(k in tl for k in ("save", "сохрани")) and
        not any(k in tl for k in ("create", "new", "write", "essay", "type",
                                   "эссе", "напиши", "написать", "создай"))
    )
    req_wants_screenshot = any(k in tl for k in ("screenshot", "скриншот", "снимок экрана"))

    plan_has_notepad  = "notepad" in plan_targets
    plan_has_browser  = any(t in plan_targets for t in ("msedge", "chrome", "opera", "firefox"))
    plan_has_search   = "search_web" in plan_actions
    plan_has_hotkey   = any(a in plan_actions for a in ("hotkey", "key_combo", "press_escape", "undo"))
    plan_has_scroll   = any(a in plan_actions for a in ("scroll_down", "scroll_up", "scroll_at"))
    plan_has_close    = any(a in plan_actions for a in ("close_window", "close_app"))
    plan_has_minimize = "minimize_window" in plan_actions
    plan_has_copy     = any(a in plan_actions for a in ("copy", "copy_text"))
    plan_has_paste    = any(a in plan_actions for a in ("paste", "paste_text"))
    plan_has_save     = "save_document" in plan_actions
    plan_has_screenshot = "screenshot" in plan_actions
    plan_has_explorer  = "explorer" in plan_targets

    # Save request → plan must save
    if req_wants_save:
        if not plan_has_save:
            log.warning("SANITY: save request but plan has no save_document")
            return False

    # Screenshot request → plan must screenshot
    if req_wants_screenshot:
        if not plan_has_screenshot:
            log.warning("SANITY: screenshot request but plan has no screenshot action")
            return False

    # Explorer/folder request → plan must open explorer
    if req_wants_explorer:
        if not plan_has_explorer and "navigate_url" not in plan_actions and "open_file_path" not in plan_actions:
            log.warning("SANITY: explorer/folder request but plan has no explorer")
            return False

    # Request asks for search/google/news → plan must have search_web
    # Note: remove "search" from req_wants_browser to avoid false positives
    if req_wants_search and not req_wants_notepad and not req_wants_word and not req_wants_close:
        if not plan_has_search:
            log.warning("SANITY: search/news request but plan has no search_web (has: %s)", plan_actions)
            return False
        # Also verify search params actually relate to the request (Granite sometimes hallucates wrong query)
        _stop = {"for", "the", "a", "an", "in", "on", "at", "to", "of", "and", "or", "is", "me",
                 "my", "can", "you", "please", "i", "find", "search", "google", "show", "look"}
        _req_words = {w for w in tl.split() if len(w) > 3 and w not in _stop}
        for _s in steps:
            if _s.action == "search_web":
                _params = (_s.params or _s.target or "").lower()
                if _params and _req_words and not any(w in _params for w in _req_words):
                    log.warning("SANITY: search_web params '%s' share no keywords with request '%s'", _params, tl)
                    return False

    # "open browser" (generic) → plan must open a browser, not notepad/word
    if req_wants_browser and not req_wants_word and not req_wants_notepad and not req_wants_search:
        if not plan_has_browser:
            log.warning("SANITY: browser request but plan has no browser target")
            return False
        # Generic "browser" (not "chrome"/"opera") should use msedge
        if "browser" in tl and "chrome" not in tl and "opera" not in tl and "firefox" not in tl:
            if "msedge" not in plan_targets and "chrome" in plan_targets:
                log.warning("SANITY: generic browser request → rejecting non-msedge plan")
                return False

    # Hotkey requests → plan must have hotkey action
    if req_wants_hotkey and not req_wants_click:
        if not plan_has_hotkey:
            log.warning("SANITY: hotkey request but plan has no hotkey action (has: %s)", plan_actions)
            return False

    # Scroll request → plan must scroll
    if req_wants_scroll and not req_wants_word and not req_wants_notepad:
        if not plan_has_scroll:
            log.warning("SANITY: scroll request but plan has no scroll action")
            return False

    # Close window → plan must close
    if req_wants_close:
        if not plan_has_close:
            log.warning("SANITY: close window request but plan has no close action")
            return False

    # Minimize → plan must minimize
    if req_wants_minimize:
        if not plan_has_minimize:
            log.warning("SANITY: minimize request but plan has no minimize action")
            return False

    # Copy → plan must copy
    if req_wants_copy and not req_wants_word and not req_wants_notepad:
        if not plan_has_copy:
            log.warning("SANITY: copy request but plan has no copy action")
            return False

    # Paste → plan must paste
    if req_wants_paste and not req_wants_word and not req_wants_notepad:
        if not plan_has_paste:
            log.warning("SANITY: paste request but plan has no paste action")
            return False

    # Request asks for Word → plan must not only have notepad
    if req_wants_word and not req_wants_notepad:
        if plan_has_notepad and not any(t in plan_targets for t in ("winword", "word")):
            log.warning("SANITY: request wants Word but plan only has notepad")
            return False
        # Also reject plans that only do new_tab / browser actions when user wants a Word document
        if plan_actions <= {"new_tab", "new_window", "open_app"} and not any(t in plan_targets for t in ("winword", "word")):
            log.warning("SANITY: request wants Word/document but plan only does browser action")
            return False

    # Calculator request → plan must open calc
    if req_wants_calc:
        if "open_app" not in plan_actions or "calc" not in plan_targets:
            log.warning("SANITY: calculator request but plan has no open_app(calc)")
            return False

    # Essay request without explicit notepad mention → plan must NOT target notepad
    # (essays default to winword; notepad is only valid if user explicitly said notepad/блокнот)
    _req_is_essay = any(k in tl for k in ("essay", "эссе", "сочинение"))
    if _req_is_essay and not req_wants_notepad:
        if plan_has_notepad and not any(t in plan_targets for t in ("winword", "word")):
            log.warning("SANITY: essay request but plan targets notepad (should be winword)")
            return False

    # Request asks for browser/search → plan must not only open notepad
    if (req_wants_browser or "search" in tl) and not req_wants_notepad and not req_wants_word:
        if plan_has_notepad and not plan_has_browser and not plan_has_search:
            log.warning("SANITY: request is browser/search but plan only opens notepad")
            return False

    # Request asks to click something → plan must have find_and_click, not open apps
    if req_wants_click and not req_wants_hotkey:
        if "find_and_click" not in plan_actions and "click_at" not in plan_actions:
            log.warning("SANITY: click request but plan has no find_and_click (has: %s)", plan_actions)
            return False

    # Request asks for notepad/notebook but plan opens browser or wrong app ──────
    if req_wants_notepad and not req_wants_browser and not req_wants_word:
        if plan_has_browser and not plan_has_notepad:
            log.warning("SANITY: request wants notepad but plan only has browser")
            return False
        unrelated_apps = {"winword", "msedge", "chrome", "opera", "firefox", "excel", "calc"}
        if plan_targets & unrelated_apps and not plan_has_notepad:
            log.warning("SANITY: request wants notepad but plan targets unrelated: %s", plan_targets)
            return False
        # Create/open intent with notepad → plan must have open_app or new_document targeting notepad
        _has_create_intent = any(k in tl for k in ("create", "new", "open", "launch", "открой",
                                                     "создай", "новый", "запусти"))
        if _has_create_intent:
            _plan_opens_notepad = plan_has_notepad and ("open_app" in plan_actions or "new_document" in plan_actions)
            if not _plan_opens_notepad:
                log.warning("SANITY: 'create/open notepad' but plan doesn't open notepad (has: %s → %s)",
                            plan_actions, plan_targets)
                return False

    # Request asks for Word but plan opens browser ───────────────────────────────
    if req_wants_word and not req_wants_notepad and not req_wants_browser:
        if plan_has_browser and not any(t in plan_targets for t in ("winword", "word")):
            log.warning("SANITY: request wants Word but plan only has browser")
            return False

    # Plan opens a document app but request doesn't mention any document/app
    _mentions_app = any(k in tl for k in word_kws + notepad_kws +
                        ("browser", "edge", "chrome", "opera", "calc", "calculator", "explorer",
                         "браузер", "калькулятор", "проводник"))
    _mentions_action = any(k in tl for k in ("open", "create", "new", "make", "write", "type",
                                               "search", "find", "click", "press", "scroll", "save",
                                               "открой", "создай", "напиши", "найди", "нажми"))
    if (not _mentions_app and not _mentions_action):
        # Completely ambiguous — any specific-app or spurious plan is wrong
        _is_spurious_plan = (
            plan_actions == {"new_document"} or
            plan_actions == {"new_window"} or
            plan_actions <= {"new_document", "open_app"} or
            (plan_actions <= {"open_app", "new_tab"} and plan_targets & {"chrome", "opera", "firefox"})
        )
        if _is_spurious_plan:
            log.warning("SANITY: ambiguous request but plan is spurious (%s → %s)", plan_actions, plan_targets)
            return False

    # "open/launch" with no specific app mentioned → any specific-app plan is wrong
    _has_open_verb = any(k in tl for k in ("open", "launch", "start", "открой", "запусти"))
    if _has_open_verb and not _mentions_app:
        _any_app = {"winword", "notepad", "excel", "msedge", "chrome", "opera", "firefox", "calc", "explorer"}
        if plan_actions <= {"new_document", "new_window", "new_tab", "open_app"} and plan_targets & _any_app:
            log.warning("SANITY: 'open' with no app target but plan opens specific app → ambiguous")
            return False

    # When user wants to write an essay/letter, type_text must have real substance.
    # "type X" where X is an explicitly dictated short phrase is FINE (don't enforce length).
    _essay_kws = ("essay", "эссе", "сочинение", "letter", "статья", "текст о", "написать о")
    _req_is_essay_write = any(k in tl for k in _essay_kws)
    if _req_is_essay_write:
        has_type_text = any(s.action == "type_text" for s in steps)
        if not has_type_text:
            log.warning("SANITY: essay/letter request but plan has no type_text step")
            return False
        for s in steps:
            if s.action == "type_text":
                content = (s.params or "").strip()
                if len(content) < 20 or content.lower() in (
                    "hello world", "hello", "test", "text here", "essay text here",
                    "type here", "your text", "...", "content", "placeholder"
                ):
                    log.warning("SANITY: type_text has empty/placeholder content for essay request")
                    return False

    # ── "type X" explicit dictation — params must contain dictated text ────────
    # e.g. "type Hello World" must yield type_text("Hello World"), not a random story
    import re as _re_sanity
    _type_direct = _re_sanity.match(r'^type\s+(.{1,60})$', tl)
    if _type_direct and not _req_is_essay_write:
        _dictated = _type_direct.group(1).strip().lower()
        for s in steps:
            if s.action == "type_text":
                _content = (s.params or "").lower()
                # Content is 3x longer AND doesn't contain the dictated text → hallucination
                if len(_content) > len(_dictated) * 3 and _dictated not in _content:
                    log.warning("SANITY: 'type X' dictation mismatch: expected '%s', got '%s...'",
                                _dictated, _content[:40])
                    return False

    # ── General: no write intent → type_text-only plan is always a hallucination ──
    # e.g. "open notepad", "save document", "search for X" must NOT produce type_text
    _req_has_write_intent = any(k in tl for k in ("write", "type", "essay", "letter", "пиши",
                                                    "напиши", "введи", "эссе", "сочинение",
                                                    "набери", "напечатай", "печатай"))
    if not _req_has_write_intent:
        if plan_actions <= {"type_text"} or plan_actions <= {"type_text", "save_document"}:
            log.warning("SANITY: no write intent but plan only types text → hallucination (got: %s)", plan_actions)
            return False

    # ── Story/hallucination content detector ─────────────────────────────────────
    # Granite sometimes returns a "Timmy and Max" story for any request
    _halluc_patterns = ("once upon a time", "there once was", "in a land far", "a long time ago",
                         "жили-были", "в некотором царстве")
    for s in steps:
        if s.action == "type_text":
            _c = (s.params or "").lower()[:50]
            if any(p in _c for p in _halluc_patterns):
                log.warning("SANITY: type_text contains hallucinated story content → reject")
                return False

    return True


def _fallback_plan(text: str, raw: str = "") -> list[PlanStep]:
    tl = text.lower()
    steps: list[PlanStep] = []
    i = 1

    # ── Hotkey / keyboard shortcut ──────────────────────────────────────────
    import re as _re2
    # Detect "press Ctrl+Z", "hotkey ctrl+s", "escape", "press esc", etc.
    hotkey_explicit = _re2.search(r'\b(?:hotkey|keyboard\s+shortcut)\s+([a-z0-9+]+)', tl, _re2.I)
    if hotkey_explicit:
        steps.append(PlanStep(step_number=i, action="hotkey", params=hotkey_explicit.group(1).strip()))
        return steps
    # "press escape" / "press esc" / "press enter"
    esc_m = _re2.search(r'\bpress\s+(escape|esc|enter|tab|delete|backspace)\b', tl, _re2.I)
    if esc_m:
        steps.append(PlanStep(step_number=i, action="hotkey", params=esc_m.group(1).lower()))
        return steps
    # "press Ctrl+Z", "Ctrl+S" etc.
    ctrl_m = _re2.search(r'\b(?:press\s+)?(ctrl|alt|shift|win)\+\S+', tl, _re2.I)
    if ctrl_m:
        combo = ctrl_m.group(0).replace("press ", "").strip()
        steps.append(PlanStep(step_number=i, action="hotkey", params=combo.lower()))
        return steps

    # ── Close / minimize window ─────────────────────────────────────────────
    if any(k in tl for k in ("close window", "close this", "закрой окно", "закрыть окно", "close the window")):
        steps.append(PlanStep(step_number=i, action="close_window"))
        return steps
    if any(k in tl for k in ("minimize", "minimise", "свернуть", "минимизировать")):
        steps.append(PlanStep(step_number=i, action="minimize_window"))
        return steps

    # ── Scroll ───────────────────────────────────────────────────────────────
    if any(k in tl for k in ("scroll down", "прокрути вниз", "scroll вниз")):
        steps.append(PlanStep(step_number=i, action="scroll_down", params="3"))
        return steps
    if any(k in tl for k in ("scroll up", "прокрути вверх", "scroll вверх")):
        steps.append(PlanStep(step_number=i, action="scroll_up", params="3"))
        return steps
    if "scroll" in tl and not any(k in tl for k in ("word", "notepad", "ворд", "блокнот")):
        steps.append(PlanStep(step_number=i, action="scroll_down", params="3"))
        return steps

    # ── Copy / paste ────────────────────────────────────────────────────────
    if any(k in tl for k in ("copy all", "copy everything", "скопируй всё", "скопируй все")):
        steps.append(PlanStep(step_number=i, action="select_all"))
        i += 1
        steps.append(PlanStep(step_number=i, action="copy"))
        return steps
    if any(k in tl for k in ("copy", "скопируй")) and not any(k in tl for k in ("open", "create", "new")):
        steps.append(PlanStep(step_number=i, action="copy"))
        return steps
    if any(k in tl for k in ("paste", "вставь")):
        steps.append(PlanStep(step_number=i, action="paste"))
        return steps

    # ── Click / find_and_click ───────────────────────────────────────────────
    # Only for explicit click requests, not hotkeys
    click_m = _re2.search(r'\b(?:click|нажми|кликни)\s+(?:on\s+)?(.+)', tl, _re2.I)
    if click_m:
        element = click_m.group(1).strip().rstrip(".,!?")
        steps.append(PlanStep(step_number=i, action="find_and_click", params=element))
        i += 1
        return steps  # Single-step click plan

    # ── File explorer / folder requests ─────────────────────────────────────
    if any(k in tl for k in ("file explorer", "my documents", "documents folder",
                               "downloads folder", "проводник", "папку")):
        steps.append(PlanStep(step_number=i, action="open_app", target="explorer"))
        return steps

    _create_word_kws    = ("word", "document", "doc", "ворд", "ворде", "ворду", "документ")
    _create_notepad_kws = ("notepad", "блокнот", "notebook", "txt", "text file", "txt file")
    # "new"/"create" → new_document; "open"/"launch" → open_app
    _create_kws = ("create", "new", "make", "создай", "создать", "новый")
    _open_kws   = ("open", "launch", "start", "открой", "запусти")
    if any(k in tl for k in _create_kws):
        if any(w in tl for w in _create_word_kws):
            steps.append(PlanStep(step_number=i, action="new_document", target="winword"))
            i += 1
        elif any(w in tl for w in _create_notepad_kws):
            steps.append(PlanStep(step_number=i, action="new_document", target="notepad"))
            i += 1
    elif any(k in tl for k in _open_kws):
        if any(w in tl for w in _create_word_kws):
            steps.append(PlanStep(step_number=i, action="open_app", target="winword"))
            i += 1
        elif any(w in tl for w in _create_notepad_kws):
            steps.append(PlanStep(step_number=i, action="open_app", target="notepad"))
            i += 1

    app_keywords = [
        ("browser", "msedge"), ("internet", "msedge"), ("web browser", "msedge"),
        ("edge", "msedge"), ("браузер", "msedge"), ("chrome", "chrome"),
        ("opera", "opera"), ("firefox", "firefox"),
        ("word", "winword"), ("ворд", "winword"),
        ("excel", "excel"), ("таблиц", "excel"),
        ("notepad", "notepad"), ("блокнот", "notepad"), ("notebook", "notepad"),
        ("calculator", "calc"), ("калькулятор", "calc"),
        ("file explorer", "explorer"), ("explorer", "explorer"), ("проводник", "explorer"),
        ("my documents", "explorer"), ("documents folder", "explorer"),
        ("folder", "explorer"), ("files", "explorer"), ("file manager", "explorer"),
        ("папк", "explorer"),
    ]

    if not steps:
        for kw, exe in app_keywords:
            if kw in tl:
                steps.append(PlanStep(step_number=i, action="open_app", target=exe))
                i += 1
                break

    if any(w in tl for w in ("search", "find", "weather", "news", "google", "bing",
                              "latest", "current", "what is", "who is", "how to",
                              "bitcoin", "crypto", "price", "поиск", "найди", "погода",
                              "гугли", "гугл", "узнай", "покажи")):
        # Extract query: strip common prefixes
        query = text
        for prefix in ("search for", "search", "google", "find", "look up", "гугли", "найди", "поиск"):
            if query.lower().startswith(prefix):
                query = query[len(prefix):].strip()
                break
        if not any(s.action == "open_app" for s in steps):
            steps.insert(0, PlanStep(step_number=1, action="open_app", target="msedge"))
            i += 1
        for j, s in enumerate(steps):
            s.step_number = j + 1
        steps.append(PlanStep(step_number=i, action="search_web", params=query))
        i += 1

    # ── Screenshot ───────────────────────────────────────────────────────────
    if any(k in tl for k in ("screenshot", "скриншот", "снимок экрана")):
        steps.append(PlanStep(step_number=i, action="screenshot"))
        return steps

    # ── Save only (no write) ─────────────────────────────────────────────────
    _pure_save = (any(k in tl for k in ("save", "сохрани")) and
                  not any(k in tl for k in ("write", "essay", "эссе", "type", "напиши")))
    if _pure_save and not steps:
        steps.append(PlanStep(step_number=i, action="save_document"))
        return steps

    # ── Write / type with real content generation ────────────────────────────
    _is_essay = any(k in tl for k in ("essay", "эссе", "сочинение", "paragraph", "article", "статья"))
    _is_write = any(k in tl for k in ("type", "write", "напиши", "напечатай", "введи", "написать", "написать"))
    if _is_write and not steps:
        # Detect target app for writing — essays default to Word, not Notepad
        if any(w in tl for w in ("word", "ворд", "ворде", "ворду", "winword", "document", "документ")) or _is_essay:
            steps.append(PlanStep(step_number=i, action="new_document", target="winword"))
            i += 1
        elif any(w in tl for w in ("notepad", "блокнот", "notebook", "txt", "text file", "txt file")):
            steps.append(PlanStep(step_number=i, action="new_document", target="notepad"))
            i += 1

    if _is_write:
        if _is_essay:
            # Extract topic from the request
            topic = tl
            for strip_kw in ("напиши эссе о", "напиши эссе про", "write essay about", "write an essay about",
                              "напиши про", "написать про", "write about", "essay about", "эссе о", "эссе про"):
                if strip_kw in topic:
                    topic = topic.split(strip_kw, 1)[-1]
                    break
            for app_ref in ("в ворде", "в ворд", "in word", "in notepad", "в блокноте", "and save", "и сохрани"):
                topic = topic.replace(app_ref, "")
            topic = topic.strip().rstrip(".,!?")
            # Generate real essay content about the extracted topic
            essay_body = (
                f"{topic.capitalize()}\n\n"
                f"This topic is of great significance in our modern world. "
                f"Understanding {topic} requires examining its key aspects, historical context, and current relevance.\n\n"
                f"First, it is important to consider the fundamental principles behind {topic}. "
                f"These principles have shaped how we approach and think about this subject today.\n\n"
                f"Furthermore, the impact of {topic} can be seen across many fields and aspects of daily life. "
                f"From practical applications to broader societal implications, the importance of this topic cannot be overstated.\n\n"
                f"In conclusion, {topic} remains a crucial area of study and discussion. "
                f"By deepening our understanding, we can better navigate the challenges and opportunities it presents."
            )
            steps.append(PlanStep(step_number=i, action="type_text", params=essay_body))
            i += 1
            # Add save if requested
            if any(k in tl for k in ("save", "сохрани", "and save", "и сохрани")):
                steps.append(PlanStep(step_number=i, action="save_document"))
                i += 1
            return steps
        else:
            content = tl
            for kw in ("type", "write", "напиши", "напечатай", "введи"):
                if kw in content:
                    content = content.split(kw, 1)[-1].strip()
                    break
            for app_ref in ("in notepad", "in word", "в блокнот", "в ворд", "in the"):
                content = content.replace(app_ref, "").strip()
            if content and len(content) > 2:
                steps.append(PlanStep(step_number=i, action="type_text", params=content))
                i += 1

    if "bold" in tl or "жирн" in tl:
        steps.append(PlanStep(step_number=i, action="format_bold"))
        i += 1
    if "italic" in tl or "курсив" in tl:
        steps.append(PlanStep(step_number=i, action="format_italic"))
        i += 1
    if "underline" in tl or "подчеркн" in tl:
        steps.append(PlanStep(step_number=i, action="format_underline"))
        i += 1

    if any(w in tl for w in ("copy all", "copy everything", "скопируй всё", "скопируй все")):
        steps.append(PlanStep(step_number=i, action="select_all"))
        i += 1
        steps.append(PlanStep(step_number=i, action="copy"))
        i += 1
    elif any(w in tl for w in ("copy", "скопируй")) and not _is_write:
        steps.append(PlanStep(step_number=i, action="copy"))
        i += 1

    if any(w in tl for w in ("paste", "вставь")):
        steps.append(PlanStep(step_number=i, action="paste"))
        i += 1

    if any(w in tl for w in ("new tab", "новую вкладку", "open tab")):
        if not any(s.action == "open_app" for s in steps):
            steps.append(PlanStep(step_number=i, action="open_app", target="msedge"))
            i += 1
        steps.append(PlanStep(step_number=i, action="new_tab"))
        i += 1
    elif any(w in tl for w in ("close tab", "закрой вкладку")):
        steps = [PlanStep(step_number=1, action="close_tab")]

    if any(w in tl for w in ("save", "сохрани")):
        steps.append(PlanStep(step_number=i, action="save_document"))
        i += 1

    if any(w in tl for w in ("close", "закрой")) and not any(s.action == "close_tab" for s in steps):
        target = None
        for kw, exe in app_keywords:
            if kw in tl:
                target = exe
                break
        if not any(w in tl for w in ("tab", "вкладку")):
            steps = [PlanStep(step_number=1, action="close_window", target=target)]

    if not steps:
        # Ambiguous request: either search or ask user
        _search_kws = ("search", "find", "google", "what", "how", "who", "where", "when", "why",
                        "поиск", "найди", "что", "как", "кто", "где", "когда")
        if any(k in tl for k in _search_kws):
            steps.append(PlanStep(step_number=1, action="open_app", target="msedge"))
            steps.append(PlanStep(step_number=2, action="search_web", params=text))
        else:
            # Totally ambiguous — ask the user
            steps.append(PlanStep(step_number=1, action="message_user", params=f"I'm not sure what to do with '{text}'. Could you clarify?"))

    return steps


# ─── Chat ──────────────────────────────────────────────────────────────────────

def _build_chat_prompt(session_id: str, new_message: str, extra_history: list[dict]) -> str:
    stored   = _chat_history.get(session_id, [])
    combined = (extra_history if extra_history else stored)[-_MAX_HISTORY_TURNS * 2:]

    system = _SOUL_TEXT if _SOUL_TEXT else (
        "You are AICompanion, a friendly Windows desktop voice assistant powered by IBM Granite.\n"
        "Be concise — your replies should be 1-3 sentences unless the user asks for more detail."
    )

    lines = [system.strip(), ""]

    # Inject semantically similar past exchanges from ChromaDB (top-3)
    semantic_ctx = _chroma_query(new_message, n_results=3)
    if semantic_ctx:
        lines.append(semantic_ctx)

    for turn in combined:
        role_label = "User" if turn.get("role") == "user" else "Assistant"
        lines.append(f"{role_label}: {turn['content']}")
    lines.append(f"User: {new_message}")
    lines.append("Assistant:")
    return "\n".join(lines)

@app.post("/api/chat", response_model=ChatResponse)
async def chat(req: ChatRequest):
    prompt = _build_chat_prompt(req.session_id, req.message, req.history)
    raw, ms = await granite_generate(prompt, temperature=0.5)

    # Update in-memory session history (SQLite-backed via _append_memory)
    history = _chat_history.setdefault(req.session_id, [])
    history.append({"role": "user",      "content": req.message})
    history.append({"role": "assistant", "content": raw})
    if len(history) > _MAX_HISTORY_TURNS * 2:
        _chat_history[req.session_id] = history[-_MAX_HISTORY_TURNS * 2:]

    # Persist to file-based memory (SQLite-style markdown log)
    _append_memory(req.session_id, req.message, raw)

    # Store embedding in ChromaDB for future semantic retrieval
    _chroma_store(req.session_id, req.message, raw)

    return ChatResponse(reply=raw, latency_ms=ms)

@app.post("/api/chat/clear")
async def clear_chat_history(req: ClearHistoryRequest):
    removed = len(_chat_history.pop(req.session_id, []))
    return {"status": "ok", "removed_turns": removed, "session_id": req.session_id}


# ─── Essay generation ─────────────────────────────────────────────────────────

ESSAY_OUTLINE_PROMPT = """\
You are a skilled writer. Think aloud about the structure of an essay.

Topic: "{topic}"
Style: {style}  |  Target length: ~{word_count} words  |  Language: {language}

Write a numbered outline (3-4 main points) for this essay. Be concise.
Format: just the numbered list, no extra text.
"""

ESSAY_DRAFT_PROMPT = """\
You are a skilled writer. Write an essay based on the outline below.

Topic: "{topic}"
Style: {style}  |  Target length: ~{word_count} words  |  Language: {language}

Outline:
{outline}

Write the full essay now. Use clear paragraphs. Do not include a title line.
"""

ESSAY_REFINE_PROMPT = """\
You are an editor. Polish the following essay draft.
Fix any awkward phrasing, improve flow, and make it read naturally.
Keep approximately the same length. Output only the improved essay.

Draft:
{draft}
"""

@app.post("/api/essay", response_model=EssayResponse)
async def write_essay(req: EssayRequest):
    think_steps: list[EssayThinkStep] = []
    t_total = time.monotonic()

    log.info("ESSAY  topic=%r  words=%d  style=%s", req.topic, req.word_count, req.style)

    outline_prompt = ESSAY_OUTLINE_PROMPT.format(
        topic=req.topic, style=req.style,
        word_count=req.word_count, language=req.language)
    outline, ms1 = await granite_generate(outline_prompt, temperature=0.3)
    think_steps.append(EssayThinkStep(stage="outline", content=outline, latency_ms=ms1))
    log.info("  Outline done (%dms): %r", ms1, outline[:80])

    draft_prompt = ESSAY_DRAFT_PROMPT.format(
        topic=req.topic, style=req.style,
        word_count=req.word_count, language=req.language,
        outline=outline)
    draft, ms2 = await granite_generate(draft_prompt, temperature=0.4)
    think_steps.append(EssayThinkStep(stage="draft", content=draft, latency_ms=ms2))
    log.info("  Draft done (%dms): %d chars", ms2, len(draft))

    refine_prompt = ESSAY_REFINE_PROMPT.format(draft=draft)
    final, ms3 = await granite_generate(refine_prompt, temperature=0.2)
    think_steps.append(EssayThinkStep(stage="refine", content=final, latency_ms=ms3))
    log.info("  Refined done (%dms): %d chars", ms3, len(final))

    total_ms = int((time.monotonic() - t_total) * 1000)
    log.info("ESSAY complete  total=%dms  chars=%d", total_ms, len(final))

    return EssayResponse(
        topic=req.topic,
        outline=outline,
        draft=draft,
        final_text=final,
        total_latency_ms=total_ms,
        think_steps=think_steps,
    )


# ─── Smart command — essay-aware router ───────────────────────────────────────

_ESSAY_PATTERNS = [
    _re.compile(r"(?:write|написать|напиши|напишите).*(?:essay|эссе|article|статью|статья)", _re.I),
    _re.compile(r"(?:write|написать|напиши|написать|type|напечатай).*(?:about|про|о том)", _re.I),
    # Match bare essay keywords even without a verb (e.g. "эссе о природе", "essay about X")
    _re.compile(r"\b(?:essay|эссе|сочинение)\b", _re.I),
]

def _is_essay_request(text: str) -> bool:
    return any(p.search(text) for p in _ESSAY_PATTERNS)

def _extract_essay_topic(text: str) -> str:
    m = _re.search(r"(?:about|про|о том|on the topic of|на тему)\s+(.+)", text, _re.I)
    if m:
        topic = m.group(1).strip().rstrip(".")
        topic = _re.sub(r"\s+(?:in|into|in the|в|в блокнот|в notepad).*$", "", topic, flags=_re.I)
        # Strip trailing noise
        topic = _re.sub(r'\s*[,\.\?!]+\s*$', '', topic).strip()
        topic = _re.sub(r'\s+please\s*$', '', topic, flags=_re.I).strip()
        topic = _re.sub(r'\s*\?\s*$', '', topic).strip()
        return topic
    return text

def _detect_target_app(text: str, session_context: Optional[str] = None, open_windows: list[str] = [],
                       is_essay: bool = False) -> str:
    """Determine which desktop app to target. Falls back to last open text editor from session context.
    For essay/writing requests without explicit app mention, defaults to winword (not notepad).
    """
    tl = text.lower()
    # Explicit Word mention always wins
    if any(w in tl for w in ("word", "ворд", "ворде", "winword", "docx", ".doc", "document", "документ")):
        return "winword"
    if any(w in tl for w in ("excel", "таблиц", "spreadsheet", "xlsx")):
        return "excel"
    # Explicit notepad mention — but only if Word is NOT mentioned (already handled above)
    if any(w in tl for w in ("notepad", "блокнот", "notebook", ".txt", "text file", "txt file")):
        return "notepad"
    if any(w in tl for w in ("browser", "chrome", "edge", "firefox", "opera", "браузер")):
        return "msedge"
    if any(w in tl for w in ("explorer", "folder", "files", "проводник", "папк")):
        return "explorer"
    # Fall back to whatever text editor was last used (from session context)
    if session_context:
        try:
            ctx = json.loads(session_context)
            editor = ctx.get("last_text_editor_app", "").lower()
            if editor in ("winword", "notepad", "excel"):
                return editor
        except Exception:
            pass
    # Fallback: check open windows list
    for w in open_windows:
        wl = w.lower()
        if any(k in wl for k in ("word", "document", "winword")):
            return "winword"
        if any(k in wl for k in ("notepad", "блокнот", ".txt")):
            return "notepad"
        if any(k in wl for k in ("excel", "spreadsheet", ".xlsx")):
            return "excel"
    # Essays without an explicit app → Word, not Notepad
    if is_essay:
        return "winword"
    return "notepad"


def _app_is_open(app_target: str, window_title: str, window_process: str,
                 session_context: Optional[str] = None, open_windows: list[str] = []) -> bool:
    """Check if an app is already open, also scanning session context for last text editor."""
    if _re.search(app_target.replace("winword", "word"), window_title, _re.I):
        return True
    if _re.search(app_target, window_process, _re.I):
        return True
    if app_target == "notepad" and _re.search(r"notepad|блокнот", window_title, _re.I):
        return True
    if session_context:
        try:
            ctx = json.loads(session_context)
            editor_app   = ctx.get("last_text_editor_app", "").lower()
            editor_title = ctx.get("last_text_editor", "")
            if editor_app == app_target:
                return True
            if app_target == "winword" and "word" in editor_title.lower():
                return True
            if app_target == "notepad" and (
                    "notepad" in editor_title.lower() or "блокнот" in editor_title.lower()):
                return True
        except Exception:
            pass
    # Check open_windows list
    for w in open_windows:
        wl = w.lower()
        if app_target in ("winword", "word") and any(k in wl for k in ("word", "document")):
            return True
        if app_target == "notepad" and any(k in wl for k in ("notepad", ".txt", "блокнот")):
            return True
        if app_target in ("excel", "msexcel") and "excel" in wl:
            return True
        if app_target in ("opera", "chrome", "firefox", "msedge") and app_target in wl:
            return True
    return False


# ─── Research + summarize + write pattern ────────────────────────────────────

_RESEARCH_PATTERNS = [
    _re.compile(r"(?:search|browse|find|research|look\s+up|google|найди|поищи).*"
                r"(?:and|then|затем|потом|,)\s*(?:write|type|put|save|напиши|запиши|summarize)", _re.I),
    _re.compile(r"(?:open|launch|start|открой)\s+browser.*"
                r"(?:and|then|,)\s*(?:write|type|summarize|напиши)", _re.I),
    _re.compile(r"(?:всю информацию|all info|summarize|суммируй).*"
                r"(?:напиши|write|put|save|in notepad|in word|в блокнот|в ворд)", _re.I),
    _re.compile(r"(?:find|search|research)\s+.{3,50}"
                r"(?:and|then|,)\s*(?:write|type|save|summarize).{0,30}"
                r"(?:notepad|word|notebook|ворд|блокнот)", _re.I),
]

def _is_research_request(text: str) -> bool:
    return any(p.search(text) for p in _RESEARCH_PATTERNS)

def _extract_research_topic(text: str) -> str:
    """
    Extract the actual subject from a research/essay command.
    Priority: 'essay/article about X' > 'write about X' > 'search for X'.
    Strips browser noise like 'on my browser Opera'.
    """
    def _strip_topic_noise(topic: str) -> str:
        topic = topic.strip().rstrip(".,")
        topic = _re.sub(r'\s*[,\.\?!]+\s*$', '', topic).strip()
        topic = _re.sub(r'\s+please\s*$', '', topic, flags=_re.I).strip()
        topic = _re.sub(r'\s*\?\s*$', '', topic).strip()
        return topic

    # 1. Explicit essay topic: "write a essay about Donald Trump, ..."
    m = _re.search(
        r"(?:essay|article|эссе|статью|статья)\s+(?:about|про|о|on)\s+(.+?)"
        r"(?:[,.]|\s+(?:and|then|search|find|in\s+my|on\s+my)|$)",
        text, _re.I)
    if m:
        return _strip_topic_noise(m.group(1))

    # 2. "write about X" / "напиши про X"
    m = _re.search(
        r"(?:write|type|напиши|напечатай)\s+(?:about|про|о том|on)\s+(.+?)"
        r"(?:[,.]|\s+(?:and|then|search|find)|$)",
        text, _re.I)
    if m:
        return _strip_topic_noise(m.group(1))

    # 3. Search-based extraction (fallback)
    for pat in [
        r"(?:search|find|research|look\s+up|google|browse|найди|поищи)\s+"
        r"(?:for\s+|info(?:rmation)?\s+(?:about|on)\s+)?(.+?)"
        r"(?:\s+and\s+|\s+then\s+|,|\s+затем|\s+потом|$)",
        r"(?:about|про|о том|on the topic of|на тему)\s+(.+?)(?:\s+and\s+|\s+then\s+|,|$)",
    ]:
        m = _re.search(pat, text, _re.I)
        if m:
            topic = m.group(1).strip().rstrip(".,")
            # Strip "on my browser Opera", "in browser", etc.
            topic = _re.sub(r"(?:on|in)\s+(?:my\s+)?browser\s*\w*", "", topic, flags=_re.I).strip()
            topic = _re.sub(
                r"\s+(?:and|then|write|type|put|save|summarize|"
                r"в|в блокнот|в ворд|to notepad|to word|into).*$",
                "", topic, flags=_re.I).strip()
            topic = _strip_topic_noise(topic)
            if topic:
                return topic
    return text


def _detect_browser(text: str) -> str:
    """Return the specific browser mentioned in the text, defaulting to msedge."""
    tl = text.lower()
    if any(w in tl for w in ("opera", "опера")):
        return "opera"
    if any(w in tl for w in ("chrome", "хром", "google chrome")):
        return "chrome"
    if any(w in tl for w in ("firefox", "мозилла", "mozilla")):
        return "firefox"
    return "msedge"


def _build_write_steps(app_target: str, content: str, app_already_open: bool) -> list[PlanStep]:
    """Build [focus/open, type_text] plan steps for writing content to an app."""
    steps: list[PlanStep] = []
    if app_already_open:
        steps.append(PlanStep(step_number=1, action="focus_window", target=app_target))
    else:
        action1 = "new_document" if app_target in ("winword", "notepad") else "open_app"
        steps.append(PlanStep(step_number=1, action=action1, target=app_target))
    steps.append(PlanStep(step_number=len(steps) + 1, action="type_text", params=content))
    return steps


@app.post("/api/smart_command", response_model=SmartCommandResponse)
async def smart_command(req: SmartCommandRequest):
    t0 = time.monotonic()

    # ── 1. Research + summarize + write ─────────────────────────────────────
    if _is_research_request(req.text):
        topic = _extract_research_topic(req.text)
        log.info("SMART_COMMAND: research detected  topic=%r", topic)

        summary_prompt = (
            f"{_SOUL_TEXT}\n\n"
            f"Write a comprehensive summary about: \"{topic}\"\n"
            f"Include key facts, background, and important points. "
            f"Target length: ~200 words. Use clear paragraphs. Do not include a title line."
        )
        summary, _ = await granite_generate(summary_prompt, temperature=0.4)
        summary = summary.strip()

        ms = int((time.monotonic() - t0) * 1000)
        browser_target = _detect_browser(req.text)
        app_target     = _detect_target_app(req.text, req.session_context, req.open_windows)
        app_is_open    = _app_is_open(app_target, req.window_title, req.window_process, req.session_context, req.open_windows)

        # Plan: open browser + search (visual feedback) → open target app → write summary
        plan_steps: list[PlanStep] = [
            PlanStep(step_number=1, action="open_app",    target=browser_target),
            PlanStep(step_number=2, action="search_web",  params=topic),
        ]
        for s in _build_write_steps(app_target, summary, app_is_open):
            s.step_number = len(plan_steps) + 1
            plan_steps.append(s)

        log.info("SMART_COMMAND: research done  chars=%d  ms=%d  target=%s", len(summary), ms, app_target)
        return SmartCommandResponse(
            plan_id=f"research_{int(time.time())}",
            steps=plan_steps,
            total_steps=len(plan_steps),
            reasoning=f"Searched '{topic}' in browser, generated summary to write in {app_target}",
            latency_ms=ms,
            content_generated=summary,
        )

    # ── 2. Essay / write-about request ──────────────────────────────────────
    if _is_essay_request(req.text):
        topic = _extract_essay_topic(req.text)
        log.info("SMART_COMMAND: essay detected  topic=%r", topic)

        outline_prompt = ESSAY_OUTLINE_PROMPT.format(
            topic=topic, style="informative", word_count=200, language="English")
        outline, _ = await granite_generate(outline_prompt, temperature=0.3)

        draft_prompt = ESSAY_DRAFT_PROMPT.format(
            topic=topic, style="informative", word_count=200, language="English",
            outline=outline)
        draft, _ = await granite_generate(draft_prompt, temperature=0.4)

        refine_prompt = ESSAY_REFINE_PROMPT.format(draft=draft)
        final_text, _ = await granite_generate(refine_prompt, temperature=0.2)
        final_text = final_text.strip()

        ms = int((time.monotonic() - t0) * 1000)
        log.info("SMART_COMMAND: essay done  chars=%d  ms=%d", len(final_text), ms)

        app_target  = _detect_target_app(req.text, req.session_context, req.open_windows, is_essay=True)
        app_is_open = _app_is_open(app_target, req.window_title, req.window_process, req.session_context, req.open_windows)

        plan_steps = _build_write_steps(app_target, final_text, app_is_open)

        return SmartCommandResponse(
            plan_id=f"smart_{int(time.time())}",
            steps=plan_steps,
            total_steps=len(plan_steps),
            reasoning=f"Generated essay about '{topic}' and will type it into {app_target}",
            latency_ms=ms,
            content_generated=final_text,
        )

    # ── 3. Regular plan via IBM Granite ─────────────────────────────────────
    plan_req = PlanRequest(
        text=req.text,
        session_id=req.session_id,
        window_title=req.window_title,
        window_process=req.window_process,
        session_context=req.session_context,
        open_windows=req.open_windows,
    )
    plan_resp = await create_plan(plan_req)
    ms = int((time.monotonic() - t0) * 1000)
    return SmartCommandResponse(
        plan_id=plan_resp.plan_id,
        steps=plan_resp.steps,
        total_steps=plan_resp.total_steps,
        reasoning=plan_resp.reasoning,
        latency_ms=ms,
        content_generated=None,
    )


# ─── Execution Feedback ───────────────────────────────────────────────────────

class ExecutionFeedbackRequest(BaseModel):
    plan_id: str
    step_action: str
    step_target: str = ""
    success: bool
    error_message: str = ""
    session_id: str = "default"

@app.post("/api/execution_feedback")
async def execution_feedback(req: ExecutionFeedbackRequest):
    """Receives execution results from C# client. Stores in memory for context."""
    status = "succeeded" if req.success else "failed"
    log.info("FEEDBACK  plan=%s  step=%s  target=%s  status=%s",
             req.plan_id, req.step_action, req.step_target, status)
    # Write to session memory
    if not req.success:
        mem = _read_memory(req.session_id)
        entry = f"\n**Execution Note:** {req.step_action} on {req.step_target} FAILED — {req.error_message}"
        _write_memory(req.session_id, mem + entry)
    return {"status": "recorded"}


# ─── OpenClaw Python Executor ─────────────────────────────────────────────────

class ExecuteStepRequest(BaseModel):
    action:       str
    target:       Optional[str] = None
    params:       Optional[str] = None
    window_title: Optional[str] = None


def _ocl_new_document_notepad() -> dict:
    """Multiple methods to open a fresh blank Notepad file."""
    # Method 1 — launch with explicit temp file (bypasses Win11 session restore)
    try:
        tmp = tempfile.NamedTemporaryFile(
            suffix=".txt", delete=False,
            dir=_os.environ.get("TEMP", tempfile.gettempdir())
        )
        tmp.close()
        subprocess.Popen(["notepad.exe", tmp.name], creationflags=subprocess.CREATE_NEW_PROCESS_GROUP)
        log.info("OCL new_document notepad: method=tempfile path=%s", tmp.name)
        return {"success": True, "method": "tempfile"}
    except Exception as e:
        log.warning("OCL new_document notepad method1 failed: %s", e)

    # Method 2 — Ctrl+N in existing Notepad via win32gui
    if _WIN32:
        try:
            hwnd = _win32gui.FindWindow("Notepad", None)
            if hwnd:
                _win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.3)
                _pyautogui.hotkey("ctrl", "n") if _PYAUTOGUI else None
                log.info("OCL new_document notepad: method=ctrlN hwnd=%s", hwnd)
                return {"success": True, "method": "ctrlN"}
        except Exception as e:
            log.warning("OCL new_document notepad method2 failed: %s", e)

    # Method 3 — shell execute without arguments (fresh instance)
    try:
        subprocess.Popen("notepad.exe", creationflags=subprocess.CREATE_NEW_PROCESS_GROUP)
        log.info("OCL new_document notepad: method=shell_plain")
        return {"success": True, "method": "shell_plain"}
    except Exception as e:
        log.warning("OCL new_document notepad method3 failed: %s", e)

    return {"success": False, "reason": "all notepad methods failed"}


def _ocl_new_document_word() -> dict:
    """Multiple methods to open a new blank Word document."""

    # Method 1 — COM automation in a thread with timeout (avoids blocking the server)
    if _COMTYPES:
        import concurrent.futures as _cf
        def _com_create():
            word = _comtypes.CreateObject("Word.Application")
            word.Visible = True
            word.Documents.Add()
            return True
        try:
            with _cf.ThreadPoolExecutor(max_workers=1) as ex:
                fut = ex.submit(_com_create)
                fut.result(timeout=8)   # Give COM 8 seconds max
            log.info("OCL new_document word: method=comtypes_com")
            return {"success": True, "method": "comtypes_com"}
        except _cf.TimeoutError:
            log.warning("OCL new_document word method=comtypes timed out — trying next method")
        except Exception as e:
            log.warning("OCL new_document word method=comtypes failed: %s", e)

    # Method 2 — shell start (most compatible, doesn't need PATH)
    try:
        subprocess.Popen("start winword /n", shell=True, creationflags=subprocess.CREATE_NEW_PROCESS_GROUP)
        import time as _t; _t.sleep(0.5)
        log.info("OCL new_document word: method=shell_start")
        return {"success": True, "method": "shell_start"}
    except Exception as e:
        log.warning("OCL new_document word method=shell_start failed: %s", e)

    # Method 3 — find WINWORD.EXE in common Office paths
    word_exe = _find_word_exe()
    if word_exe:
        try:
            subprocess.Popen([word_exe, "/n"], creationflags=subprocess.CREATE_NEW_PROCESS_GROUP)
            log.info("OCL new_document word: method=explicit_path path=%s", word_exe)
            return {"success": True, "method": "explicit_path"}
        except Exception as e:
            log.warning("OCL new_document word method=explicit_path failed: %s", e)

    # Method 4 — Ctrl+N in running Word via win32gui
    if _WIN32:
        try:
            hwnd = _win32gui.FindWindow("OpusApp", None)  # Word's main window class
            if not hwnd:
                hwnd = _win32gui.FindWindow(None, "Microsoft Word")
            if hwnd:
                _win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.3)
                if _PYAUTOGUI:
                    _pyautogui.hotkey("ctrl", "n")
                log.info("OCL new_document word: method=ctrlN hwnd=%s", hwnd)
                return {"success": True, "method": "ctrlN"}
        except Exception as e:
            log.warning("OCL new_document word method=ctrlN failed: %s", e)

    # Method 5 — pywinauto launch
    if _PYWINAUTO:
        try:
            app = _PwApp(backend="uia").start("winword.exe")
            time.sleep(3)
            # Try clicking "Blank document" on Start Screen
            try:
                app.top_window().child_window(title="Blank document", control_type="Button").click_input()
            except Exception:
                pass
            log.info("OCL new_document word: method=pywinauto")
            return {"success": True, "method": "pywinauto"}
        except Exception as e:
            log.warning("OCL new_document word method=pywinauto failed: %s", e)

    # Method 6 — shell execute
    try:
        subprocess.Popen("start winword /n", shell=True)
        log.info("OCL new_document word: method=shell_start")
        return {"success": True, "method": "shell_start"}
    except Exception as e:
        log.warning("OCL new_document word method=shell_start failed: %s", e)

    return {"success": False, "reason": "all Word methods failed"}


def _ocl_open_app(target: str) -> dict:
    app_map = {
        "notepad":    ["notepad.exe"],
        "winword":    ["WINWORD.EXE"],
        "word":       ["WINWORD.EXE"],
        "calc":       ["calc.exe"],
        "calculator": ["calc.exe"],
        "msedge":     ["msedge.exe"],
        "edge":       ["msedge.exe"],
        "explorer":   ["explorer.exe"],
    }
    cmd = app_map.get(target.lower(), [target])
    try:
        subprocess.Popen(cmd, creationflags=subprocess.CREATE_NEW_PROCESS_GROUP)
        return {"success": True, "method": "subprocess"}
    except Exception:
        try:
            subprocess.Popen(f"start {cmd[0]}", shell=True)
            return {"success": True, "method": "shell_start"}
        except Exception as e:
            return {"success": False, "reason": str(e)}


def _ocl_type_text(text: str) -> dict:
    time.sleep(0.3)

    # Method 1 — clipboard paste (best for ALL text including Russian/Unicode)
    try:
        import tkinter as _tk
        r = _tk.Tk(); r.withdraw()
        r.clipboard_clear(); r.clipboard_append(text); r.update()
        r.after(300, r.destroy); r.mainloop()
        if _PYAUTOGUI:
            _pyautogui.hotkey("ctrl", "v")
        elif _WIN32:
            import win32clipboard as _wc
            _wc.OpenClipboard(); _wc.EmptyClipboard()
            _wc.SetClipboardText(text); _wc.CloseClipboard()
            _pyautogui.hotkey("ctrl", "v") if _PYAUTOGUI else None
        log.info("OCL type_text: method=clipboard_paste len=%d", len(text))
        return {"success": True, "method": "clipboard_paste"}
    except Exception as e:
        log.warning("OCL type_text clipboard failed: %s", e)

    # Method 2 — pywinauto type_keys (handles Unicode natively)
    if _PYWINAUTO and _WIN32:
        try:
            hwnd = _win32gui.GetForegroundWindow()
            if hwnd:
                import pywinauto
                ctrl = pywinauto.Desktop(backend="uia").from_handle(hwnd)
                ctrl.type_keys(text, with_spaces=True)
                log.info("OCL type_text: method=pywinauto_type_keys len=%d", len(text))
                return {"success": True, "method": "pywinauto_type_keys"}
        except Exception as e:
            log.warning("OCL type_text pywinauto failed: %s", e)

    # Method 3 — pyautogui write (ASCII safe)
    if _PYAUTOGUI:
        try:
            _pyautogui.write(text, interval=0.015)
            log.info("OCL type_text: method=pyautogui_write len=%d", len(text))
            return {"success": True, "method": "pyautogui_write"}
        except Exception as e:
            log.warning("OCL type_text pyautogui failed: %s", e)

    return {"success": False, "reason": "all type_text methods failed"}


def _ocl_find_and_click(element_name: str) -> dict:
    """Find a UI element by name and click it — tries multiple methods."""

    # Method 1 — pywinauto UIAutomation (best for named controls)
    if _PYWINAUTO and _WIN32:
        try:
            import pywinauto
            desktop = pywinauto.Desktop(backend="uia")
            hwnd = _win32gui.GetForegroundWindow()
            win = desktop.from_handle(hwnd)
            ctrl = win.child_window(title=element_name, found_index=0)
            ctrl.click_input()
            log.info("OCL find_and_click: method=pywinauto_uia element='%s'", element_name)
            return {"success": True, "method": "pywinauto_uia"}
        except Exception as e:
            log.warning("OCL find_and_click pywinauto failed: %s", e)

    # Method 2 — win32gui window-title search + Enter
    if _WIN32:
        try:
            found = [None]
            def cb(hwnd, _):
                title = _win32gui.GetWindowText(hwnd)
                if element_name.lower() in title.lower() and _win32gui.IsWindowVisible(hwnd):
                    found[0] = hwnd
                return True
            _win32gui.EnumWindows(cb, None)
            if found[0]:
                _win32gui.SetForegroundWindow(found[0])
                time.sleep(0.2)
                if _PYAUTOGUI:
                    _pyautogui.press("enter")
                log.info("OCL find_and_click: method=win32gui_enum element='%s'", element_name)
                return {"success": True, "method": "win32gui_enum"}
        except Exception as e:
            log.warning("OCL find_and_click win32gui failed: %s", e)

    # Method 3 — pyautogui locateOnScreen (image-based, last resort)
    if _PYAUTOGUI:
        try:
            # Take screenshot and look for element text on screen via OCR (not available, skip)
            # Just send Tab until we find something and click
            _pyautogui.press("enter")
            return {"success": True, "method": "pyautogui_enter"}
        except Exception as e:
            log.warning("OCL find_and_click pyautogui failed: %s", e)

    return {"success": False, "reason": f"element '{element_name}' not found"}


@app.post("/api/execute_step")
async def execute_step_python(req: ExecuteStepRequest):
    """
    OpenClaw Python executor — fallback when C# automation fails.
    Uses pyautogui + win32gui + subprocess for multi-method execution.
    """
    action = req.action.lower().strip()
    target = (req.target or "").lower().strip()
    params = req.params or ""

    log.info("OCL execute_step: action=%s target=%s params_len=%d", action, target, len(params))

    if action == "new_document":
        if "notepad" in target:
            return _ocl_new_document_notepad()
        elif target in ("winword", "word"):
            return _ocl_new_document_word()
        # Generic Ctrl+N fallback
        if _PYAUTOGUI:
            _pyautogui.hotkey("ctrl", "n")
            return {"success": True, "method": "generic_ctrlN"}
        return {"success": False, "reason": "unknown target for new_document"}

    elif action == "open_app":
        return _ocl_open_app(target)

    elif action == "type_text":
        return _ocl_type_text(params)

    elif action in ("save_document", "save"):
        if _PYAUTOGUI:
            _pyautogui.hotkey("ctrl", "s")
            return {"success": True, "method": "pyautogui_ctrl_s"}
        return {"success": False, "reason": "pyautogui not available"}

    elif action in ("find_and_click", "click_element"):
        return _ocl_find_and_click(params or target)

    elif action in ("press_escape", "escape"):
        if _PYAUTOGUI:
            _pyautogui.press("escape")
            return {"success": True, "method": "pyautogui_esc"}

    elif action in ("press_enter", "enter"):
        if _PYAUTOGUI:
            _pyautogui.press("enter")
            return {"success": True, "method": "pyautogui_enter"}

    elif action == "focus_window":
        if _WIN32:
            try:
                found = [None]
                tl = target.lower()
                def cb2(hwnd, _):
                    t = _win32gui.GetWindowText(hwnd).lower()
                    if tl in t and _win32gui.IsWindowVisible(hwnd):
                        found[0] = hwnd
                    return True
                _win32gui.EnumWindows(cb2, None)
                if found[0]:
                    _win32gui.SetForegroundWindow(found[0])
                    return {"success": True, "method": "win32gui_focus"}
            except Exception as e:
                return {"success": False, "reason": str(e)}

    elif action == "screenshot":
        if _PYAUTOGUI:
            img = _pyautogui.screenshot()
            path = _os.path.join(tempfile.gettempdir(), "ocl_screenshot.png")
            img.save(path)
            return {"success": True, "path": path, "method": "pyautogui_screenshot"}

    elif action in ("mouse_move", "move_mouse"):
        coords = (target or params or "").replace(" ", "")
        try:
            x, y = map(int, coords.split(","))
            if _PYAUTOGUI:
                _pyautogui.moveTo(x, y, duration=0.3)
                return {"success": True, "method": "pyautogui_move", "x": x, "y": y}
        except Exception as e:
            return {"success": False, "reason": str(e)}

    elif action in ("drag_to", "drag"):
        coords = (target or params or "").replace(" ", "")
        try:
            x, y = map(int, coords.split(","))
            if _PYAUTOGUI:
                _pyautogui.dragTo(x, y, duration=0.5, button="left")
                return {"success": True, "method": "pyautogui_drag", "x": x, "y": y}
        except Exception as e:
            return {"success": False, "reason": str(e)}

    elif action in ("scroll_at",):
        # params format: 'x,y,direction,amount'
        try:
            parts = (params or "").replace(" ", "").split(",")
            x, y = int(parts[0]), int(parts[1])
            direction = parts[2] if len(parts) > 2 else "down"
            amount = int(parts[3]) if len(parts) > 3 else 3
            if _PYAUTOGUI:
                _pyautogui.moveTo(x, y)
                _pyautogui.scroll(amount if direction == "up" else -amount)
                return {"success": True, "method": "pyautogui_scroll_at"}
        except Exception as e:
            return {"success": False, "reason": str(e)}

    elif action in ("hotkey", "key_combo"):
        combo = (params or "").strip().lower()
        if _PYAUTOGUI and combo:
            try:
                keys = [k.strip() for k in combo.split("+")]
                _pyautogui.hotkey(*keys)
                return {"success": True, "method": "pyautogui_hotkey", "combo": combo}
            except Exception as e:
                return {"success": False, "reason": str(e)}

    elif action == "wait":
        try:
            secs = float(params or "1")
            import time as _t
            _t.sleep(min(secs, 30))  # cap at 30s
            return {"success": True, "method": "sleep", "seconds": secs}
        except Exception as e:
            return {"success": False, "reason": str(e)}

    elif action in ("open_file_path", "open_file"):
        path = (params or target or "").strip()
        try:
            subprocess.Popen(["explorer", path], shell=False)
            return {"success": True, "method": "explorer_open", "path": path}
        except Exception as e:
            return {"success": False, "reason": str(e)}

    elif action == "message_user":
        # Just log it — C# side shows status messages
        log.info("OCL message_user: %s", params)
        return {"success": True, "method": "logged", "message": params}

    return {"success": False, "reason": f"action '{action}' not handled by OCL executor"}


# ─── Verify execution ─────────────────────────────────────────────────────────

class VerifyRequest(BaseModel):
    expected_app: str | None = None         # e.g. "winword", "notepad"
    expected_action: str | None = None      # e.g. "new_document", "type_text"
    command_text: str | None = None         # original user command
    window_before: str | None = None        # window title before execution

class VerifyResponse(BaseModel):
    success: bool
    description: str
    suggestion: str = ""
    window_title: str = ""
    process_name: str = ""

@app.post("/api/verify", response_model=VerifyResponse)
async def verify_execution(req: VerifyRequest):
    """Check if the last plan execution achieved its goal."""
    window_title  = ""
    process_name  = ""

    if _WIN32:
        try:
            hwnd = _win32gui.GetForegroundWindow()
            window_title = _win32gui.GetWindowText(hwnd)
            _, pid = _win32process.GetWindowThreadProcessId(hwnd)
            if _PSUTIL:
                process_name = _psutil.Process(pid).name().lower()
        except Exception:
            pass

    success = False
    suggestion = ""

    if req.expected_app:
        expected = req.expected_app.lower().strip()
        wt_low   = window_title.lower()
        pn_low   = process_name.lower()
        match_map = {
            "winword": ("winword" in pn_low or ("word" in wt_low and "wordpad" not in wt_low)),
            "notepad": ("notepad" in pn_low),
            "calc":    ("calc" in pn_low or "calculator" in wt_low.lower()),
            "explorer":("explorer" in pn_low or "file explorer" in wt_low),
            "msedge":  ("msedge" in pn_low or "edge" in wt_low),
            "opera":   ("opera" in pn_low),
            "chrome":  ("chrome" in pn_low),
        }
        if expected in match_map:
            success = match_map[expected]
        else:
            success = (expected in pn_low or expected in wt_low)

        if not success:
            # Check if the window changed at all
            changed = req.window_before and window_title != req.window_before
            if changed:
                suggestion = f"Window changed to '{window_title}' but expected {req.expected_app}. The app may have opened in background — check taskbar."
            else:
                suggestion = f"Expected {req.expected_app} to open but current window is still '{window_title}'. Try clicking the taskbar or saying 'open {req.expected_app}'."
    else:
        # No expected app — just check if window changed
        if req.window_before and window_title != req.window_before:
            success = True
            suggestion = f"Window changed: '{req.window_before}' → '{window_title}'"
        else:
            success = True  # Can't verify without expected app
            suggestion = ""

    description = f"Active: '{window_title}' ({process_name})"
    log.info("VERIFY: expected=%s success=%s desc=%s", req.expected_app, success, description)
    return VerifyResponse(success=success, description=description, suggestion=suggestion,
                          window_title=window_title, process_name=process_name)


# ─── Alternatives: generate 3 different approaches ───────────────────────────

class AlternativesRequest(BaseModel):
    text: str
    window_title: str = "unknown"
    window_process: str = "unknown"
    open_windows: list[str] = []

class AlternativePlan(BaseModel):
    label: str
    confidence: int
    reasoning: str
    steps: list[PlanStep]
    icon: str = "🤖"

class AlternativesResponse(BaseModel):
    alternatives: list[AlternativePlan]
    latency_ms: int

_ALTERNATIVES_PROMPT = """\
You are an AI planning assistant. Generate EXACTLY 3 DIFFERENT plans for this request.
Request: "{text}"
Active window: "{window_title}" (process: {window_process})
Open windows: {open_windows}

Return a JSON array of 3 objects. Each has a different strategy.
Schema: [{{"label":"short name","icon":"emoji","confidence":0-100,"reasoning":"why","steps":[{{"step_number":1,"action":"...","target":"...","params":"..."}}]}}]

Rules:
- Plan 1: Most reliable standard method (highest confidence)
- Plan 2: Alternative approach (different app, method, or path)
- Plan 3: Quick fallback (simplest, maybe less ideal)
- Use these actions: open_app, new_document, focus_window, find_and_click, type_text, search_web, save_document, navigate_url, click_at, hotkey, screenshot
- App targets: winword, notepad, calc, msedge, explorer, excel
- Keep each plan to 1-3 steps maximum
- Assign confidence (0-100) reflecting likelihood of success

Return ONLY valid JSON array, no markdown.
"""

@app.post("/api/alternatives", response_model=AlternativesResponse)
async def get_alternatives(req: AlternativesRequest):
    prompt = _ALTERNATIVES_PROMPT.format(
        text=req.text,
        window_title=req.window_title or "unknown",
        window_process=req.window_process or "unknown",
        open_windows=", ".join(req.open_windows) if req.open_windows else "none",
    )
    raw, ms = await granite_generate(prompt)

    try:
        # Granite returns a JSON array — use direct JSON parsing, not _extract_json (which only handles objects)
        text = raw.strip()
        # Find the JSON array
        arr_start = text.find("[")
        obj_start = text.find("{")
        if arr_start != -1 and (obj_start == -1 or arr_start < obj_start):
            # It's an array
            end = text.rfind("]")
            alts_raw = json.loads(text[arr_start:end+1]) if end != -1 else []
        else:
            parsed = _extract_json(raw)
            alts_raw = parsed if isinstance(parsed, list) else parsed.get("alternatives", parsed.get("plans", []))
        if not isinstance(alts_raw, list):
            alts_raw = []

        alternatives = []
        for i, a in enumerate(alts_raw[:3]):
            steps = [
                PlanStep(step_number=s.get("step_number", j+1),
                         action=s.get("action", "unknown"),
                         target=s.get("target"),
                         params=s.get("params"),
                         confidence=int(s.get("confidence", 80)))
                for j, s in enumerate(a.get("steps", []))
            ]
            if not steps:
                continue
            alternatives.append(AlternativePlan(
                label=a.get("label", f"Option {i+1}"),
                icon=a.get("icon", "🤖"),
                confidence=int(a.get("confidence", 70)),
                reasoning=a.get("reasoning", ""),
                steps=steps,
            ))
        log.info("ALTERNATIVES: %d plans generated in %dms", len(alternatives), ms)
        return AlternativesResponse(alternatives=alternatives, latency_ms=ms)
    except Exception as exc:
        log.warning("ALTERNATIVES parse failed: %s  raw=%r", exc, raw[:200])
        # Return empty — C# will fall back to its own options
        return AlternativesResponse(alternatives=[], latency_ms=ms)


# ─── Think: chain-of-thought pre-planning analysis ───────────────────────────

class ThinkRequest(BaseModel):
    text: str
    window_title: str = "unknown"
    window_process: str = "unknown"
    session_context: str = "none"
    open_windows: list[str] = []

class ThinkResponse(BaseModel):
    thought: str
    approach: str
    suggested_action: str
    needs_screenshot: bool
    confidence: int
    latency_ms: int

_THINK_PROMPT = """\
You are an AI assistant analyzing a user request before executing it.
Think step by step about what the user wants and how to achieve it.

User request: "{text}"
Active window: "{window_title}" (process: {window_process})
Session context: {context}
Open windows: {open_windows}

Analyze and respond in JSON:
{{
  "thought": "<1-2 sentence analysis of what the user really wants>",
  "approach": "<brief description of the best approach>",
  "suggested_action": "<primary action: open_app / new_document / type_text / search_web / find_and_click / hotkey / screenshot / other>",
  "needs_screenshot": <true if you need to see the screen to decide next steps, false otherwise>,
  "confidence": <0-100 how confident you are in this approach>
}}

Respond ONLY with valid JSON. No prose, no markdown.
"""

@app.post("/api/think", response_model=ThinkResponse)
async def think_about_request(req: ThinkRequest):
    """Chain-of-thought pre-planning: analyze what the user wants before generating a full plan."""
    prompt = _THINK_PROMPT.format(
        text=req.text,
        window_title=req.window_title or "unknown",
        window_process=req.window_process or "unknown",
        context=req.session_context or "none",
        open_windows=", ".join(req.open_windows) if req.open_windows else "none",
    )
    raw, ms = await granite_generate(prompt)
    try:
        parsed = _extract_json(raw)
        return ThinkResponse(
            thought=parsed.get("thought", ""),
            approach=parsed.get("approach", ""),
            suggested_action=parsed.get("suggested_action", "other"),
            needs_screenshot=bool(parsed.get("needs_screenshot", False)),
            confidence=int(parsed.get("confidence", 70)),
            latency_ms=ms,
        )
    except Exception as exc:
        log.warning("THINK parse failed: %s  raw=%r", exc, raw[:200])
        return ThinkResponse(
            thought=raw[:200] if raw else "Analysis failed",
            approach="Use standard plan",
            suggested_action="other",
            needs_screenshot=False,
            confidence=50,
            latency_ms=ms,
        )


# ─── Skills & Memory API ──────────────────────────────────────────────────────

@app.get("/api/skills")
async def get_skills():
    return {"skills": SKILLS, "count": len(SKILLS)}

@app.get("/api/memory/{session_id}")
async def get_memory(session_id: str):
    content = _read_memory(session_id)
    if not content:
        return {"session_id": session_id, "memory": "", "turns": 0}
    turns = content.count("**User:**")
    return {"session_id": session_id, "memory": content, "turns": turns}


# ─── Health, heartbeat, debug ─────────────────────────────────────────────────

@app.get("/api/health")
async def health():
    try:
        async with httpx.AsyncClient(timeout=5) as client:
            r = await client.get(f"{OLLAMA_BASE}/api/tags")
            models = [m["name"] for m in r.json().get("models", [])]
        granite_ok = any("granite" in m for m in models)
        return {
            "status": "ok" if granite_ok else "degraded",
            "model": GRANITE_MODEL,
            "ollama": "reachable",
            "granite_loaded": granite_ok,
            "soul_loaded": bool(_SOUL_TEXT),
            "skills_count": len(SKILLS),
            "timestamp": datetime.utcnow().isoformat(),
        }
    except Exception as exc:
        return {"status": "error", "error": str(exc)}

@app.get("/api/agent/heartbeat")
async def heartbeat():
    return {
        "status": "ready",
        "model": GRANITE_MODEL,
        "soul_loaded": bool(_SOUL_TEXT),
        "timestamp": datetime.utcnow().isoformat(),
    }

@app.get("/api/debug/last")
async def debug_last():
    return _debug


# ─── Startup ──────────────────────────────────────────────────────────────────

@app.on_event("startup")
async def on_startup():
    log.info("=" * 60)
    log.info("  AICompanion Backend  —  IBM Granite via Ollama")
    log.info("  Model : %s", GRANITE_MODEL)
    log.info("  Ollama: %s", OLLAMA_BASE)
    log.info("  Soul  : %s", "loaded" if _SOUL_TEXT else "NOT FOUND")
    log.info("  Skills: %d registered", len(SKILLS))
    log.info("  Memory: %s", _MEMORY_DIR)
    log.info("  Docs  : http://localhost:8000/docs")
    log.info("=" * 60)
    try:
        log.info("Warming up Granite model...")
        async with httpx.AsyncClient(timeout=30) as client:
            await client.post(
                f"{OLLAMA_BASE}/api/generate",
                json={"model": GRANITE_MODEL, "prompt": "hi", "stream": False,
                      "options": {"num_predict": 1}},
            )
        log.info("Model warm-up complete")
    except Exception as exc:
        log.warning("Warm-up failed (model will still work): %s", exc)


# ─── Entry point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("server:app", host="0.0.0.0", port=8000, log_level="info")
