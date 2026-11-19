"""Claude access through the user's own Claude Code login (no API key).

Uses the Claude Agent SDK, which drives the Claude Code CLI bundled with it. Each user signs in once
(`claude auth login`, also available as a button in the app) and calls then run under that login.
Tools are disabled: Claude is only used as a text/JSON generator.
"""
import asyncio
import json
import re
import subprocess
import sys
from pathlib import Path

import claude_agent_sdk as sdk
from claude_agent_sdk import AssistantMessage, ClaudeAgentOptions, ResultMessage, TextBlock, query

_BUNDLED = Path(sdk.__file__).parent / "_bundled"
CLI = next((p for p in (_BUNDLED / "claude.exe", _BUNDLED / "claude") if p.exists()), None)
NOFLAGS = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def status() -> dict:
    """{'loggedIn': bool, 'detail': str}"""
    if CLI is None:
        return {"loggedIn": False, "detail": "Bundled Claude Code CLI not found - reinstall claude-agent-sdk."}
    try:
        r = subprocess.run([str(CLI), "auth", "status"], capture_output=True, text=True, timeout=30, creationflags=NOFLAGS)
        d = json.loads(r.stdout)
        who = d.get("email") or d.get("account") or d.get("authMethod")
        return {"loggedIn": bool(d.get("loggedIn")), "detail": f"Signed in ({who})" if d.get("loggedIn") else "Not signed in to Claude."}
    except Exception as e:
        return {"loggedIn": False, "detail": f"Could not check Claude login: {e}"}


def login() -> None:
    """Start the browser sign-in flow (returns immediately; the user completes it in the browser)."""
    if CLI is None:
        raise RuntimeError("Bundled Claude Code CLI not found.")
    subprocess.Popen([str(CLI), "auth", "login"], stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                     stderr=subprocess.DEVNULL, creationflags=NOFLAGS)


async def _ask(prompt: str, model: str | None) -> str:
    opts = ClaudeAgentOptions(
        max_turns=1, tools=[], allowed_tools=[], setting_sources=[], model=model,
        system_prompt="You are a precise SAP S/4HANA implementation assistant. Output only what is asked, with no preamble.")
    text, err = [], None
    async for m in query(prompt=prompt, options=opts):
        if isinstance(m, AssistantMessage):
            text += [b.text for b in m.content if isinstance(b, TextBlock)]
        elif isinstance(m, ResultMessage) and m.is_error:
            err = getattr(m, "result", None) or "Claude returned an error"
    if err and not text:
        raise RuntimeError(err)
    return "".join(text)


def ask(prompt: str, model: str | None = None, timeout: int = 240) -> str:
    return asyncio.run(asyncio.wait_for(_ask(prompt, model), timeout))


def _parse_json(t: str) -> dict:
    """First JSON object in the text; tolerates preamble, code fences, // comments and trailing commas."""
    t = t.replace("```json", "").replace("```", "")
    start = t.index("{")
    try:
        return json.JSONDecoder().raw_decode(t[start:])[0]
    except json.JSONDecodeError:
        cleaned = re.sub(r"(?m)^\s*//.*$", "", t[start:])
        cleaned = re.sub(r",\s*([}\]])", r"\1", cleaned)
        return json.JSONDecoder().raw_decode(cleaned)[0]


def ask_json(prompt: str, model: str | None = None) -> dict:
    t = ask(prompt, model)
    try:
        return _parse_json(t)
    except Exception as e:
        retry = (prompt + "\n\nYour previous reply was not valid JSON (" + str(e) + "). Reply again with ONLY one valid JSON object: "
                 "no comments, no trailing commas, no text outside the braces.")
        return _parse_json(ask(retry, model))


if __name__ == "__main__":
    print(status())
    if "--test" in sys.argv:
        print(ask_json('Return {"ok": true}'))
