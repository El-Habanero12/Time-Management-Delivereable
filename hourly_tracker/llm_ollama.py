from __future__ import annotations

import json
import shutil
import subprocess
from dataclasses import dataclass
from typing import Iterable, List, Optional

from .tagging import Suggestion


@dataclass
class OllamaStatus:
    installed: bool
    models: List[str]


def detect_ollama() -> OllamaStatus:
    exe = shutil.which("ollama")
    if not exe:
        return OllamaStatus(installed=False, models=[])

    try:
        proc = subprocess.run(
            [exe, "list"],
            capture_output=True,
            text=True,
            timeout=5,
            check=False,
        )
        models: List[str] = []
        for line in proc.stdout.splitlines()[1:]:
            parts = line.strip().split()
            if parts:
                models.append(parts[0])
        return OllamaStatus(installed=True, models=models)
    except Exception:
        return OllamaStatus(installed=True, models=[])


def _run_ollama(prompt: str, model: str, timeout_seconds: int) -> Optional[str]:
    exe = shutil.which("ollama")
    if not exe:
        return None
    try:
        proc = subprocess.run(
            [exe, "run", model],
            input=prompt,
            capture_output=True,
            text=True,
            timeout=timeout_seconds,
            check=False,
        )
        if proc.returncode != 0:
            return None
        return (proc.stdout or "").strip()
    except Exception:
        return None


def ollama_narrative_summary(
    model: str,
    timeout_seconds: int,
    summary_input: str,
) -> Optional[dict]:
    prompt = (
        "You are an offline assistant summarizing private time-tracking check-ins. "
        "Return strictly valid JSON with keys: narrative (string), suggestions (array of 3 strings).\n"
        "Keep it concise and practical.\n"
        f"Input:\n{summary_input}\n"
    )
    raw = _run_ollama(prompt, model, timeout_seconds)
    if not raw:
        return None
    try:
        start = raw.find("{")
        end = raw.rfind("}")
        if start == -1 or end == -1:
            return None
        return json.loads(raw[start : end + 1])
    except Exception:
        return None


def ollama_classify_category(
    activity: str,
    categories: Iterable[str],
    model: str,
    timeout_seconds: int,
) -> Optional[Suggestion]:
    cats = list(categories)
    prompt = (
        "Classify the activity into one of the provided categories. "
        "Return strictly valid JSON with keys: category (string), confidence (0-1 float).\n"
        f"Categories: {cats}\n"
        f"Activity: {activity}\n"
    )
    raw = _run_ollama(prompt, model, timeout_seconds)
    if not raw:
        return None
    try:
        start = raw.find("{")
        end = raw.rfind("}")
        if start == -1 or end == -1:
            return None
        data = json.loads(raw[start : end + 1])
        category = str(data.get("category", "")).strip()
        confidence = float(data.get("confidence", 0.0))
        if category not in cats:
            return None
        confidence = max(0.0, min(1.0, confidence))
        return Suggestion(category=category, confidence=confidence, source="ollama")
    except Exception:
        return None
