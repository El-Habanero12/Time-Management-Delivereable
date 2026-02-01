from __future__ import annotations

import json
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional, Tuple

WordCounts = Dict[str, Dict[str, int]]


BASELINE_KEYWORDS: Dict[str, str] = {
    "meeting": "Work",
    "standup": "Work",
    "email": "Admin",
    "inbox": "Admin",
    "plan": "Admin",
    "planning": "Admin",
    "study": "Study",
    "homework": "Study",
    "lecture": "Study",
    "class": "Study",
    "gym": "Exercise",
    "workout": "Exercise",
    "run": "Exercise",
    "walk": "Exercise",
    "break": "Break",
    "lunch": "Break",
    "nap": "Break",
    "clean": "Chores",
    "laundry": "Chores",
    "dishes": "Chores",
    "call": "Social",
    "friends": "Social",
    "family": "Social",
}

BASELINE_REGEX: List[Tuple[re.Pattern[str], str]] = [
    (re.compile(r"\b(code|coding|debug|implement|pr|review)\b", re.IGNORECASE), "Work"),
    (re.compile(r"\b(read|reading|research|write|writing)\b", re.IGNORECASE), "Study"),
    (re.compile(r"\b(commute|driving|drive|transit)\b", re.IGNORECASE), "Break"),
]


@dataclass
class Suggestion:
    category: str
    confidence: float
    source: str


class CategorySuggester:
    def __init__(
        self,
        learned_rules_path: Path,
        categories: Iterable[str],
        llm_classifier: Optional[Callable[[str, List[str]], Optional[Suggestion]]] = None,
        llm_enabled: bool = False,
    ) -> None:
        # Normalise to Path in case callers pass strings
        self.learned_rules_path = Path(learned_rules_path)
        self.categories = list(categories)
        self.llm_classifier = llm_classifier
        self.llm_enabled = llm_enabled
        self.learned_keywords: WordCounts = {}
        self._load()

    def _load(self) -> None:
        if not self.learned_rules_path.exists():
            return
        try:
            data = json.loads(self.learned_rules_path.read_text(encoding="utf-8"))
            self.learned_keywords = data.get("learned_keywords", {})
        except Exception:
            self.learned_keywords = {}

    def _save(self) -> None:
        self.learned_rules_path.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "learned_keywords": self.learned_keywords,
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        }
        self.learned_rules_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    @staticmethod
    def _tokenize(text: str) -> List[str]:
        tokens = re.findall(r"[A-Za-z]{4,}", text.lower())
        return tokens[:20]

    def _learned_suggestion(self, activity: str) -> Optional[Suggestion]:
        tokens = self._tokenize(activity)
        scores: Dict[str, int] = {}
        for token in tokens:
            cat_counts = self.learned_keywords.get(token)
            if not cat_counts:
                continue
            for cat, count in cat_counts.items():
                scores[cat] = scores.get(cat, 0) + int(count)
        if not scores:
            return None
        best_cat = max(scores, key=scores.get)
        total = sum(scores.values())
        confidence = (scores[best_cat] / total) if total else 0.5
        return Suggestion(category=best_cat, confidence=min(0.95, max(0.5, confidence)), source="learned")

    def _regex_suggestion(self, activity: str) -> Optional[Suggestion]:
        for pattern, cat in BASELINE_REGEX:
            if pattern.search(activity):
                return Suggestion(category=cat, confidence=0.75, source="regex")
        return None

    def _keyword_suggestion(self, activity: str) -> Optional[Suggestion]:
        text = activity.lower()
        for keyword, cat in BASELINE_KEYWORDS.items():
            if keyword in text:
                return Suggestion(category=cat, confidence=0.65, source="keyword")
        return None

    def suggest(self, activity: str) -> Optional[Suggestion]:
        activity = activity.strip()
        if not activity:
            return None

        learned = self._learned_suggestion(activity)
        if learned and learned.category in self.categories:
            return learned

        regex = self._regex_suggestion(activity)
        if regex and regex.category in self.categories:
            return regex

        keyword = self._keyword_suggestion(activity)
        if keyword and keyword.category in self.categories:
            return keyword

        if self.llm_enabled and self.llm_classifier:
            llm_result = self.llm_classifier(activity, self.categories)
            if llm_result and llm_result.category in self.categories:
                return llm_result

        return None

    def learn_override(self, activity: str, chosen_category: str) -> None:
        if not activity.strip() or not chosen_category:
            return
        tokens = self._tokenize(activity)
        if not tokens:
            return
        for token in tokens:
            cat_counts = self.learned_keywords.setdefault(token, {})
            cat_counts[chosen_category] = int(cat_counts.get(chosen_category, 0)) + 1
        self._save()
