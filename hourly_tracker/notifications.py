from __future__ import annotations

from typing import Optional

try:
    from win10toast import ToastNotifier  # type: ignore
except Exception:  # pragma: no cover - optional dependency behavior
    ToastNotifier = None  # type: ignore


class Notifier:
    def __init__(self) -> None:
        self._notifier: Optional[ToastNotifier] = ToastNotifier() if ToastNotifier else None

    def notify(self, title: str, message: str, duration: int = 5) -> None:
        if self._notifier:
            try:
                self._notifier.show_toast(title, message, duration=duration, threaded=True)
                return
            except Exception:
                pass
        # Fallback: no-op if toast fails.
        _ = (title, message, duration)
