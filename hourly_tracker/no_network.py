from __future__ import annotations

import socket
from typing import Any, Tuple


class NetworkBlockedError(RuntimeError):
    pass


def _is_loopback(address: Tuple[Any, Any]) -> bool:
    host = str(address[0]).lower()
    return host in {"127.0.0.1", "localhost", "::1"}


def enforce_no_network(allow_loopback: bool = True) -> None:
    """Best-effort guardrail that blocks outbound sockets.

    This is not a security boundary, but it reduces accidental network calls.
    """
    original_create_connection = socket.create_connection
    original_connect = socket.socket.connect

    def guarded_create_connection(address, *args, **kwargs):  # type: ignore[no-untyped-def]
        if allow_loopback and _is_loopback(address):
            return original_create_connection(address, *args, **kwargs)
        raise NetworkBlockedError(f"Network disabled by configuration: {address}")

    def guarded_connect(self, address):  # type: ignore[no-untyped-def]
        if allow_loopback and _is_loopback(address):
            return original_connect(self, address)
        raise NetworkBlockedError(f"Network disabled by configuration: {address}")

    socket.create_connection = guarded_create_connection  # type: ignore[assignment]
    socket.socket.connect = guarded_connect  # type: ignore[assignment]
