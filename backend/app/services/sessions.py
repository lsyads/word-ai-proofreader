from __future__ import annotations

from dataclasses import dataclass
from datetime import UTC, datetime
from uuid import uuid4

from app.schemas import SessionResponse


@dataclass
class AISession:
    session_id: str
    created_at: str
    last_response_id: str | None = None


_sessions: dict[str, AISession] = {}


class SessionNotFoundError(KeyError):
    """Raised when a proofread request references an unknown AI session."""


def create_session() -> SessionResponse:
    session = AISession(
        session_id=f"session_{uuid4().hex}",
        created_at=datetime.now(UTC).isoformat(),
    )
    _sessions[session.session_id] = session

    return SessionResponse(session_id=session.session_id, created_at=session.created_at)


def get_last_response_id(session_id: str | None) -> str | None:
    if not session_id:
        return None

    session = _sessions.get(session_id)
    if not session:
        raise SessionNotFoundError(session_id)

    return session.last_response_id


def update_last_response_id(session_id: str | None, response_id: str | None) -> None:
    if not session_id or not response_id:
        return

    session = _sessions.get(session_id)
    if not session:
        raise SessionNotFoundError(session_id)

    session.last_response_id = response_id


def clear_sessions_for_tests() -> None:
    _sessions.clear()
