from __future__ import annotations

from dataclasses import dataclass
from datetime import UTC, datetime
from uuid import uuid4

from app.schemas import SessionResponse


@dataclass
class AISession:
    session_id: str
    created_at: str


_sessions: dict[str, AISession] = {}


def create_session() -> SessionResponse:
    session = AISession(
        session_id=f"session_{uuid4().hex}",
        created_at=datetime.now(UTC).isoformat(),
    )
    _sessions[session.session_id] = session

    return SessionResponse(session_id=session.session_id, created_at=session.created_at)


def clear_sessions_for_tests() -> None:
    _sessions.clear()
