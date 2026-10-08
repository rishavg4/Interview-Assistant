"""
server.py  —  Interview Assistant Backend (FastAPI / Python)
Runs at http://127.0.0.1:3001
"""

import uuid
from datetime import datetime, timezone
from typing import Optional

import uvicorn
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from config import settings
from services import ai_service, excel_service, cache_service

# ── App ────────────────────────────────────────────────────────────────────
app = FastAPI(title="Interview Assistant API", version="2.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://127.0.0.1:5500",
        "http://localhost:5500",
        f"http://127.0.0.1:{settings.port}",
        f"http://localhost:{settings.port}",
    ],
    allow_methods=["GET", "POST"],
    allow_headers=["Content-Type"],
)

# Serve frontend from the same server (optional convenience)
import pathlib
FRONTEND_DIR = pathlib.Path(__file__).resolve().parents[1] / "frontend"
if FRONTEND_DIR.exists():
    app.mount("/app", StaticFiles(directory=str(FRONTEND_DIR), html=True), name="frontend")

ROLES = [
    "Test Manager",
    "Scrum Master",
    "Defect Manager",
    "SAP Tester",
    "SAP Test Manager",
    "Project Manager",
]

# ── Request / Response models ──────────────────────────────────────────────

class SessionStartRequest(BaseModel):
    role: str

class AnswerRequest(BaseModel):
    role: str
    question: str
    session_id: Optional[str] = None

class AnswerResponse(BaseModel):
    points: list[str]
    illustrated: str
    source: str         # 'cache' | 'cache_cross' | 'ai'
    matched_role: str   # which role's sheet the answer came from (same as requested role for AI)


# ── Routes ─────────────────────────────────────────────────────────────────

@app.get("/api/roles")
def get_roles():
    """Return the list of supported roles."""
    return {"roles": ROLES}


@app.get("/api/stats")
def get_stats():
    """Return per-role interview session counts."""
    return {"stats": cache_service.get_stats()}


@app.post("/api/session/start")
def start_session(body: SessionStartRequest):
    """
    Start a new interview session for the given role.
    Returns sessionId, interview count, and whether Excel cache data exists.
    """
    if body.role not in ROLES:
        raise HTTPException(status_code=400, detail="Invalid role")

    session_id = str(uuid.uuid4())
    interview_count = cache_service.get_interview_count(body.role)
    cache_service.increment_interview_count(body.role)
    cache_size = cache_service.get_cache_size(body.role)

    mode = "cache+ai" if cache_size > 0 else "ai"

    return {
        "sessionId": session_id,
        "interviewCount": interview_count + 1,
        "mode": mode,
        "cacheSize": cache_size,
    }


@app.post("/api/answer", response_model=AnswerResponse)
async def get_answer(body: AnswerRequest):
    """
    Answer an interview question.
    Strategy: Excel cache first (fast, free) → AI fallback if no match.
    All answers are saved to Excel for future cache hits.
    """
    if not body.role or not body.question:
        raise HTTPException(status_code=400, detail="role and question are required")

    points: list[str] = []
    illustrated: str = ""
    source: str = "ai"
    matched_role: str = body.role

    # ── Step 1: Check Excel cache (primary role, then cross-role fallback) ─
    cached = cache_service.find_similar_answer(body.role, body.question)
    if cached:
        points       = cached["points"]
        illustrated  = cached["illustrated"]
        matched_role = cached["matched_role"]
        source       = "cache" if matched_role == body.role else "cache_cross"

    # ── Step 2: Fall back to AI only if no cache match anywhere ───────────
    if source == "ai":
        try:
            result      = await ai_service.generate_answer(body.role, body.question)
            points      = result["points"]
            illustrated = result["illustrated"]
        except Exception as exc:
            raise HTTPException(
                status_code=500,
                detail="Failed to generate answer. Check your OPENAI_API_KEY."
            ) from exc

    # ── Step 3: Always persist to Excel for future cache ──────────────────
    excel_service.save_qa({
        "id":         str(uuid.uuid4()),
        "role":       body.role,
        "question":   body.question,
        "points":     points,
        "illustrated": illustrated,
        "source":     source,
        "session_id": body.session_id or "unknown",
        "timestamp":  datetime.now(timezone.utc).isoformat(),
    })

    return AnswerResponse(
        points=points,
        illustrated=illustrated,
        source=source,
        matched_role=matched_role,
    )


@app.get("/api/history/{role}")
def get_history(role: str):
    """Return all saved Q&A entries for a given role."""
    history = excel_service.get_history(role)
    return {"history": history}


# ── Entry point ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    uvicorn.run(
        "server:app",
        host=settings.host,
        port=settings.port,
        reload=False,
    )
