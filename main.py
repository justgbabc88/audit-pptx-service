"""
Revenue Leak Audit — PPTX Generation Service
=============================================
FastAPI wrapper around slide_templates.py.
Receives a data_dict via POST, runs generate(), returns .pptx bytes.

Deploy to Railway. No persistent storage needed.
"""
import os
import io
import time
import logging
from typing import Dict, Any

from fastapi import FastAPI, HTTPException, Header, Response
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from slide_templates import generate

# ── Logging ───────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("pptx-service")

# ── App setup ─────────────────────────────────────────────────────────────
app = FastAPI(
    title="Revenue Leak Audit PPTX Service",
    version="1.0.0",
    description="Generates audit decks from structured data_dict input.",
)

# CORS — tighten this to your Supabase edge function domain in production
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # TODO: restrict to your Supabase project URL
    allow_credentials=False,
    allow_methods=["POST", "GET"],
    allow_headers=["*"],
)

# ── Auth ──────────────────────────────────────────────────────────────────
# Simple shared-secret auth. The edge function passes this header.
SERVICE_SECRET = os.environ.get("SERVICE_SECRET")
if not SERVICE_SECRET:
    logger.warning("SERVICE_SECRET env var not set — refusing to start in production mode")


def require_auth(x_service_secret: str = Header(None)):
    if not SERVICE_SECRET:
        raise HTTPException(500, "Server not configured: SERVICE_SECRET missing")
    if x_service_secret != SERVICE_SECRET:
        raise HTTPException(401, "Invalid or missing x-service-secret header")


# ── Schemas ───────────────────────────────────────────────────────────────
class GenerateRequest(BaseModel):
    data_dict: Dict[str, Any]
    audit_id: str | None = None  # optional, for logging only


# ── Routes ────────────────────────────────────────────────────────────────
@app.get("/healthz")
def healthz():
    """Railway uses this for health checks. Keep it cheap."""
    return {"status": "ok", "service": "pptx-generator"}


@app.post("/generate")
def generate_deck(
    req: GenerateRequest,
    _auth: None = None,  # populated via dependency below
    x_service_secret: str = Header(None),
):
    """
    Generate a PPTX from a data_dict.
    Returns the raw .pptx bytes with the correct content type.
    """
    require_auth(x_service_secret)

    audit_id = req.audit_id or "unknown"
    started = time.time()
    logger.info(f"[{audit_id}] Starting generation")

    # ── Arithmetic verification ───────────────────────────────────────────
    # The guarantee surplus must be positive. If not, refuse before generating.
    try:
        guarantee = req.data_dict.get("guarantee", {})
        surplus = guarantee.get("surplus_amount")
        if surplus is not None and isinstance(surplus, (int, float)) and surplus <= 0:
            logger.warning(f"[{audit_id}] Refusing generation: non-positive surplus {surplus}")
            raise HTTPException(
                status_code=400,
                detail=f"Guarantee arithmetic invalid: surplus is {surplus}. Refusing to generate.",
            )
    except HTTPException:
        raise
    except Exception as e:
        logger.warning(f"[{audit_id}] Surplus check skipped: {e}")

    # ── Generate ──────────────────────────────────────────────────────────
    buffer = io.BytesIO()
    try:
        generate(req.data_dict, buffer)
    except KeyError as e:
        logger.error(f"[{audit_id}] Missing key in data_dict: {e}")
        raise HTTPException(400, f"Missing required data_dict key: {e}")
    except Exception as e:
        logger.exception(f"[{audit_id}] Generation failed")
        raise HTTPException(500, f"Generation failed: {type(e).__name__}: {e}")

    buffer.seek(0)
    pptx_bytes = buffer.getvalue()
    duration_ms = int((time.time() - started) * 1000)
    logger.info(f"[{audit_id}] Done in {duration_ms}ms, {len(pptx_bytes)} bytes")

    return Response(
        content=pptx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={
            "X-Generation-Duration-Ms": str(duration_ms),
            "X-Pptx-Size-Bytes": str(len(pptx_bytes)),
        },
    )
