import os
import secrets

from fastapi import Header, HTTPException

_SERVICE_TOKEN = os.getenv("MCP_API_TOKEN")


def verify_service_token(authorization: str | None = Header(default=None)) -> None:
    """Guards service-to-service access (e.g. the ProjectTools MCP server) with a static bearer token."""
    if not _SERVICE_TOKEN:
        raise HTTPException(status_code=503, detail="MCP_API_TOKEN is not configured on the server")
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="Missing bearer token")
    token = authorization.removeprefix("Bearer ")
    if not secrets.compare_digest(token, _SERVICE_TOKEN):
        raise HTTPException(status_code=401, detail="Invalid bearer token")
