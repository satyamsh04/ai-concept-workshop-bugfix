from fastapi import FastAPI

from app.core.config import settings
from app.core.logging import setup_logging
from app.api.routes import health, emails, drafts, auth, pipeline

setup_logging(debug=settings.debug)

app = FastAPI(title=settings.app_name, version="0.1.0")

app.include_router(health.router)
app.include_router(auth.router)
app.include_router(pipeline.router)
app.include_router(emails.router)
app.include_router(drafts.router)
