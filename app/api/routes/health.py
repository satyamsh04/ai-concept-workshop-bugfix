from fastapi import APIRouter

from app.models.responses import HealthResponse

router = APIRouter()


@router.get("/health", response_model=HealthResponse, tags=["health"], include_in_schema=False)
async def health_check() -> HealthResponse:
    return HealthResponse(
        status="ok",
        services={"graph_api": "unchecked", "llm": "unchecked"},
    )
