from fastapi import APIRouter, Depends
from pydantic import BaseModel

from app.core.config import Settings, get_settings
from app.services.pipeline import PipelineService
from app.services.sharepoint_client import SharePointClient
from app.services.snowflake_client import SnowflakeClient

router = APIRouter()


class RunRequest(BaseModel):
    dry_run: bool = False


def build_pipeline(settings: Settings) -> PipelineService:
    sp_client = SharePointClient(
        tenant_id=settings.tenant_id,
        client_id=settings.client_id,
        client_secret=settings.client_secret,
        site_id=settings.sharepoint_site_id,
        drive_id=settings.sharepoint_drive_id,
        folder_path=settings.sharepoint_folder_path,
    )
    sf_client = SnowflakeClient(
        account=settings.snowflake_account,
        user=settings.snowflake_user,
        password=settings.snowflake_password,
        role=settings.snowflake_role,
        warehouse=settings.snowflake_warehouse,
        database=settings.snowflake_database,
        schema=settings.snowflake_schema,
    )
    return PipelineService(sp_client, sf_client, settings.snowflake_table)


@router.get("/health")
def health() -> dict:
    return {"status": "ok"}


@router.post("/jobs/extract")
def run_extract(request: RunRequest, settings: Settings = Depends(get_settings)) -> dict:
    pipeline = build_pipeline(settings)
    return pipeline.run(dry_run=request.dry_run)
