from functools import lru_cache

from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", env_file_encoding="utf-8", extra="ignore")

    app_name: str = "pnl-extract-service"
    app_env: str = "dev"

    # Microsoft Graph / SharePoint app credentials
    tenant_id: str = Field(default="", alias="AZURE_TENANT_ID")
    client_id: str = Field(default="", alias="AZURE_CLIENT_ID")
    client_secret: str = Field(default="", alias="AZURE_CLIENT_SECRET")
    sharepoint_site_id: str = Field(default="", alias="SHAREPOINT_SITE_ID")
    sharepoint_drive_id: str = Field(default="", alias="SHAREPOINT_DRIVE_ID")
    sharepoint_folder_path: str = Field(default="", alias="SHAREPOINT_FOLDER_PATH")

    # Snowflake connection
    snowflake_account: str = Field(default="", alias="SNOWFLAKE_ACCOUNT")
    snowflake_user: str = Field(default="", alias="SNOWFLAKE_USER")
    snowflake_password: str = Field(default="", alias="SNOWFLAKE_PASSWORD")
    snowflake_role: str = Field(default="", alias="SNOWFLAKE_ROLE")
    snowflake_warehouse: str = Field(default="", alias="SNOWFLAKE_WAREHOUSE")
    snowflake_database: str = Field(default="", alias="SNOWFLAKE_DATABASE")
    snowflake_schema: str = Field(default="", alias="SNOWFLAKE_SCHEMA")
    snowflake_table: str = Field(default="PANDL_MASTER_SKU", alias="SNOWFLAKE_TABLE")


@lru_cache
def get_settings() -> Settings:
    return Settings()
