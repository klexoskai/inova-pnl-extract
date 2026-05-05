# P&L Extract FastAPI Service

FastAPI service that:
- reads Excel files from a SharePoint folder (via Microsoft Graph),
- runs the P&L extraction logic from the notebook,
- writes the final dataset into Snowflake (table replace/update behavior via `overwrite=True`).

## API Endpoints

- `GET /health`
- `POST /jobs/extract`
  - body: `{"dry_run": false}`
  - `dry_run=true` executes download + transform but skips Snowflake load.

## Local Run

1. Create env file:
   - `cp .env.example .env`
2. Install dependencies:
   - `pip install -r requirements.txt`
3. Start API:
   - `uvicorn app.main:app --reload`

## SharePoint Access

The app uses Microsoft Graph application permissions (client credential flow). Configure:
- `AZURE_TENANT_ID`
- `AZURE_CLIENT_ID`
- `AZURE_CLIENT_SECRET`
- `SHAREPOINT_SITE_ID`
- `SHAREPOINT_DRIVE_ID`
- `SHAREPOINT_FOLDER_PATH`

For Azure AD app permissions, assign Graph permissions that allow reading files from SharePoint drives.

## Snowflake Load

The service writes using `snowflake.connector.pandas_tools.write_pandas` with overwrite enabled.
Set:
- `SNOWFLAKE_ACCOUNT`
- `SNOWFLAKE_USER`
- `SNOWFLAKE_PASSWORD`
- `SNOWFLAKE_ROLE`
- `SNOWFLAKE_WAREHOUSE`
- `SNOWFLAKE_DATABASE`
- `SNOWFLAKE_SCHEMA`
- `SNOWFLAKE_TABLE`

## Docker

Build:
- `docker build -t pnl-extract-service .`

Run:
- `docker run --env-file .env -p 8000:8000 pnl-extract-service`

## Azure Container Deployment (high-level)

1. Build and push the container image to Azure Container Registry.
2. Deploy to Azure Container Apps or Azure Web App for Containers.
3. Add all environment variables as secrets/config in Azure.
4. Trigger using:
   - HTTP call to `POST /jobs/extract` from:
     - Azure Logic Apps,
     - Azure Function timer,
     - or an external scheduler.

If you want, the next step can be adding:
- managed identity + Key Vault secret loading,
- idempotent/incremental Snowflake merge (instead of overwrite),
- auth on the trigger endpoint.
