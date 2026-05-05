from app.domain.extractor import extract_from_bytes
from app.services.sharepoint_client import SharePointClient
from app.services.snowflake_client import SnowflakeClient


class PipelineService:
    def __init__(self, sharepoint_client: SharePointClient, snowflake_client: SnowflakeClient, target_table: str) -> None:
        self.sharepoint_client = sharepoint_client
        self.snowflake_client = snowflake_client
        self.target_table = target_table

    def run(self, dry_run: bool = False) -> dict:
        files = self.sharepoint_client.download_excel_files()
        outcome = extract_from_bytes((f.relative_path, f.content) for f in files)
        row_count = len(outcome.data.index)

        load_ok = True
        inserted_rows = 0
        if not dry_run and row_count > 0:
            load_ok, inserted_rows = self.snowflake_client.replace_table(outcome.data, self.target_table)

        return {
            "total_files_downloaded": len(files),
            "processed_files": outcome.processed_files,
            "skipped_files": outcome.skipped_files,
            "rows_extracted": row_count,
            "snowflake_load_ok": load_ok,
            "snowflake_rows_written": inserted_rows,
            "target_table": self.target_table,
            "dry_run": dry_run,
        }
