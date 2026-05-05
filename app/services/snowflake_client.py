import pandas as pd
import snowflake.connector
from snowflake.connector.pandas_tools import write_pandas


class SnowflakeClient:
    def __init__(
        self,
        account: str,
        user: str,
        password: str,
        role: str,
        warehouse: str,
        database: str,
        schema: str,
    ) -> None:
        self._conn_params = {
            "account": account,
            "user": user,
            "password": password,
            "role": role,
            "warehouse": warehouse,
            "database": database,
            "schema": schema,
        }

    def replace_table(self, df: pd.DataFrame, table_name: str) -> tuple[bool, int]:
        if df.empty:
            return True, 0
        with snowflake.connector.connect(**self._conn_params) as conn:
            ok, _, rows, _ = write_pandas(conn, df, table_name=table_name, auto_create_table=True, overwrite=True)
            return bool(ok), int(rows)
