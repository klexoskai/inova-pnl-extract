from dataclasses import dataclass

import msal
import requests


GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]


@dataclass
class SharePointFile:
    relative_path: str
    content: bytes


class SharePointClient:
    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        site_id: str,
        drive_id: str,
        folder_path: str,
    ) -> None:
        self.site_id = site_id
        self.drive_id = drive_id
        self.folder_path = folder_path.strip("/")
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        self.app = msal.ConfidentialClientApplication(client_id=client_id, client_credential=client_secret, authority=authority)

    def _token(self) -> str:
        token_response = self.app.acquire_token_for_client(scopes=GRAPH_SCOPE)
        if "access_token" not in token_response:
            raise RuntimeError(f"Failed to acquire token: {token_response.get('error_description', 'unknown error')}")
        return token_response["access_token"]

    def _headers(self) -> dict[str, str]:
        return {"Authorization": f"Bearer {self._token()}"}

    def list_excel_items(self) -> list[dict]:
        url = (
            f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{self.drive_id}/root:/"
            f"{self.folder_path}:/children?$top=999"
        )
        items: list[dict] = []
        while url:
            response = requests.get(url, headers=self._headers(), timeout=30)
            response.raise_for_status()
            payload = response.json()
            for item in payload.get("value", []):
                if item.get("folder"):
                    items.extend(self._walk_folder(item["id"], prefix=item["name"]))
                else:
                    name = item.get("name", "")
                    if name.endswith((".xlsx", ".xls", ".xlsm")) and not name.startswith("~"):
                        item["relative_path"] = name
                        items.append(item)
            url = payload.get("@odata.nextLink")
        return items

    def _walk_folder(self, item_id: str, prefix: str) -> list[dict]:
        url = (
            f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{self.drive_id}/items/"
            f"{item_id}/children?$top=999"
        )
        output: list[dict] = []
        while url:
            response = requests.get(url, headers=self._headers(), timeout=30)
            response.raise_for_status()
            payload = response.json()
            for item in payload.get("value", []):
                path = f"{prefix}/{item['name']}"
                if item.get("folder"):
                    output.extend(self._walk_folder(item["id"], prefix=path))
                elif path.endswith((".xlsx", ".xls", ".xlsm")) and not item["name"].startswith("~"):
                    item["relative_path"] = path
                    output.append(item)
            url = payload.get("@odata.nextLink")
        return output

    def download_excel_files(self) -> list[SharePointFile]:
        output: list[SharePointFile] = []
        for item in self.list_excel_items():
            item_id = item["id"]
            rel_path = item["relative_path"]
            url = (
                f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{self.drive_id}/items/"
                f"{item_id}/content"
            )
            response = requests.get(url, headers=self._headers(), timeout=60)
            response.raise_for_status()
            output.append(SharePointFile(relative_path=rel_path, content=response.content))
        return output
