import os
import sys
from typing import Any, Dict, List, Optional
from urllib.parse import urljoin

from oauthlib.oauth2 import BackendApplicationClient, TokenExpiredError
from requests import Response
from requests_oauthlib import OAuth2Session

try:
    from docassemble.base.util import get_config
except ImportError:
    if not os.environ.get("GRAPH_IGNORE_DOCASSEMBLE_IMPORT", ""):
        print(
            "WARNING: Could not import docassemble. Client credentials must be "
            "specified directly.",
            file=sys.stderr)


class MSGraphSession(OAuth2Session):
    def __new__(cls, *args, **kwargs):
        self = super().__new__(cls)
        # We do this in order to facilitate pickling
        self.auth = lambda r: r
        return self

    def __init__(self,
                 api_version: str = "v1.0",
                 config_key: str = "msgraph",
                 tenant_id_key: str = "tenant id",
                 tenant_id: Optional[str] = None,
                 client_id_key: str = "client id",
                 client_id: Optional[str] = None,
                 client_secret_key: str = "client secret",
                 *args, **kwargs):
        self.api_version = api_version
        self.config_key = config_key
        self.client_secret_key = client_secret_key
        self.tenant_id = tenant_id or get_config(config_key).get(
            tenant_id_key)
        client_id = client_id or get_config(config_key).get(
            client_id_key)
        super().__init__(client=BackendApplicationClient(client_id=client_id),
                         *args, **kwargs)

    def __getstate__(self):
        self.auth = None
        return self.__dict__

    def fetch_token(self, client_secret: Optional[str] = None,
                    scope: str = "https://graph.microsoft.com/.default",
                    *args, **kwargs):
        secret = client_secret or \
                 get_config(self.config_key).get(self.client_secret_key)
        return super().fetch_token(
            f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token",
            client_secret=secret,
            scope=scope,
            *args, **kwargs
        )

    def request(self, method, url, *args, **kwargs):
        if not url.startswith("https://"):
            url = urljoin(f"https://graph.microsoft.com/{self.api_version}/",
                          url)
        return super().request(method, url, *args, **kwargs)

    ############################################################################
    #                           CONVENIENCE METHODS                            #
    ############################################################################

    def get_file_info(self,
                      drive_id: str,
                      item_id: str):
        return self.get(f'drives/{drive_id}/items/{item_id}')

    def upload_file(self,
                    drive_id: str,
                    parent_item_id: str,
                    filename: str,
                    file) -> Response:
        return self.put(
            f'drives/{drive_id}/items/{parent_item_id}:/{filename}:/content',
            file)

    def get_table(self,
                  drive_id: str,
                  item_id: str,
                  table_id: str):
        return self.get(f"drives/{drive_id}/items/{item_id}/workbook/"
                        f"tables/{table_id}")

    def get_table_headers(self,
                          drive_id: str,
                          item_id: str,
                          table_id: str) -> List[List[str]]:
        return self.get(f"drives/{drive_id}/items/{item_id}/workbook/"
                        f"tables/{table_id}/headerRowRange").json()["values"]

    def add_table_row(self,
                      drive_id: str,
                      item_id: str,
                      table_id: str,
                      row_data: List[Any]) -> Response:
        return self.post(f"drives/{drive_id}/items/{item_id}/workbook/"
                         f"tables/{table_id}/rows/add",
                         json={"values": [row_data]})

    def add_table_data(self,
                       drive_id: str,
                       item_id: str,
                       table_id: str,
                       data: Dict[str, Any]) -> Response:
        headers = self.get_table_headers(drive_id, item_id, table_id)[-1]
        normalized_data = {key.casefold(): value for key, value in data.items()}
        row = [
            normalized_data.get(header.casefold(), None) for header in headers
        ]
        return self.add_table_row(drive_id, item_id, table_id, row)
