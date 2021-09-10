import datetime
import json

from docassemble.msgraph import MSGraphSession

tenant_id = "18109d70-b8cf-486f-b0c9-c6143ab021a6"
client_id = "2ea6b63c-a73a-486a-bc24-b396fc8fb988"
client_secret = "ACkjnr_~2KjzwB-w5E~a4.lQP-C5Ed1V-0"

drive_id = "b!AEOAy7g2BE-9Q_U0jJ7fSCXwDcLKBLtNoRJAHQDqk2YblYpFdguJQK1fXZaCc57I"
item_id = "01PRD5L72ME5KVSJGTFJCJCOV2B4FWROVS"
table_id = "{AC129F96-269F-44F5-9FAA-EBC8B6A31F4A}"

graph = MSGraphSession(tenant_id=tenant_id, client_id=client_id)
graph.fetch_token(client_secret=client_secret)
# print(graph.get_table_headers(drive_id, item_id, table_id)[-1])

# resp = graph.get(f"drives/{drive_id}/items/01PRD5L76K7WOZ7UUDSFDZFANJAAAAF4XO/children").json()
# with open("README.md", 'r') as file:
#    resp = graph.upload_file(drive_id, '01PRD5L76K7WOZ7UUDSFDZFANJAAAAF4XO',
#                             'test.md',
#                             file).json()
print(json.dumps(
        graph.get(f"drives/{drive_id}/items/{item_id}/workbook/"
                f"tables/{table_id}").json(),
    indent=2))
