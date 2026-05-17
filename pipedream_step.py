"""
Pipedream Python step — paste this into a Code step in your workflow.

Required env vars (set under Settings > Environment Variables in Pipedream):
  HARVEST_ACCESS_TOKEN
  HARVEST_ACCESS_ID

After this step runs, add a Google Drive "Upload File" action for each
path returned in steps.download_invoices.$return_value.files
"""

import os
import time
import requests


def handler(pd: "pipedream"):
    access_token = os.environ["HARVEST_ACCESS_TOKEN"].strip()
    account_id = os.environ["HARVEST_ACCESS_ID"].strip()

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Harvest-Account-Id": account_id,
        "User-Agent": "PipedreamHarvestDownloader/1.0",
    }

    # Fetch all 2026 invoices (paginated)
    invoices = []
    params = {"from": "2026-01-01", "to": "2026-12-31", "per_page": 100, "page": 1}

    while True:
        resp = requests.get(
            "https://api.harvestapp.com/v2/invoices", headers=headers, params=params
        )
        resp.raise_for_status()
        data = resp.json()
        invoices.extend(data["invoices"])
        if not data["next_page"]:
            break
        params["page"] = data["next_page"]

    # Download each PDF to /tmp (Pipedream's writable scratch space)
    os.makedirs("/tmp/invoices_2026", exist_ok=True)
    files = []

    for invoice in invoices:
        invoice_id = invoice["id"]
        number = invoice.get("number", str(invoice_id))
        issued = invoice.get("issue_date", "unknown")
        client = invoice.get("client", {}).get("name", "unknown")
        safe_client = "".join(c if c.isalnum() or c in " -_" else "_" for c in client)
        path = f"/tmp/invoices_2026/{issued}_{safe_client}_{number}.pdf"

        pdf_resp = requests.get(
            f"https://api.harvestapp.com/v2/invoices/{invoice_id}.pdf",
            headers={**headers, "Accept": "application/pdf"},
        )
        pdf_resp.raise_for_status()

        with open(path, "wb") as f:
            f.write(pdf_resp.content)

        files.append({"path": path, "name": os.path.basename(path)})
        time.sleep(0.2)  # stay within Harvest rate limits

    return {"count": len(files), "files": files}
