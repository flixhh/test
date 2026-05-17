#!/usr/bin/env python3
"""
Downloads all 2026 Harvest invoices as PDFs via a Pipedream proxy.

Set PIPEDREAM_PROXY_URL to your Pipedream workflow's HTTP trigger URL,
e.g. https://eoXXXXXXXX.m.pipedream.net
"""

import os
import sys
import time
import base64
import requests
from pathlib import Path

PROXY_URL = os.environ.get("PIPEDREAM_PROXY_URL", "").rstrip("/")
OUTPUT_DIR = Path("invoices_2026")


def proxy_get(path, params=None, accept="application/json"):
    resp = requests.get(
        PROXY_URL + path,
        params=params,
        headers={"Accept": accept},
    )
    resp.raise_for_status()
    return resp.json()


def list_invoices_2026():
    invoices = []
    params = {"from": "2026-01-01", "to": "2026-12-31", "per_page": 100, "page": 1}

    while True:
        data = proxy_get("/v2/invoices", params=params)
        invoices.extend(data["invoices"])
        print(f"  Page {params['page']}: {len(data['invoices'])} invoices")
        if not data["next_page"]:
            break
        params["page"] = data["next_page"]

    return invoices


def download_pdf(invoice):
    invoice_id = invoice["id"]
    number = invoice.get("number", str(invoice_id))
    issued = invoice.get("issue_date", "unknown")
    client = invoice.get("client", {}).get("name", "unknown")
    safe_client = "".join(c if c.isalnum() or c in " -_" else "_" for c in client)
    filename = OUTPUT_DIR / f"{issued}_{safe_client}_{number}.pdf"

    if filename.exists():
        print(f"  Skipping {filename.name} (already exists)")
        return

    data = proxy_get(f"/v2/invoices/{invoice_id}.pdf", accept="application/pdf")
    pdf_bytes = base64.b64decode(data["body"])
    filename.write_bytes(pdf_bytes)
    print(f"  Downloaded {filename.name} ({len(pdf_bytes):,} bytes)")
    time.sleep(0.2)


def main():
    if not PROXY_URL:
        print("Error: set PIPEDREAM_PROXY_URL to your Pipedream workflow URL.")
        sys.exit(1)

    OUTPUT_DIR.mkdir(exist_ok=True)

    print("Fetching 2026 invoices via Pipedream proxy...")
    invoices = list_invoices_2026()
    print(f"Found {len(invoices)} invoices.\n")

    if not invoices:
        print("No invoices found for 2026.")
        sys.exit(0)

    print("Downloading PDFs...")
    for i, invoice in enumerate(invoices, 1):
        print(f"[{i}/{len(invoices)}]", end=" ")
        download_pdf(invoice)

    print(f"\nDone. PDFs saved to ./{OUTPUT_DIR}/")


if __name__ == "__main__":
    main()
