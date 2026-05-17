#!/usr/bin/env python3
import os
import sys
import time
import requests
from pathlib import Path

ACCESS_TOKEN = os.environ["HARVEST_ACCESS_TOKEN"].strip()
ACCOUNT_ID = os.environ["HARVEST_ACCESS_ID"].strip()

HEADERS = {
    "Authorization": f"Bearer {ACCESS_TOKEN}",
    "Harvest-Account-Id": ACCOUNT_ID,
    "User-Agent": "HarvestInvoiceDownloader/1.0",
}

BASE_URL = "https://api.harvestapp.com/v2"
OUTPUT_DIR = Path("invoices_2026")


def list_invoices_2026():
    invoices = []
    url = f"{BASE_URL}/invoices"
    params = {"from": "2026-01-01", "to": "2026-12-31", "per_page": 100, "page": 1}

    while url:
        resp = requests.get(url, headers=HEADERS, params=params)
        resp.raise_for_status()
        data = resp.json()
        invoices.extend(data["invoices"])
        print(f"  Fetched page {params['page']}: {len(data['invoices'])} invoices")

        if data["next_page"]:
            params["page"] = data["next_page"]
        else:
            break

    return invoices


def download_pdf(invoice):
    invoice_id = invoice["id"]
    invoice_number = invoice.get("number", str(invoice_id))
    issued = invoice.get("issue_date", "unknown")
    client_name = invoice.get("client", {}).get("name", "unknown")
    safe_client = "".join(c if c.isalnum() or c in " -_" else "_" for c in client_name)
    filename = OUTPUT_DIR / f"{issued}_{safe_client}_{invoice_number}.pdf"

    if filename.exists():
        print(f"  Skipping {filename.name} (already exists)")
        return

    url = f"{BASE_URL}/invoices/{invoice_id}.pdf"
    resp = requests.get(url, headers={**HEADERS, "Accept": "application/pdf"})
    resp.raise_for_status()

    filename.write_bytes(resp.content)
    print(f"  Downloaded {filename.name} ({len(resp.content):,} bytes)")
    time.sleep(0.2)  # stay within rate limits


def main():
    OUTPUT_DIR.mkdir(exist_ok=True)

    print("Fetching 2026 invoices from Harvest...")
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
