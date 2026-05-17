"""
Pipedream HTTP-triggered proxy step.

Create a new Pipedream workflow with an HTTP / Webhook trigger,
then paste this as the Python Code step.

The workflow URL will look like: https://eoXXXXXXXX.m.pipedream.net
Set that as PIPEDREAM_PROXY_URL in download_harvest_invoices.py.

Required Pipedream env vars:
  HARVEST_ACCESS_TOKEN
  HARVEST_ACCESS_ID
"""

import os
import base64
import requests


def handler(pd: "pipedream"):
    event = pd.steps["trigger"]["event"]
    path = event.get("path", "/v2/invoices")
    query = event.get("query", {})
    accept = event.get("headers", {}).get("accept", "application/json")

    access_token = os.environ["HARVEST_ACCESS_TOKEN"].strip()
    account_id = os.environ["HARVEST_ACCESS_ID"].strip()

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Harvest-Account-Id": account_id,
        "User-Agent": "PipedreamHarvestProxy/1.0",
        "Accept": accept,
    }

    resp = requests.get(
        f"https://api.harvestapp.com{path}", headers=headers, params=query
    )
    resp.raise_for_status()

    if "pdf" in accept:
        return {
            "body": base64.b64encode(resp.content).decode(),
            "encoding": "base64",
        }

    return resp.json()
