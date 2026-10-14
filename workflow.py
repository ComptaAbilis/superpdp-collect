# ============================================================
# FICHIER 1/3 : workflow.py
# ============================================================

"""Workflow Mistral : collecte des factures recues via l'API Super PDP.

Recupere toutes les factures recues (direction=in) depuis un identifiant
de facture donne (starting_after_id), telecharge chaque fichier depose
(XML CII/UBL ou PDF Factur-X) et retourne le tout en base64 pour que
l'agent Vibe les depose dans SharePoint, classes par SIRET puis par mois.

API Super PDP v1.beta :
- Token   : POST https://api.superpdp.tech/oauth2/token (client_credentials)
- Liste   : GET  /v1.beta/invoices?direction=in&starting_after_id=... (100/page, has_after)
- Fichier : GET  /v1.beta/invoices/{id}/download (Accept: application/xml, application/pdf)
"""

import base64
import re
from datetime import timedelta
from typing import Any

import httpx
import pydantic
import mistralai.workflows as workflows

SUPERPDP_ENDPOINT = "https://api.superpdp.tech"
SIRET_RE = re.compile(r"\b\d{14}\b")


class CollectInput(pydantic.BaseModel):
    client_id: str
    client_secret: str
    starting_after_id: int = 0
    max_files: int = 100


def _get_token(client_id: str, client_secret: str) -> str:
    r = httpx.post(
        f"{SUPERPDP_ENDPOINT}/oauth2/token",
        data={
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
        },
        timeout=30,
    )
    r.raise_for_status()
    return r.json()["access_token"]


def _list_new_invoices(token: str, starting_after_id: int, max_files: int) -> list:
    headers = {"Authorization": f"Bearer {token}"}
    invoices = []
    cursor = starting_after_id
    while len(invoices) < max_files:
        r = httpx.get(
            f"{SUPERPDP_ENDPOINT}/v1.beta/invoices",
            params={"direction": "in", "starting_after_id": cursor, "limit": 100},
            headers=headers,
            timeout=30,
        )
        r.raise_for_status()
        body = r.json()
        data = body.get("data", [])
        invoices.extend(data)
        if not body.get("has_after") or not data:
            break
        cursor = data[-1]["id"]
    return invoices[:max_files]


def _guess_siret(invoice: dict, xml_text) -> str:
    for key in ("buyer", "receiver", "customer"):
        party = invoice.get(key) or {}
        if isinstance(party, dict):
            for ident in party.get("identifiers", []):
                val = str(ident.get("value", ""))
                if len(val) == 14 and val.isdigit():
                    return val
    ea = str(invoice.get("buyer_electronic_address", "") or invoice.get("electronic_address", "") or "")
    m = SIRET_RE.search(ea)
    if m:
        return m.group(0)
    if xml_text:
        m = SIRET_RE.search(xml_text)
        if m:
            return m.group(0)
    return ""


def _filename(invoice: dict, content_type: str) -> str:
    inv_id = invoice.get("id")
    ext = "pdf" if "pdf" in content_type else "xml"
    num = invoice.get("number") or invoice.get("invoice_number")
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", str(num)) if num else f"facture_{inv_id}"
    return f"{safe}_{inv_id}.{ext}"


@workflows.activity(name="collect-invoices", start_to_close_timeout=timedelta(minutes=10))
async def collect_invoices(payload: CollectInput) -> dict:
    token = _get_token(payload.client_id, payload.client_secret)
    headers = {"Authorization": f"Bearer {token}"}
    invoices = _list_new_invoices(token, payload.starting_after_id, payload.max_files)
    results = []
    errors = []
    max_id = payload.starting_after_id
    for inv in invoices:
        try:
            r = httpx.get(
                f"{SUPERPDP_ENDPOINT}/v1.beta/invoices/{inv['id']}/download",
                headers={**headers, "Accept": "application/xml, application/pdf"},
                timeout=60,
            )
            r.raise_for_status()
            content_type = r.headers.get("content-type", "application/xml")
            raw = r.content
            xml_text = raw.decode("utf-8", errors="ignore") if "xml" in content_type else None
            results.append({
                "id": inv["id"],
                "filename": _filename(inv, content_type),
                "content_type": content_type,
                "created_at": inv.get("created_at") or inv.get("issue_date"),
                "siret": _guess_siret(inv, xml_text),
                "base64": base64.b64encode(raw).decode("ascii"),
            })
            max_id = max(max_id, inv["id"])
        except Exception as exc:
            errors.append({"id": inv.get("id"), "error": str(exc)[:300]})
            max_id = max(max_id, inv.get("id", max_id))
    return {"invoices": results, "max_id": max_id, "errors": errors, "count": len(results)}


@workflows.workflow.define(name="superpdp-collect")
class SuperPdpCollectWorkflow:
    @workflows.workflow.entrypoint
    async def run(self, payload: CollectInput) -> dict:
        return await collect_invoices(payload)
