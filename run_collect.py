#!/usr/bin/env python3
"""Collecte quotidienne des factures Super PDP (Groupe Abilis) -> SharePoint.
Stdlib uniquement (urllib). Secrets fournis par variables d'environnement."""

import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request

SP_CLIENT_ID = os.environ["SUPERPDP_CLIENT_ID"]
SP_CLIENT_SECRET = os.environ["SUPERPDP_CLIENT_SECRET"]
SP_TENANT = os.environ["SP_TENANT"]
GRAPH_CLIENT_ID = os.environ["SP_CLIENT_ID"]
GRAPH_CLIENT_SECRET = os.environ["SP_CLIENT_SECRET"]
DRIVE = os.environ["SP_DRIVE_ID"]
ROOT = os.environ["SP_ROOT_FOLDER_ID"]
STATE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "state", "last_id.txt")
BASE_FOLDERS = ["GROUPE ABILIS", "Super PDP Groupe Abilis"]
SIREN_FALLBACK = "807897970"

REPORT = {"deposited": [], "facturx_ok": 0, "facturx_unavailable": [], "errors": []}


def http(url, data=None, headers=None, method=None):
    req = urllib.request.Request(url, data=data, headers=headers or {}, method=method)
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            return r.status, dict(r.headers), r.read()
    except urllib.error.HTTPError as e:
        return e.code, dict(e.headers), e.read()


def spdp_token():
    data = urllib.parse.urlencode({
        "grant_type": "client_credentials",
        "client_id": SP_CLIENT_ID,
        "client_secret": SP_CLIENT_SECRET,
    }).encode()
    st, _, body = http("https://api.superpdp.tech/oauth2/token", data=data,
                       headers={"Content-Type": "application/x-www-form-urlencoded"})
    if st != 200:
        raise RuntimeError(f"Token Super PDP {st}: {body[:300]}")
    return json.loads(body)["access_token"]


def spdp_get(path, token, accept="application/json"):
    return http(f"https://api.superpdp.tech{path}",
                headers={"Authorization": f"Bearer {token}", "Accept": accept})


def graph_token():
    data = urllib.parse.urlencode({
        "grant_type": "client_credentials",
        "client_id": GRAPH_CLIENT_ID,
        "client_secret": GRAPH_CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
    }).encode()
    st, _, body = http(
        f"https://login.microsoftonline.com/{SP_TENANT}/oauth2/v2.0/token", data=data,
        headers={"Content-Type": "application/x-www-form-urlencoded"})
    if st != 200:
        raise RuntimeError(f"Token Graph {st}: {body[:300]}")
    return json.loads(body)["access_token"]


def graph(token, url, data=None, method=None):
    return http(url, data=json.dumps(data).encode() if data is not None else None,
                headers={"Authorization": f"Bearer {token}",
                         "Content-Type": "application/json"},
                method=method)


def child_folder(token, parent_id, name):
    st, _, body = graph(token,
        f"https://graph.microsoft.com/v1.0/drives/{DRIVE}/items/{parent_id}/children",
        data={"name": name, "folder": {},
              "@microsoft.graph.conflictBehavior": "fail"}, method="POST")
    if st in (200, 201):
        return json.loads(body)["id"]
    if st == 409:  # existe deja -> re-GET
        st2, _, body2 = graph(token,
            f"https://graph.microsoft.com/v1.0/drives/{DRIVE}/items/{parent_id}/children")
        if st2 == 200:
            for it in json.loads(body2).get("value", []):
                if it.get("name") == name and "folder" in it:
                    return it["id"]
    raise RuntimeError(f"Dossier '{name}' : {st} {body[:300]}")


def upload(token, folder_id, filename, content):
    url = (f"https://graph.microsoft.com/v1.0/drives/{DRIVE}/items/"
           f"{folder_id}:/{urllib.parse.quote(filename)}:/content"
           "?@microsoft.graph.conflictBehavior=replace")
    st, _, body = http(url, data=content,
                       headers={"Authorization": f"Bearer {token}",
                                "Content-Type": "application/octet-stream"},
                       method="PUT")
    if st not in (200, 201):
        raise RuntimeError(f"Upload {filename} : {st} {body[:300]}")


def sanitize(s):
    return re.sub(r"[^A-Za-z0-9._-]", "_", (s or "").strip()) or "SANS_NUMERO"


def read_last_id():
    with open(STATE) as f:
        return int(f.read().strip())


def write_last_id(i):
    os.makedirs(os.path.dirname(STATE), exist_ok=True)
    with open(STATE, "w") as f:
        f.write(str(i))


def main():
    last_id = read_last_id()
    token = spdp_token()
    gtoken = graph_token()

    # parent des SIRET = racine > GROUPE ABILIS > Super PDP Groupe Abilis
    parent = ROOT
    for name in BASE_FOLDERS:
        parent = child_folder(gtoken, parent, name)

    invoices, after, has_after = [], last_id, True
    while has_after:
        st, _, body = spdp_get(
            f"/v1.beta/invoices?direction=in&starting_after_id={after}&limit=100", token)
        if st != 200:
            raise RuntimeError(f"Liste factures {st}: {body[:300]}")
        page = json.loads(body)
        page_items = page.get("invoices") or page.get("data") or []
        invoices.extend(page_items)
        has_after = page.get("has_after", False)
        if has_after and page_items:
            after = page_items[-1]["id"]
        elif has_after:
            has_after = False

    print(f"Factures a traiter (id > {last_id}) : {len(invoices)}")
    for inv in invoices:
        iid = inv["id"]
        try:
            st, _, body = spdp_get(f"/v1.beta/invoices/{iid}", token)
            if st != 200:
                REPORT["errors"].append(f"{iid}: detail {st}")
                continue
            detail = json.loads(body)
            en = detail.get("en_invoice") or detail
            ids = ((en.get("buyer") or {}).get("identifiers")) or []
            siret = next((i for i in ids if re.fullmatch(r"\d{14}", str(i))), None)
            folder_key = str(siret) if siret else SIREN_FALLBACK
            month = (en.get("issue_date") or "")[:7] or "SANS_DATE"
            number = sanitize(en.get("number") or detail.get("number") or "")

            month_folder = child_folder(gtoken, child_folder(gtoken, parent, folder_key), month)

            # 1) fichier depot (XML ou PDF tel que depose)
            st, hdr, content = spdp_get(f"/v1.beta/invoices/{iid}/download", token,
                                        accept="application/xml, application/pdf")
            if st != 200:
                REPORT["errors"].append(f"{iid}: download {st}")
                continue
            ext = "pdf" if content[:4] == b"%PDF" else "xml"
            fname = f"{number}_{iid}.{ext}"
            upload(gtoken, month_folder, fname, content)
            REPORT["deposited"].append(
                {"id": iid, "file": fname, "folder": f"{folder_key}/{month}"})

            # 2) PDF officiel factur-x
            if ext == "pdf":
                REPORT["facturx_ok"] += 1  # depot deja PDF : un seul fichier suffit
            else:
                st, hdr, pdf = spdp_get(
                    f"/v1.beta/invoices/{iid}?format=factur-x", token,
                    accept="application/pdf")
                if st == 200 and pdf[:4] == b"%PDF":
                    upload(gtoken, month_folder, f"{number}_{iid}.pdf", pdf)
                    REPORT["facturx_ok"] += 1
                else:
                    REPORT["facturx_unavailable"].append(iid)

            write_last_id(iid)
        except Exception as e:  # une facture en echec ne bloque pas la collecte
            REPORT["errors"].append(f"{iid}: {e}")

    print("=== RAPPORT ===")
    print(f"Factures deposees : {len(REPORT['deposited'])}")
    for d in REPORT["deposited"]:
        print(f"  - {d['file']} -> {d['folder']}")
    print(f"PDF factur-x deposes : {REPORT['facturx_ok']}")
    print(f"PDF factur-x indisponibles : {REPORT['facturx_unavailable']}")
    print(f"Erreurs : {REPORT['errors'] or 'aucune'}")
    if REPORT["errors"]:
        sys.exit(1)


if __name__ == "__main__":
    main()
