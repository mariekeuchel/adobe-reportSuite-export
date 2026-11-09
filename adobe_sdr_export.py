#!/usr/bin/env python3
"""
Adobe Analytics SDR-Export
==========================

Exportiert für eine oder mehrere Report Suites:
  * eVars inkl. Einstellungen (Allokation, Expiration, Typ, Merchandising,
    Binding Events), Props, List-Variablen, Success Events
    -> Adobe Analytics 2.0 API (/dimensions, /metrics)
    Hinweis: Die 1.4 API wurde von Adobe im August 2026 abgeschaltet. Die 2.0 API
    liefert Konfigurationsdetails nur für eVars, nicht für Props, Events und Listvars.
  * alle Dimensionen und Metriken inkl. Beschreibung, Kategorie, Tags
    -> Adobe Analytics 2.0 API (/dimensions, /metrics)
  * Segmente inkl. Scope (Hit/Visit/Visitor), Definition als lesbarer Text,
    verwendete Variablen, Owner, Tags, Freigaben
    -> Adobe Analytics 2.0 API (/segments)
  * Berechnete Metriken inkl. Formel als Text
    -> Adobe Analytics 2.0 API (/calculatedmetrics)

Ergebnis: eine Excel-Datei mit einem Tab pro Komponententyp und leeren
Kontextspalten zum Anreichern.

Authentifizierung: OAuth Server-to-Server (Adobe Developer Console).

Aufruf-Beispiele:
  python adobe_sdr_export.py --config config.json --list-rsids
  python adobe_sdr_export.py --config config.json --rsids rs_prod rs_dev
  python adobe_sdr_export.py --config config.json --rsids rs_prod --all-segments --out sdr.xlsx
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import requests

IMS_TOKEN_URL = "https://ims-na1.adobelogin.com/ims/token/v3"
API_BASE = "https://analytics.adobe.io"
DEFAULT_SCOPES = (
    "openid,AdobeID,read_organizations,"
    "additional_info.projectedProductContext,additional_info.job_function"
)

# Leere Spalten, die du später mit eigenem Kontext befüllen kannst
CONTEXT_COLUMNS = ["Business-Definition", "Fachlicher Owner", "Status", "Kommentar"]

SEGMENT_EXPANSION = (
    "reportSuiteName,ownerFullName,modified,tags,compatibility,definition,"
    "publishingStatus,definitionLastModified,categories,sharesFullName"
)
CALCMETRIC_EXPANSION = (
    "reportSuiteName,ownerFullName,modified,tags,definition,categories,sharesFullName"
)
DIM_METRIC_EXPANSION = "tags,allowedForReporting,categories"
# eVar-Einstellungen (nur für eVars befüllt), siehe
# https://developer.adobe.com/analytics-apis/docs/2.0/guides/endpoints/dimensions/
EVAR_EXPANSION = (
    "allocationType,expirationType,expirationCustomDays,dataType,"
    "bindingEvents,merchandisingSyntax,attributionModel"
)
EVAR_FIELDS = ["allocationType", "expirationType", "expirationCustomDays", "dataType",
               "merchandisingSyntax", "bindingEvents", "ecomAllocationType", "attributionModel"]


def log(msg: str) -> None:
    print(f"[{datetime.now():%H:%M:%S}] {msg}", file=sys.stderr)


# --------------------------------------------------------------------------- #
# API-Client
# --------------------------------------------------------------------------- #
class AdobeClient:
    def __init__(self, cfg: dict):
        self.client_id = cfg["client_id"]
        self.client_secret = cfg["client_secret"]
        self.scopes = cfg.get("scopes", DEFAULT_SCOPES)
        self.company_id = cfg.get("global_company_id") or None
        self.ims_org = cfg.get("ims_org_id")
        if self.company_id and self.company_id.endswith("@AdobeOrg"):
            # Häufige Verwechslung: IMS Org ID statt Global Company ID eingetragen
            log("Hinweis: global_company_id ist eine IMS Org ID – Company ID wird automatisch ermittelt.")
            self.ims_org, self.company_id = self.company_id, None
        self.session = requests.Session()
        self._token: str | None = None
        self._token_expiry = 0.0

    # --- Auth -------------------------------------------------------------- #
    def _access_token(self) -> str:
        if self._token and time.time() < self._token_expiry - 60:
            return self._token
        r = self.session.post(
            IMS_TOKEN_URL,
            data={
                "grant_type": "client_credentials",
                "client_id": self.client_id,
                "client_secret": self.client_secret,
                "scope": self.scopes,
            },
            timeout=30,
        )
        if r.status_code != 200:
            raise SystemExit(f"Token-Abruf fehlgeschlagen ({r.status_code}): {r.text}")
        data = r.json()
        self._token = data["access_token"]
        self._token_expiry = time.time() + int(data.get("expires_in", 86400))
        return self._token

    def _headers(self) -> dict:
        h = {
            "Authorization": f"Bearer {self._access_token()}",
            "x-api-key": self.client_id,
            "Accept": "application/json",
        }
        if self.company_id:
            h["x-proxy-global-company-id"] = self.company_id
        if self.ims_org:
            h["x-gw-ims-org-id"] = self.ims_org
        return h

    def _request(self, method: str, url: str, **kw) -> requests.Response:
        r = None
        for attempt in range(6):
            r = self.session.request(method, url, headers=self._headers(), timeout=120, **kw)
            if r.status_code in (429, 500, 502, 503, 504):
                wait = int(r.headers.get("Retry-After", 2 ** attempt))
                log(f"  HTTP {r.status_code} – neuer Versuch in {wait}s")
                time.sleep(wait)
                continue
            return r
        return r

    # --- Discovery --------------------------------------------------------- #
    def discover_company(self) -> str:
        if self.company_id:
            return self.company_id
        r = self._request("GET", f"{API_BASE}/discovery/me")
        if r.status_code != 200:
            raise SystemExit(f"Discovery fehlgeschlagen ({r.status_code}): {r.text[:500]}")
        companies = [
            c for org in r.json().get("imsOrgs", [])
            if not self.ims_org or org.get("imsOrgId") == self.ims_org
            for c in org.get("companies", [])
        ]
        if not companies:
            raise SystemExit("Keine Analytics-Company für diesen technischen Account gefunden.")
        if len(companies) > 1:
            log("Mehrere Companies gefunden – nutze die erste. Per 'global_company_id' in config.json festlegen:")
            for c in companies:
                log(f"   {c['globalCompanyId']}  {c.get('companyName')}")
        self.company_id = companies[0]["globalCompanyId"]
        log(f"Global Company ID: {self.company_id}")
        return self.company_id

    # --- 2.0 API ----------------------------------------------------------- #
    def get20(self, path: str, params: dict | None = None, fallback_expansion: str | None = None) -> Any:
        url = f"{API_BASE}/api/{self.company_id}{path}"
        params = dict(params or {})
        r = self._request("GET", url, params=params)
        if r.status_code == 400 and fallback_expansion and params.get("expansion") != fallback_expansion:
            log(f"  400 bei {path} mit erweiterter expansion – versuche Basis-Expansion: {r.text[:200]}")
            params["expansion"] = fallback_expansion
            r = self._request("GET", url, params=params)
        if r.status_code == 400 and "expansion" in params:
            # Falls Adobe einen Expansion-Wert nicht (mehr) kennt: ohne Expansion erneut
            log(f"  400 bei {path} mit expansion – versuche ohne: {r.text[:200]}")
            params.pop("expansion")
            r = self._request("GET", url, params=params)
        if r.status_code != 200:
            raise RuntimeError(f"GET {path} -> {r.status_code}: {r.text[:500]}")
        return r.json()

    def get20_paged(self, path: str, params: dict) -> list[dict]:
        out, page = [], 0
        while True:
            data = self.get20(path, {**params, "limit": 1000, "page": page})
            out.extend(data.get("content", []))
            if data.get("lastPage", True) or page + 1 >= data.get("totalPages", 1):
                return out
            page += 1


# --------------------------------------------------------------------------- #
# Segment- und Calculated-Metric-Definitionen lesbar machen
# --------------------------------------------------------------------------- #
CONTEXT_LABEL = {"hits": "Hit", "visits": "Visit", "visitors": "Visitor"}

OPERATORS = {
    "streq": "equals", "not-streq": "does not equal",
    "strlt": "<", "strgt": ">", "strle": "≤", "strge": "≥",
    "streq-in": "equals any of", "not-streq-in": "equals none of",
    "contains": "contains", "not-contains": "does not contain",
    "contains-any": "contains any of", "not-contains-any": "contains none of",
    "contains-all": "contains all of", "not-contains-all": "does not contain all of",
    "starts-with": "starts with", "not-starts-with": "does not start with",
    "ends-with": "ends with", "not-ends-with": "does not end with",
    "matches": "matches", "not-matches": "does not match",
    "eq": "=", "ne": "≠", "gt": ">", "lt": "<", "ge": "≥", "le": "≤",
    "eq-any": "equals any of", "not-eq-any": "equals none of",
    "eq-in": "equals any of", "not-eq-in": "equals none of",
}


class DefinitionRenderer:
    """Übersetzt Segment-/Calc-Metric-JSON in lesbaren Text."""

    def __init__(self, names: dict[str, str], segment_names: dict[str, str]):
        self.names = names              # "variables/evar5" -> "eVar5 Seitentyp"
        self.segment_names = segment_names

    def label(self, ref: Any) -> str:
        if isinstance(ref, dict):
            if "name" in ref:
                n = ref["name"]
                friendly = self.names.get(n)
                return f"{friendly} [{n.split('/')[-1]}]" if friendly else n
            if ref.get("func") == "segment" and "segment_id" in ref:
                return self._seg_ref(ref["segment_id"])
            return self.render(ref)
        return str(ref)

    def _seg_ref(self, sid: str) -> str:
        return f"Segment »{self.segment_names.get(sid, sid)}«"

    @staticmethod
    def _value(p: dict) -> str:
        for key in ("str", "num", "list", "glob"):
            if key in p:
                v = p[key]
                return ", ".join(map(str, v)) if isinstance(v, list) else repr(v) if key == "str" else str(v)
        return ""

    # Segment ----------------------------------------------------------------
    def render(self, node: Any, depth: int = 0) -> str:
        if not isinstance(node, dict):
            return str(node)
        f = node.get("func", "")

        if f == "segment":
            if "container" in node:
                return self.render(node["container"], depth)
            if "segment_id" in node:
                return self._seg_ref(node["segment_id"])
        if f == "container":
            ctx = CONTEXT_LABEL.get(node.get("context"), node.get("context", "?"))
            return f"{ctx}-Container [ {self.render(node.get('pred'), depth + 1)} ]"
        if f in ("and", "or"):
            parts = [self.render(p, depth + 1) for p in node.get("preds", [])]
            joined = f" {f.upper()} ".join(parts)
            return f"({joined})" if depth and len(parts) > 1 else joined
        if f == "without":
            return f"EXCLUDE ( {self.render(node.get('pred'), depth + 1)} )"
        if f == "not":
            return f"NOT ( {self.render(node.get('pred'), depth + 1)} )"
        if f.startswith("sequence"):
            items = [self.render(s, depth + 1) for s in node.get("stream", []) or node.get("checkpoints", [])]
            prefix = {"sequence-prefix": "ONLY BEFORE SEQUENCE: ", "sequence-suffix": "ONLY AFTER SEQUENCE: "}.get(f, "")
            return f"{prefix}SEQUENCE ( {' THEN '.join(items)} )"
        if f in ("time-restriction", "dimension-restriction"):
            rest = {k: v for k, v in node.items() if k != "func"}
            return f"[{f}: {json.dumps(rest, ensure_ascii=False)}]"
        if f in ("exists", "not-exists", "event-exists", "not-event-exists"):
            subject = self.label(node.get("val") or node.get("evt"))
            return f"{subject} {'does not exist' if f.startswith('not') else 'exists'}"
        if f in OPERATORS:
            subject = self.label(node.get("val") or node.get("evt"))
            return f"{subject} {OPERATORS[f]} {self._value(node)}".strip()
        if f in ("attr", "event", "metric"):
            return self.label(node)

        # Fallback: unbekannte Struktur roh ausgeben, damit nichts verloren geht
        rest = {k: v for k, v in node.items() if k != "func"}
        return f"{f}{json.dumps(rest, ensure_ascii=False)}"

    # Calculated Metric ------------------------------------------------------
    ARITH = {"divide": "/", "multiply": "*", "add": "+", "subtract": "-"}

    def render_formula(self, node: Any) -> str:
        if not isinstance(node, dict):
            return str(node)
        f = node.get("func", "")
        if f == "calc-metric":
            return self.render_formula(node.get("formula"))
        if f in self.ARITH:
            return f"({self.render_formula(node.get('col1'))} {self.ARITH[f]} {self.render_formula(node.get('col2'))})"
        if f == "metric":
            txt = self.label(node)
            if node.get("filters"):
                txt += " {Filter: " + ", ".join(self._filter_label(x) for x in node["filters"]) + "}"
            return txt
        if f in ("number", "literal"):
            return str(node.get("val"))
        if f == "segment":
            return self.render(node)
        args = {k: v for k, v in node.items() if k != "func"}
        inner = ", ".join(
            self.render_formula(v) if isinstance(v, dict) else json.dumps(v, ensure_ascii=False)
            for v in args.values()
        )
        return f"{f.upper()}({inner})"

    def _filter_label(self, x: Any) -> str:
        if isinstance(x, str):
            return self._seg_ref(x)
        return self.render(x)


def comp_label(names: dict[str, str], cid: str) -> str:
    n = names.get(cid)
    return f"{n} [{cid.split('/')[-1]}]" if n else cid


def used_components(definition: Any) -> tuple[list[str], list[str]]:
    """Sammelt alle referenzierten Variablen/Metriken und Segment-IDs."""
    comps, segs = set(), set()

    def walk(n):
        if isinstance(n, dict):
            name = n.get("name")
            if isinstance(name, str) and name.startswith(("variables/", "metrics/")):
                comps.add(name)
            if "segment_id" in n:
                segs.add(n["segment_id"])
            for v in n.values():
                walk(v)
        elif isinstance(n, list):
            for v in n:
                walk(v)

    walk(definition)
    return sorted(comps), sorted(segs)


def top_context(definition: dict) -> str:
    ctx = (definition or {}).get("container", {}).get("context")
    return CONTEXT_LABEL.get(ctx, ctx or "")


# --------------------------------------------------------------------------- #
# Hilfsfunktionen DataFrames / Excel
# --------------------------------------------------------------------------- #
_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _cell(v: Any) -> Any:
    if isinstance(v, (list, dict)):
        if isinstance(v, list) and all(isinstance(x, (str, int, float)) for x in v):
            v = ", ".join(map(str, v))
        elif isinstance(v, list) and all(isinstance(x, dict) and "name" in x for x in v):
            v = ", ".join(str(x["name"]) for x in v)  # z.B. Tags
        else:
            v = json.dumps(v, ensure_ascii=False)
    if isinstance(v, str):
        v = _ILLEGAL.sub("", v)
        if len(v) > 32000:  # Excel-Zelllimit
            v = v[:32000] + " …[gekürzt]"
    return v


def to_df(rows: list[dict], first_cols: list[str] | None = None) -> pd.DataFrame:
    df = pd.DataFrame([{k: _cell(v) for k, v in r.items()} for r in rows])
    if df.empty:
        return df
    first = [c for c in (first_cols or []) if c in df.columns]
    df = df[first + [c for c in df.columns if c not in first]]
    for c in CONTEXT_COLUMNS:
        df[c] = ""
    return df


def write_excel(sheets: dict[str, pd.DataFrame], path: Path) -> None:
    from openpyxl.utils import get_column_letter

    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for name, df in sheets.items():
            if df is None or df.empty:
                df = pd.DataFrame({"Hinweis": ["Keine Daten (siehe Log)"]})
            df.to_excel(xw, sheet_name=name[:31], index=False)
            ws = xw.sheets[name[:31]]
            ws.freeze_panes = "B2"
            ws.auto_filter.ref = ws.dimensions
            for i, col in enumerate(df.columns, start=1):
                sample = [len(str(col))] + [len(str(v)) for v in df[col].head(200)]
                ws.column_dimensions[get_column_letter(i)].width = min(max(sample) + 2, 70)


# --------------------------------------------------------------------------- #
# Export-Logik
# --------------------------------------------------------------------------- #
def export(client: AdobeClient, rsids: list[str], out: Path,
           all_segments: bool, locale: str) -> None:
    client.discover_company()

    # ---- Report-Suite-Namen ------------------------------------------------
    suites = client.get20_paged("/collections/suites", {"rsids": ",".join(rsids)})
    rs_names = {s["rsid"]: s.get("name", "") for s in suites}
    missing = [r for r in rsids if r not in rs_names]
    if missing:
        log(f"WARNUNG: Report Suites nicht gefunden/kein Zugriff: {missing}")

    # ---- 2.0: Dimensionen & Metriken --------------------------------------
    dims, mets = [], []
    for rsid in rsids:
        log(f"2.0 Dimensionen/Metriken für {rsid}")
        common = {"rsid": rsid, "locale": locale, "expansion": DIM_METRIC_EXPANSION}
        dim_params = {**common, "expansion": f"{DIM_METRIC_EXPANSION},{EVAR_EXPANSION}"}
        for d in client.get20("/dimensions", dim_params, fallback_expansion=DIM_METRIC_EXPANSION):
            dims.append({"rsid": rsid, **d})
        for m in client.get20("/metrics", common):
            mets.append({"rsid": rsid, **m})

    names = {}
    for x in dims + mets:
        names.setdefault(x["id"], x.get("name", ""))

    # ---- 2.0: Segmente ----------------------------------------------------
    log("2.0 Segmente")
    seg_params = {"includeType": "all", "expansion": SEGMENT_EXPANSION, "locale": locale}
    if not all_segments:
        seg_params["rsids"] = ",".join(rsids)
    try:
        segments = client.get20_paged("/segments", seg_params)
    except RuntimeError as e:
        log(f"  includeType=all nicht erlaubt (Admin-Recht fehlt?) – nutze owned/shared: {e}")
        seg_params["includeType"] = "shared,templates"
        segments = client.get20_paged("/segments", seg_params)
    segments = list({s["id"]: s for s in segments}.values())
    seg_names = {s["id"]: s.get("name", "") for s in segments}

    # ---- 2.0: Calculated Metrics ------------------------------------------
    log("2.0 Berechnete Metriken")
    cm_params = {"includeType": "all", "expansion": CALCMETRIC_EXPANSION, "locale": locale}
    if not all_segments:
        cm_params["rsids"] = ",".join(rsids)
    try:
        calcs = client.get20_paged("/calculatedmetrics", cm_params)
    except RuntimeError as e:
        log(f"  includeType=all nicht erlaubt – nutze owned/shared: {e}")
        cm_params["includeType"] = "shared,templates"
        calcs = client.get20_paged("/calculatedmetrics", cm_params)
    calcs = list({c["id"]: c for c in calcs}.values())
    for c in calcs:
        names.setdefault(f"metrics/{c['id']}", c.get("name", ""))

    renderer = DefinitionRenderer(names, seg_names)

    seg_rows = []
    for s in segments:
        definition = s.get("definition") or {}
        comps, refs = used_components(definition)
        try:
            readable = renderer.render(definition)
        except Exception as e:  # noqa: BLE001
            readable = f"(nicht darstellbar: {e})"
        seg_rows.append({
            "id": s.get("id"),
            "name": s.get("name"),
            "description": s.get("description"),
            "scope_container": top_context(definition),
            "definition_lesbar": readable,
            "verwendete_variablen": ", ".join(comp_label(names, c) for c in comps),
            "verwendete_segmente": ", ".join(seg_names.get(r, r) for r in refs),
            "rsid": s.get("rsid"),
            "report_suite": s.get("reportSuiteName"),
            "owner": (s.get("owner") or {}).get("name") or (s.get("owner") or {}).get("login"),
            "tags": s.get("tags"),
            "shares": s.get("sharesFullName") or s.get("shares"),
            "modified": s.get("modified"),
            "compatibility": s.get("compatibility"),
            "definition_json": definition,
        })

    calc_rows = []
    for c in calcs:
        definition = c.get("definition") or {}
        comps, refs = used_components(definition)
        try:
            readable = renderer.render_formula(definition)
        except Exception as e:  # noqa: BLE001
            readable = f"(nicht darstellbar: {e})"
        calc_rows.append({
            "id": c.get("id"),
            "name": c.get("name"),
            "description": c.get("description"),
            "formel_lesbar": readable,
            "type": c.get("type"),
            "polarity": c.get("polarity"),
            "precision": c.get("precision"),
            "verwendete_komponenten": ", ".join(comp_label(names, x) for x in comps),
            "verwendete_segmente": ", ".join(seg_names.get(r, r) for r in refs),
            "rsid": c.get("rsid"),
            "report_suite": c.get("reportSuiteName"),
            "owner": (c.get("owner") or {}).get("name") or (c.get("owner") or {}).get("login"),
            "tags": c.get("tags"),
            "shares": c.get("sharesFullName") or c.get("shares"),
            "modified": c.get("modified"),
            "definition_json": definition,
        })

    # ---- eVars, Props, Listvars, Events aus der 2.0 API --------------------
    def from_20(rows: list[dict], pattern: str) -> list[dict]:
        return [{"report_suite": rs_names.get(r["rsid"], ""), **r}
                for r in rows if re.fullmatch(pattern, r["id"])]

    evars = from_20(dims, r"variables/evar\d+")
    props = from_20(dims, r"variables/prop\d+")
    listvars = from_20(dims, r"variables/listvariable\d+")
    events = from_20(mets, r"metrics/event\d+")

    for e in evars:  # Binding Events lesbar machen (Metrik-IDs -> Namen)
        if isinstance(e.get("bindingEvents"), list):
            e["bindingEvents"] = ", ".join(comp_label(names, b) for b in e["bindingEvents"])
    evar_settings = sum(1 for e in evars if e.get("allocationType") or e.get("expirationType"))
    if evars and not evar_settings:
        log("  WARNUNG: Die 2.0 API hat keine eVar-Einstellungen (Allokation/Expiration) geliefert.")
    else:
        log(f"  eVar-Einstellungen für {evar_settings} von {len(evars)} eVars erhalten")

    def sort_key(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty or "id" not in df.columns:
            return df
        num = df["id"].astype(str).str.extract(r"(\d+)$")[0].astype(float)
        return df.assign(_n=num).sort_values(["rsid", "_n"]).drop(columns="_n")

    overview = pd.DataFrame([{
        "rsid": r,
        "report_suite": rs_names.get(r, "(nicht gefunden)"),
        "eVars": sum(1 for e in evars if e["rsid"] == r),
        "eVars mit Einstellungen": sum(1 for e in evars if e["rsid"] == r
                                       and (e.get("allocationType") or e.get("expirationType"))),
        "Props": sum(1 for p in props if p["rsid"] == r),
        "Listvars": sum(1 for x in listvars if x["rsid"] == r),
        "Events": sum(1 for e in events if e["rsid"] == r),
        "Dimensionen gesamt": sum(1 for d in dims if d["rsid"] == r),
        "Metriken gesamt": sum(1 for m in mets if m["rsid"] == r),
        "Exportiert am": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "Quelle": "Adobe Analytics 2.0 API",
    } for r in rsids])

    head = ["rsid", "report_suite", "id", "name", "description"]
    sheets = {
        "Übersicht": overview,
        "eVars": sort_key(to_df(evars, head + EVAR_FIELDS)),
        "Props": sort_key(to_df(props, head)),
        "Listvars": sort_key(to_df(listvars, head)),
        "Events": sort_key(to_df(events, head)),
        "Segmente": to_df(seg_rows, ["id", "name"]),
        "Calc Metrics": to_df(calc_rows, ["id", "name"]),
        "Dimensionen (alle)": to_df(dims, ["rsid", "id", "name", "description"]),
        "Metriken (alle)": to_df(mets, ["rsid", "id", "name", "description"]),
    }
    write_excel(sheets, out)
    log(f"Fertig: {out.resolve()}")
    for n, df in sheets.items():
        log(f"   {n:<20} {len(df):>6} Zeilen")


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def main() -> None:
    ap = argparse.ArgumentParser(description="Adobe Analytics SDR-Export (Events, Variablen, Segmente)")
    ap.add_argument("--config", default="config.json", help="Pfad zur config.json")
    ap.add_argument("--rsids", nargs="*", help="Report Suite IDs (überschreibt config.json)")
    ap.add_argument("--out", help="Ausgabedatei (.xlsx)")
    ap.add_argument("--list-rsids", action="store_true", help="Nur verfügbare Report Suites auflisten")
    # Ohne Funktion: die 1.4 API ist abgeschaltet. Bleibt erhalten, damit alte Aufrufe weiter laufen.
    ap.add_argument("--skip-14", action="store_true", help=argparse.SUPPRESS)
    ap.add_argument("--all-segments", action="store_true",
                    help="Alle Segmente/Calc Metrics der Company, nicht nur die der gewählten RS")
    ap.add_argument("--locale", default=None, help="z.B. de_DE oder en_US (Standard: config bzw. en_US)")
    args = ap.parse_args()

    cfg_path = Path(args.config)
    cfg = json.loads(cfg_path.read_text(encoding="utf-8")) if cfg_path.exists() else {}
    # Umgebungsvariablen haben Vorrang vor config.json (Secrets nicht in Dateien ablegen)
    import os
    for key, env in (("client_id", "ADOBE_CLIENT_ID"), ("client_secret", "ADOBE_CLIENT_SECRET"),
                     ("global_company_id", "ADOBE_GLOBAL_COMPANY_ID"),
                     ("ims_org_id", "ADOBE_IMS_ORG_ID")):
        if os.environ.get(env):
            cfg[key] = os.environ[env]
    if not cfg.get("client_id") or not cfg.get("client_secret"):
        raise SystemExit("client_id/client_secret fehlen (config.json oder ADOBE_CLIENT_ID/ADOBE_CLIENT_SECRET).")
    client = AdobeClient(cfg)

    if args.list_rsids:
        client.discover_company()
        for s in client.get20_paged("/collections/suites", {}):
            print(f"{s['rsid']}\t{s.get('name', '')}")
        return

    rsids = args.rsids or cfg.get("rsids") or []
    if not rsids:
        raise SystemExit("Keine Report Suites angegeben (--rsids oder 'rsids' in config.json).")
    out = Path(args.out or f"adobe_sdr_{datetime.now():%Y%m%d_%H%M}.xlsx")
    export(client, rsids, out, all_segments=args.all_segments,
           locale=args.locale or cfg.get("locale", "en_US"))


if __name__ == "__main__":
    main()
