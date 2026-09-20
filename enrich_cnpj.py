"""
CNPJ Data Enrichment

Enriches a spreadsheet of Brazilian company registrations (CNPJ) with public
records from the Receita Federal, queried through several independent public
APIs with automatic failover.

Built for vendor due diligence: it fills in the registered name, trade name,
registration status, primary economic activity (CNAE), business segment,
company size, incorporation date, location and public contact details, and
raises a risk flag when a registration is not active or the company is under
judicial recovery.

Each provider exposes the same registry data in a different JSON shape. A
per-provider adapter normalizes every response to one internal record, so
adding a provider means writing one function, not touching the main loop.

The process is resumable. Only rows marked PENDING are processed and the
workbook is checkpointed periodically, so an interrupted run resumes cleanly.

Usage:
    pip install -r requirements.txt
    python enrich_cnpj.py vendors.xlsx --sheet VENDORS

Author: Daniel Batista
License: MIT
"""

import argparse
import json
import re
import sys
import time
import unicodedata
from datetime import date
from pathlib import Path

import requests
from openpyxl import load_workbook

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

REPO_URL = "https://github.com/yl-dan/cnpj-enricher"
USER_AGENT = f"cnpj-enricher/2.0 (+{REPO_URL})"

REQUEST_TIMEOUT = 25        # seconds
THROTTLE_SECONDS = 1.5      # delay between rows
COOLDOWN_SECONDS = 120      # how long a rate-limited provider is skipped
SAVE_EVERY = 25             # rows between checkpoint saves

STATUS_PENDING = "PENDING"
STATUS_DONE = "ENRICHED"
STATUS_NOT_FOUND = "NOT FOUND"
STATUS_FAILED = "LOOKUP FAILED"
STATUS_INVALID = "INVALID CNPJ"

REQUIRED_COLUMNS = [
    "cnpj", "status_enrichment", "registered_name", "trade_name",
    "registration_status", "primary_cnae", "segment", "incorporation_date",
    "company_size", "share_capital", "city", "state", "public_email",
    "public_phone", "risk_flag", "source", "query_date",
]

# Optional: used only to label progress output.
VENDOR_NAME_COLUMN = "vendor_name"

DEFAULT_SEGMENTS = {
    "10.": "Food manufacturing",
    "17.": "Paper and packaging manufacturing",
    "20.": "Chemical manufacturing",
    "46.": "Wholesale trade",
    "47.": "Retail trade",
    "49.": "Transport",
    "52.": "Logistics and warehousing",
    "62.": "Information technology",
    "69.": "Legal and accounting services",
}
SEGMENT_FALLBACK = "Other"


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------

def strip_accents(text):
    """
    Removes diacritics so comparisons are accent-insensitive.

    Each provider normalizes registry text differently: the same status can
    arrive as RECUPERACAO JUDICIAL from one and RECUPERAÇÃO JUDICIAL from
    another. Without this, whether a risk flag fires would depend on which
    provider happened to answer first.
    """
    normalized = unicodedata.normalize("NFKD", str(text or ""))
    return "".join(c for c in normalized if not unicodedata.combining(c))


def digits_only(value):
    return re.sub(r"\D", "", str(value or ""))


def format_cnae(code, description):
    """Formats a CNAE code as NN.NN-N-NN, appending its description."""
    digits = digits_only(code).zfill(7)
    if not digits.strip("0"):
        return ""
    formatted = f"{digits[:2]}.{digits[2:4]}-{digits[4]}-{digits[5:]}"
    return f"{formatted} - {description or ''}".strip(" -")


def format_date(value):
    """Accepts YYYY-MM-DD or DD/MM/YYYY and returns DD/MM/YYYY."""
    text = str(value or "")[:10]
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", text):
        year, month, day = text.split("-")
        return f"{day}/{month}/{year}"
    return text if re.fullmatch(r"\d{2}/\d{2}/\d{4}", text) else ""


def format_phone(*parts):
    """Joins area code and number fragments into (NN) NNNNN-NNNN."""
    digits = digits_only("".join(str(p or "") for p in parts))
    if len(digits) < 10:
        return ""
    return f"({digits[:2]}) {digits[2:-4]}-{digits[-4:]}"


def classify_segment(cnae, segments):
    if not cnae:
        return ""
    for prefix in sorted(segments, key=len, reverse=True):
        if cnae.startswith(prefix):
            return segments[prefix]
    return SEGMENT_FALLBACK


def load_segments(path):
    if not path:
        return DEFAULT_SEGMENTS
    try:
        with open(path, encoding="utf-8") as handle:
            mapping = json.load(handle)
    except (OSError, json.JSONDecodeError) as error:
        sys.exit(f"Could not read segment map '{path}': {error}")
    if not isinstance(mapping, dict):
        sys.exit(f"Segment map '{path}' must be a JSON object of prefix -> label.")
    return mapping


# ---------------------------------------------------------------------------
# Provider adapters
#
# Each adapter receives one provider's raw JSON and returns the same internal
# record. Adding a provider means adding one function and one Provider entry.
# ---------------------------------------------------------------------------

def parse_brasilapi(payload):
    """BrasilAPI and Minha Receita expose the same response shape."""
    return {
        "registered_name": payload.get("razao_social"),
        "trade_name": payload.get("nome_fantasia"),
        "registration_status": payload.get("descricao_situacao_cadastral"),
        "cnae": format_cnae(payload.get("cnae_fiscal"), payload.get("cnae_fiscal_descricao")),
        "incorporation_date": format_date(payload.get("data_inicio_atividade")),
        "company_size": payload.get("porte"),
        "share_capital": payload.get("capital_social"),
        "city": payload.get("municipio"),
        "state": payload.get("uf"),
        "email": payload.get("email"),
        "phone": format_phone(payload.get("ddd_telefone_1")),
    }


def parse_cnpjws(payload):
    establishment = payload.get("estabelecimento") or {}
    activity = establishment.get("atividade_principal") or {}
    return {
        "registered_name": payload.get("razao_social"),
        "trade_name": establishment.get("nome_fantasia"),
        "registration_status": establishment.get("situacao_cadastral"),
        "cnae": format_cnae(activity.get("subclasse"), activity.get("descricao")),
        "incorporation_date": format_date(establishment.get("data_inicio_atividade")),
        "company_size": (payload.get("porte") or {}).get("descricao"),
        "share_capital": payload.get("capital_social"),
        "city": (establishment.get("cidade") or {}).get("nome"),
        "state": (establishment.get("estado") or {}).get("sigla"),
        "email": establishment.get("email"),
        "phone": format_phone(establishment.get("ddd1"), establishment.get("telefone1")),
    }


def parse_cnpja(payload):
    company = payload.get("company") or {}
    activity = payload.get("mainActivity") or {}
    address = payload.get("address") or {}
    emails = payload.get("emails") or []
    phones = payload.get("phones") or []
    return {
        "registered_name": company.get("name"),
        "trade_name": payload.get("alias"),
        "registration_status": (payload.get("status") or {}).get("text"),
        "cnae": format_cnae(activity.get("id"), activity.get("text")),
        "incorporation_date": format_date(payload.get("founded")),
        "company_size": (company.get("size") or {}).get("text"),
        "share_capital": company.get("equity"),
        "city": address.get("city"),
        "state": address.get("state"),
        "email": emails[0].get("address") if emails else "",
        "phone": format_phone(phones[0].get("area"), phones[0].get("number")) if phones else "",
    }


def parse_receitaws(payload):
    activities = payload.get("atividade_principal") or [{}]
    activity = activities[0] if activities else {}
    return {
        "registered_name": payload.get("nome"),
        "trade_name": payload.get("fantasia"),
        "registration_status": payload.get("situacao"),
        "cnae": format_cnae(activity.get("code"), activity.get("text")),
        "incorporation_date": format_date(payload.get("abertura")),
        "company_size": payload.get("porte"),
        "share_capital": payload.get("capital_social"),
        "city": payload.get("municipio"),
        "state": payload.get("uf"),
        "email": payload.get("email"),
        "phone": format_phone(payload.get("telefone")),
    }


class Provider:
    """
    One upstream API, with its own availability state.

    Rate limiting and refusal are handled differently on purpose. HTTP 429
    means "slow down", so the provider is put on a cooldown and reconsidered
    later. HTTP 403 means "no", so it is dropped for the rest of the run.
    Treating both as permanent would discard the best provider on its first
    busy minute.
    """

    def __init__(self, name, url_template, parser):
        self.name = name
        self.url_template = url_template
        self.parser = parser
        self.disabled = False
        self.cooldown_until = 0.0

    def available(self):
        return not self.disabled and time.monotonic() >= self.cooldown_until

    def throttle(self):
        self.cooldown_until = time.monotonic() + COOLDOWN_SECONDS

    def refuse(self):
        self.disabled = True

    def url(self, cnpj):
        return self.url_template.format(cnpj)


PROVIDERS = [
    Provider("BrasilAPI", "https://brasilapi.com.br/api/cnpj/v1/{}", parse_brasilapi),
    Provider("Minha Receita", "https://minhareceita.org/{}", parse_brasilapi),
    Provider("CNPJ.ws", "https://publica.cnpj.ws/cnpj/{}", parse_cnpjws),
    Provider("CNPJA", "https://open.cnpja.com/office/{}", parse_cnpja),
    Provider("ReceitaWS", "https://receitaws.com.br/v1/cnpj/{}", parse_receitaws),
]


# ---------------------------------------------------------------------------
# Lookup
# ---------------------------------------------------------------------------

def lookup(session, cnpj):
    """
    Queries providers in order until one returns a usable record.

    Returns (outcome, record, provider_name) where outcome is one of:
        "found"     a provider returned the company
        "not_found" every provider that answered said the CNPJ does not exist
        "failed"    no provider gave a definitive answer this run

    The distinction between not_found and failed matters: marking a network
    failure as NOT FOUND would quietly assert that a real company does not
    exist, and that row would never be revisited.
    """
    saw_definitive_miss = False

    for provider in PROVIDERS:
        if not provider.available():
            continue

        try:
            response = session.get(provider.url(cnpj), timeout=REQUEST_TIMEOUT)
        except requests.RequestException:
            continue

        if response.status_code == 200:
            try:
                record = provider.parser(response.json())
            except (ValueError, KeyError, TypeError, IndexError, AttributeError):
                # Malformed or unexpected payload from this provider only.
                continue
            if record.get("registered_name"):
                return "found", record, provider.name
            continue

        if response.status_code in (400, 404):
            saw_definitive_miss = True
            continue

        if response.status_code == 429:
            provider.throttle()
            print(f"    {provider.name} is rate limiting; pausing it for {COOLDOWN_SECONDS}s")
            continue

        if response.status_code == 403:
            provider.refuse()
            print(f"    {provider.name} refused access; dropping it for this run")
            continue

        # 5xx and anything else: transient, try the next provider.

    return ("not_found" if saw_definitive_miss else "failed"), None, ""


def build_updates(record, segments, provider_name, today):
    """Maps an internal record onto the sheet's column names."""
    status = (record.get("registration_status") or "").upper()
    registered_name = record.get("registered_name") or ""

    flags = []
    if status and strip_accents(status) != "ATIVA":
        flags.append(f"REGISTRATION {status}")
    if "RECUPERACAO JUDICIAL" in strip_accents(registered_name).upper():
        flags.append("JUDICIAL RECOVERY")

    return {
        "registered_name": registered_name,
        "trade_name": record.get("trade_name"),
        "registration_status": status,
        "primary_cnae": record.get("cnae"),
        "segment": classify_segment(record.get("cnae"), segments),
        "incorporation_date": record.get("incorporation_date"),
        "company_size": (record.get("company_size") or "").upper(),
        "share_capital": record.get("share_capital"),
        "city": (record.get("city") or "").upper(),
        "state": record.get("state"),
        "public_email": (record.get("email") or "").lower(),
        "public_phone": record.get("phone"),
        "risk_flag": "; ".join(flags),
        "status_enrichment": STATUS_DONE,
        "source": f"{provider_name} (Receita Federal)",
        "query_date": today,
    }


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def parse_args():
    parser = argparse.ArgumentParser(
        description="Enrich a CNPJ spreadsheet with public Receita Federal data.",
    )
    parser.add_argument("workbook", help="path to the .xlsx file to enrich")
    parser.add_argument("--sheet", default="VENDORS", help="worksheet name (default: VENDORS)")
    parser.add_argument("--segments", help="path to a JSON map of CNAE prefix to segment label")
    parser.add_argument("--limit", type=int, help="stop after this many rows")
    return parser.parse_args()


def main():
    args = parse_args()

    path = Path(args.workbook)
    if not path.is_file():
        sys.exit(f"File not found: {path}")

    segments = load_segments(args.segments)

    workbook = load_workbook(path)
    if args.sheet not in workbook.sheetnames:
        sys.exit(f"Sheet '{args.sheet}' not found. Available: {', '.join(workbook.sheetnames)}")
    sheet = workbook[args.sheet]

    columns = {cell.value: cell.column for cell in sheet[1] if cell.value}
    missing = [name for name in REQUIRED_COLUMNS if name not in columns]
    if missing:
        sys.exit("Missing required column(s): " + ", ".join(missing))

    label_column = columns.get(VENDOR_NAME_COLUMN)

    rows = [
        row for row in range(2, sheet.max_row + 1)
        if sheet.cell(row, columns["status_enrichment"]).value == STATUS_PENDING
    ]
    if args.limit:
        rows = rows[:args.limit]

    if not rows:
        print("Nothing to do: no rows marked PENDING.")
        return

    print(f"{len(rows)} row(s) pending. {len(PROVIDERS)} providers configured.")

    today = date.today().isoformat()
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT})

    counts = {"found": 0, "not_found": 0, "failed": 0, "invalid": 0}

    def write(row, column_name, value):
        if value not in (None, ""):
            sheet.cell(row, columns[column_name]).value = value

    for index, row in enumerate(rows, start=1):
        if not any(p.available() for p in PROVIDERS):
            print("Every provider is unavailable. Stopping; rerun later to resume.")
            break

        label = str(sheet.cell(row, label_column).value)[:35] if label_column else f"row {row}"
        cnpj = digits_only(sheet.cell(row, columns["cnpj"]).value)

        if len(cnpj) != 14:
            sheet.cell(row, columns["status_enrichment"]).value = STATUS_INVALID
            sheet.cell(row, columns["query_date"]).value = today
            counts["invalid"] += 1
            print(f"[{index}/{len(rows)}] {label}: invalid CNPJ")
            continue

        outcome, record, provider_name = lookup(session, cnpj)

        if outcome == "found":
            for name, value in build_updates(record, segments, provider_name, today).items():
                write(row, name, value)
            sheet.cell(row, columns["status_enrichment"]).value = STATUS_DONE
            counts["found"] += 1
            status = record.get("registration_status") or "ok"
            print(f"[{index}/{len(rows)}] {label}: {status} [{provider_name}]")
        elif outcome == "not_found":
            sheet.cell(row, columns["status_enrichment"]).value = STATUS_NOT_FOUND
            counts["not_found"] += 1
            print(f"[{index}/{len(rows)}] {label}: not found in any registry")
        else:
            # Left as LOOKUP FAILED, never NOT FOUND: no provider actually
            # said this company does not exist. Reset to PENDING to retry.
            sheet.cell(row, columns["status_enrichment"]).value = STATUS_FAILED
            counts["failed"] += 1
            print(f"[{index}/{len(rows)}] {label}: lookup failed, retry later")

        sheet.cell(row, columns["query_date"]).value = today

        if index % SAVE_EVERY == 0:
            workbook.save(path)
            print(f"    checkpoint saved ({counts['found']} enriched)")

        time.sleep(THROTTLE_SECONDS)

    workbook.save(path)
    print(
        f"Done. enriched={counts['found']} not_found={counts['not_found']} "
        f"failed={counts['failed']} invalid={counts['invalid']}"
    )
    if counts["failed"]:
        print(f"Reset the {counts['failed']} '{STATUS_FAILED}' row(s) to '{STATUS_PENDING}' to retry.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit("\nInterrupted. Progress up to the last checkpoint was saved.")
