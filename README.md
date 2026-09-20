# CNPJ Data Enrichment

Enriches a spreadsheet of Brazilian company registrations (CNPJ) with public
records from the Receita Federal, via the BrasilAPI public endpoint.

Built for vendor due diligence. Given a list of CNPJs, it fills in the
registered name, trade name, registration status, primary economic activity
(CNAE), business segment, company size, incorporation date, location and
public contact details, and raises a risk flag when a registration is not
active or the company is under judicial recovery.

## Why this exists

Verifying a vendor base by hand means opening one registry page per company
and copying fields into a spreadsheet. At a few hundred vendors that stops
being viable, and manual transcription introduces exactly the kind of error
a compliance check is supposed to catch.

## Design notes

**Resumable by default.** Only rows marked `PENDING` are processed, and the
workbook is checkpointed every 25 rows. An interrupted run restarts without
re-requesting what it already has.

**Retries hit the same row, not the next one.** On HTTP 429 or a 5xx, the
script backs off and retries that CNPJ up to three times. Skipping ahead on
throttling would scatter silent gaps through the output, which is worse than
failing: the sheet would look complete.

**Header validation happens upfront.** A missing column is reported before
the first request, not as a `KeyError` forty minutes into a run.

**Accent-insensitive risk matching.** Registry data is inconsistent about
diacritics, and a raw string comparison silently misses
`RECUPERAÇÃO JUDICIAL` while matching `RECUPERACAO JUDICIAL`. Both forms are
normalized before comparison. This is the flag that matters most, so it is
the one least acceptable to miss.

**Failures are recorded, not hidden.** Rows that exhaust their retries are
marked `REQUEST FAILED` rather than left `PENDING`, so a run is auditable.
Set them back to `PENDING` to retry later.

**Rate limiting is respected.** Requests are spaced by 1.2 seconds against a
free public API. Running faster is a good way to lose access for everyone.

## Data protection note

CNPJ records are public and concern legal entities, so they fall outside the
scope of the Brazilian LGPD in most cases. Two caveats are worth carrying
into any deployment:

- For sole proprietorships (MEI and similar), the registered name is often
  the owner's own name, and the contact details are personal. Those records
  do involve personal data.
- Enriched output should inherit the access controls of the source vendor
  base, not be treated as public because its inputs were.

The `.gitignore` excludes `*.xlsx` and `*.csv` so source data is never
committed alongside the code.

## Setup

```bash
pip install -r requirements.txt
```

## Usage

```bash
python enrich_cnpj.py vendors.xlsx --sheet VENDORS
```

| Option | Description |
|---|---|
| `workbook` | Path to the `.xlsx` file to enrich (required) |
| `--sheet` | Worksheet name. Default: `VENDORS` |
| `--segments` | Path to a JSON map of CNAE prefix to segment label |
| `--limit` | Stop after N rows. Useful for a first test run |

Start with `--limit 5` against a copy of the file to confirm the column
mapping before committing to a full run.

## Expected columns

The first row must contain these headers, in any order:

```
cnpj                   status_enrichment      registered_name
trade_name             registration_status    primary_cnae
segment                incorporation_date     company_size
share_capital          city                   state
public_email           public_phone           risk_flag
source                 query_date
```

Rows are processed only when `status_enrichment` is `PENDING`.

### Status values

| Value | Meaning |
|---|---|
| `PENDING` | Not yet processed. The script only touches these |
| `ENRICHED` | Record found and written |
| `NOT FOUND` | The registry has no such CNPJ |
| `INVALID CNPJ` | The cell did not contain 14 digits. No request was made |
| `REQUEST FAILED` | Retries exhausted. Reset to `PENDING` to try again |

## Segment mapping

Business segments are derived from the CNAE code by longest-prefix match.
A built-in default covers common categories; supply your own to match how
your organization classifies vendors:

```bash
python enrich_cnpj.py vendors.xlsx --segments segments.json
```

```json
{
  "46.3": "Food distribution",
  "49.": "Transport",
  "62.": "Information technology"
}
```

See `segments.example.json` for a fuller example. Longer prefixes win, so
`46.3` takes precedence over `46.`.

## Tech

Python, `requests`, `openpyxl`, BrasilAPI.

## License

MIT: see [LICENSE](LICENSE).
