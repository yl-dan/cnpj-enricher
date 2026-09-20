# CNPJ Data Enrichment

Enriches a spreadsheet of Brazilian company registrations (CNPJ) with public
records from the Receita Federal, queried through five independent public APIs
with automatic failover.

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

## Providers

| Provider | Endpoint |
|---|---|
| BrasilAPI | `brasilapi.com.br` |
| Minha Receita | `minhareceita.org` |
| CNPJ.ws | `publica.cnpj.ws` |
| CNPJA | `open.cnpja.com` |
| ReceitaWS | `receitaws.com.br` |

All five expose the same registry data in different JSON shapes. A per-provider
adapter normalizes every response into one internal record, so adding a source
means writing one function and one table entry, never touching the main loop.
Providers are tried in order until one returns a usable record.

## Design notes

**Resumable by default.** Only rows marked `PENDING` are processed, and the
workbook is checkpointed every 25 rows. An interrupted run restarts without
re-requesting what it already has.

**Rate limiting and refusal are handled differently.** HTTP 429 means "slow
down", so that provider is put on a two-minute cooldown and reconsidered later.
HTTP 403 means "no", so it is dropped for the rest of the run. Treating both as
permanent would discard the best provider on its first busy minute.

**"Not found" is never guessed.** A row is marked `NOT FOUND` only when a
provider actually answered that the CNPJ does not exist. If every provider was
unreachable or throttled, the row is marked `LOOKUP FAILED` instead. Writing
`NOT FOUND` after a network failure would quietly assert that a real company
does not exist, and that row would never be revisited.

**The client identifies itself.** Requests carry a User-Agent naming this tool
and linking to its repository, not a spoofed browser string. These are free
public services; if one declines an honest client, the answer is to respect the
limit, not to disguise the request.

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

An optional `vendor_name` column, if present, is used to label progress output.

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
| `NOT FOUND` | A provider confirmed no such CNPJ exists |
| `INVALID CNPJ` | The cell did not contain 14 digits. No request was made |
| `LOOKUP FAILED` | No provider gave a definitive answer. Reset to `PENDING` to retry |

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

Python, `requests`, `openpyxl`. Public registry APIs listed above.

## License

MIT: see [LICENSE](LICENSE).
