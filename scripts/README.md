# populate_csr_store_name.py

Populates the hidden `CSR Store Name` column on the two live Smartsheets:

- `NVS List (NEW)` — sheet id `8700079041892228`, column id `6363478419083140`
- `NSO List (NEW)` — sheet id `328848479571844`, column id `6957589299761028`

The hidden TEXT_NUMBER `CSR Store Name` column is pre-created on both sheets.
The script only writes into that column; it touches nothing else.

## Source data

`data/Store_Names_Updated.xlsx` (committed alongside the script) has two sheets,
`NVS` and `NSO`, each with `Store #` / `Store Name` columns. The two sheets are
merged into a single mapping keyed by `Store #` (with `.0` and leading zeros
stripped).

## Auth

The script reads `SMARTSHEET_ACCESS_TOKEN` if set. Otherwise it pulls from AWS
Secrets Manager:

- `SMARTSHEET_AWS_SECRET_ID` (default `smartsheet/access_token`)
- `AWS_REGION` (default `us-west-2`)

In this org the standing secret is `sync-monitors/prod/smartsheet-api-token`.

## Usage

```powershell
# preferred — dry-run first
$env:SMARTSHEET_AWS_SECRET_ID = "sync-monitors/prod/smartsheet-api-token"
python scripts/populate_csr_store_name.py --dry-run

# live write once counts look sane
python scripts/populate_csr_store_name.py
```

## Behaviour

- Rows whose `Store #` matches the mapping get a `PUT /sheets/{id}/rows`
  setting `CSR Store Name` to the mapped value, in chunks of 400.
- Rows whose existing `CSR Store Name` already equals the mapped value are
  skipped (idempotent).
- Blank / `TBD` / `N/A` `Store #` rows are left alone.
- Unmatched `Store #` rows are reported but left alone.
- 429 / 5xx responses retry with exponential backoff.
