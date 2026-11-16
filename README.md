# RosterRecon

RosterRecon is an OSINT CLI that collects employee names, job titles, and
locations from a company search on LinkedIn, generates common corporate
username patterns, and exports the results to XLSX, CSV, and text files.

It is a distinct, expanded project derived from
[initstring/linkedin2username](https://github.com/initstring/linkedin2username)
under the MIT License.

## What it exports

Each run creates:

- `<company>-employees.xlsx` — formatted workbook with employee metadata,
  generated username patterns, filters, frozen headers, and a location summary.
- `<company>-employees.csv` — standards-compliant UTF-8 CSV containing
  `full_name`, `job_title`, and `location`.
- `<company>-rawnames.txt` — one full name per line.
- `<company>-first.last.txt`, `<company>-f.last.txt`,
  `<company>-flast.txt`, `<company>-firstl.txt`, `<company>-first.txt`, and
  `<company>-lastf.txt` — one username candidate per line.

Repeated search hits are deduplicated before export. Spreadsheet-bound profile
text is escaped to prevent formula injection.

## Requirements

- Python 3.10 or newer
- Firefox, Chrome, or Chromium in a location Selenium can discover
- A valid LinkedIn account (I highly recommend building sock puppets)

Install the project in a standard virtual environment:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --editable .
```

## Usage

Run from the repository:

```bash
rosterrecon --company targetco
```

Append an email domain and limit collection to five 50-result pages:

```bash
rosterrecon --company targetco --domain targetco.com --depth 5
```

Wait at least 750 milliseconds between result-page requests:

```bash
rosterrecon --company targetco --rate-limit 750
```

`--rate-limit-ms` is an equivalent, more explicit spelling. The default is
zero milliseconds; use a conservative non-zero delay to reduce request pace.

Choose an output directory:

```bash
rosterrecon --company targetco --output results/targetco
```

XLSX export is enabled by default. Use `--no-xlsx` when only CSV and text
outputs are needed:

```bash
rosterrecon --company targetco --no-xlsx
```

During a run, Selenium opens LinkedIn in a browser. Complete login (including
MFA if required), leave the browser open, and press Enter in the terminal.
RosterRecon then reuses that authenticated browser session for company search.

For every option:

```bash
rosterrecon --help
```

## Responsible use

LinkedIn may rate-limit searches or restrict an account. Use a dedicated,
authorized account, select a conservative depth and delay, and collect only
what is necessary for an approved assessment. Do not use RosterRecon for
harassment, unsolicited outreach, credential attacks, or collection outside
your authorization.

LinkedIn is a trademark of LinkedIn Corporation. RosterRecon is not affiliated
with or endorsed by LinkedIn.

## Development

Run the tests without contacting LinkedIn:

```bash
python -m pip install --editable '.[dev]'
python -m pytest
```

The fixture-based tests cover name mutation, metadata extraction,
deduplication, CSV quoting, XLSX formatting, and spreadsheet formula safety.

## License and attribution

MIT. The original copyright notice for linkedin2username is retained in
[LICENSE](LICENSE), as required by the upstream license.
