# Chinese-Language Scraper

A single-file Python scraper for building a complete, searchable archive of a Chinese A-share company's public footprint — **official exchange filings**, **corporate websites**, and an extracted **text + keyword index** — in one folder.

It handles the awkward parts of Chinese-market scraping for you: the cninfo / Shenzhen Stock Exchange disclosure APIs, Chinese-language PDFs and Office files, GB/UTF-8 encoding, bot-challenge pages, and resumable long crawls.

> Ships configured for a worked example — Jereh Group (SZSE: `002353`) — so it runs out of the box. See [Retargeting](#retargeting) to point it at a different company.

## What it collects

1. **Exchange filings** — every announcement since listing plus investor-relations records, from [cninfo.com.cn](https://www.cninfo.com.cn) (China's official disclosure site). If cninfo fails for a given year, that year falls back to [szse.cn](https://www.szse.cn) (Shenzhen Stock Exchange).
2. **Websites** — every page (saved as text) and every document (PDF, Word, Excel, PowerPoint, ZIP) on the configured corporate sites, in Chinese and English.
3. **Text & keywords** — extracts the text of every PDF and Office file and flags documents that mention a configurable keyword list (works across Chinese and English, ignoring spaces, line breaks and case).

## Requirements

- Python 3.7+
- `requests`, `beautifulsoup4` (required)
- `pypdf` (optional — PDF text extraction is skipped without it)
- `cryptography` (optional — needed to read AES-encrypted PDFs)

```bash
pip install requests beautifulsoup4 pypdf
pip install --only-binary :all: cryptography   # optional, for encrypted PDFs
```

> On macOS, installing `cryptography` with `--only-binary :all:` avoids a from-source Rust/OpenSSL build.

## Usage

```bash
python jereh_scraper.py --check       # ~1-minute test of every source
python jereh_scraper.py               # scrape everything
python jereh_scraper.py --scan-only   # rebuild the text + reports only
```

### Options

| Option | Description |
| --- | --- |
| `--source filings\|website\|all` | What to scrape (default: `all`) |
| `--since 2025-01-01` | Only filings from this date |
| `--until 2025-12-31` | Only filings up to this date |
| `--main-site-only` | Skip the secondary / group sites |
| `--max-pages 500` | Stop the website crawl after N pages |
| `--delay 1.0` | Seconds between requests to the same site |
| `--keywords "Kawasaki,LM2500"` | Extra keywords to flag |
| `--out ./docs` | Output folder |
| `--filings-from auto\|cninfo\|szse\|both` | Filing source strategy |
| `--save-html` | Also keep the raw HTML of every page |
| `--fresh` | Restart the website crawl from scratch |

A full run can take several hours. It is **resumable** — stop with `Ctrl+C` and run the same command again; finished downloads are skipped, failed ones retried, and the website crawl picks up where it stopped.

## Output

```
<out>/
  filings/cninfo/<year>/<date>_<id>_<title>.pdf   # filings (+ .txt with extracted text)
  filings/szse/<year>/...                         # only if szse.cn was used
  website/<site>/pages/*.txt                      # text of every web page
  website/<site>/files/*                          # documents from the sites
  index.csv            # every item: source, date, title, link, file, keywords
  keyword_hits.csv     # one row per document per keyword, with snippets
  scrape_log.txt       # full log of every run
  manifest.jsonl       # record of everything downloaded (used to resume)
```

## Retargeting

The scraper is driven by a few constants near the top of the script — edit these to point it at a different company:

| Constant | What to change |
| --- | --- |
| `STOCK_CODE` | The 6-digit A-share code (e.g. `"002353"`) |
| `MAIN_SITE_ROOTS` | Start URLs for the main corporate site crawl |
| `GROUP_SITE_ROOTS` | Additional / subsidiary sites to crawl (optional) |
| `KEYWORDS` | The terms to flag in `keyword_hits.csv` (Chinese + English) |

The cninfo and szse.cn filing logic works for any Shenzhen-listed company by stock code; Shanghai-listed codes would need the SSE disclosure endpoint added.

## Design notes

- Fetches politely, one request at a time, with per-host pacing, retries, and `robots.txt` respect.
- Handles redirects, HTTPS→HTTP fallback, bot-challenge detection, and Chinese/UTF-8 encoding.
- Writes files atomically and keeps a manifest so interrupted runs resume cleanly.

## Scope & use

Intended for research and analysis of **publicly disclosed** information. Downloaded filings and site content remain the issuers' own copyrighted material — the output folder is **not** committed to this repo (`.gitignore`), and you are responsible for how you use what you collect.

## License

MIT
