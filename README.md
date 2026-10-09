# Jereh Scraper

A single-file Python scraper that downloads everything public about **Jereh Group (SZSE: 002353)** into one folder — exchange filings, company websites, and an extracted text + keyword index.

## What it collects

1. **Exchange filings** — every announcement since listing plus investor-relations meeting records, from [cninfo.com.cn](https://www.cninfo.com.cn) (China's official disclosure site). If cninfo fails for a given year, that year is pulled from [szse.cn](https://www.szse.cn) (the Shenzhen Stock Exchange) instead.
2. **Websites** — every page (saved as text) and every document (PDF, Word, Excel, PowerPoint, ZIP) on `jereh.com` (Chinese + English) and the Jereh group sites.
3. **Text & keywords** — extracts the text of every PDF and Office file and flags documents that mention a configurable keyword list (J&F, FTAI, gas turbines, data centers, etc.).

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
| `--main-site-only` | Skip the Jereh group sites |
| `--max-pages 500` | Stop the website crawl after N pages |
| `--delay 1.0` | Seconds between requests to the same site |
| `--keywords "Kawasaki,LM2500"` | Extra keywords to flag |
| `--out ./jereh_docs` | Output folder |
| `--filings-from auto\|cninfo\|szse\|both` | Filing source strategy |
| `--save-html` | Also keep the raw HTML of every page |
| `--fresh` | Restart the website crawl from scratch |

A full run takes several hours. It is **resumable** — stop with `Ctrl+C` and run the same command again; finished downloads are skipped, failed ones retried, and the website crawl picks up where it stopped.

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

## Notes

- Fetches politely, one request at a time, with per-host pacing, retries, and `robots.txt` respect.
- Downloaded filings are the companies' own copyrighted disclosures; this tool is for research use. The output folder is **not** committed to this repo.

## License

MIT
