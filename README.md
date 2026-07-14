# Listings Project URL Analyzer

Google Colab-compatible crawler and analysis tool for current Listings Project NYC sublet URLs.

## What it does

1. Crawls the current NYC sublet browse pages.
2. Extracts listing URLs and available card metadata.
3. Detects UUID-suffixed slugs.
4. Compares each slug base with an approximate title-derived slug.
5. Tests suffix-stripped URLs and records redirects, canonical URLs, page identity, and HTTP outcomes.
6. Compares simple title features between suffixed and unsuffixed listings.
7. Writes a structured Excel workbook and a JSON run summary.

The analysis deliberately distinguishes current evidence from historical possibilities. A stripped URL resolving to another listing is direct evidence of a current namespace collision. A 404 or 410 does not establish why the suffix was assigned.

## Google Colab

Run these cells:

```python
!git clone https://github.com/muthy5/nyc-listings.git
%cd nyc-listings
!pip -q install -r requirements.txt
%run listings_project.py
```

The script mounts Google Drive and writes:

- `/content/drive/MyDrive/listings_real.xlsx`
- `/content/drive/MyDrive/listings_real_summary.json`

## Command-line options

```bash
python listings_project.py \
  --output ./listings_real.xlsx \
  --summary-json ./listings_real_summary.json \
  --max-pages 60 \
  --no-drive-mount
```

Add `--verify-listings` to fetch each recovered listing page for page-title, heading, canonical-URL, and resolved-URL checks. This is slower and creates substantially more requests.

## Workbook sheets

- `Run Summary`
- `All Listings`
- `Suffixed Only`
- `Base URL Tests`
- `Predictor Summary`
- `Crawl Log`
- `Errors`

## Design changes from the original one-cell version

- Configuration is centralized in a dataclass.
- Crawling, parsing, analysis, URL testing, and export are separated into testable functions.
- HTTP retries, request logs, and errors are retained.
- Duplicate records are resolved by metadata quality.
- URL identity checks use the final URL, canonical URL, and Open Graph URL.
- Output includes a machine-readable JSON summary.
- The module is import-safe and can run in Colab or as a normal Python script.

## Important limitation

The fallback sitemap may include listings outside the current NYC-sublet result set or listings no longer visible in the browse pages. The `Method` column identifies how each record was recovered. Current public data cannot reconstruct deleted or unpublished historical listings, so it cannot prove the original cause of every UUID suffix.
