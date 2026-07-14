"""Listings Project NYC sublet URL and suffix analysis.

Designed for Google Colab but also runnable as a normal Python script.
Outputs an Excel workbook with crawl results, suffix analysis, stripped-base
URL tests, predictor summaries, request logs, and errors.
"""

from __future__ import annotations

import argparse
import json
import random
import re
import time
import unicodedata
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import parse_qsl, urlencode, urljoin, urlsplit, urlunsplit

import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


LISTING_PATH_RE = re.compile(r"^/listings/[^/?#]+/?$", re.I)
UUID_RE = re.compile(
    r"(?i)(?:-)?([0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-"
    r"[89ab][0-9a-f]{3}-[0-9a-f]{12})$"
)
PRICE_RE = re.compile(
    r"\$\s?[\d,]+(?:\.\d{2})?(?:\s*/\s*(?:day|week|month))?", re.I
)
DATE_RANGE_RE = re.compile(
    r"([A-Z][a-z]+ \d{1,2}, \d{4})\s*(?:-|–|—|to)\s*"
    r"([A-Z][a-z]+ \d{1,2}, \d{4})"
)


@dataclass(frozen=True)
class Config:
    base: str = "https://www.listingsproject.com"
    browse_path: str = "/real-estate/new-york-city/sublets"
    output: str = "/content/drive/MyDrive/listings_real.xlsx"
    summary_json: str = "/content/drive/MyDrive/listings_real_summary.json"
    max_pages: int = 60
    request_timeout: int = 25
    delay_min: float = 0.8
    delay_max: float = 1.4
    stop_after_empty_pages: int = 2
    verify_listing_pages: bool = False
    use_browser_fallback: bool = True
    max_sitemaps: int = 100

    @property
    def browse_url(self) -> str:
        return urljoin(self.base, self.browse_path)


HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X) "
        "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.0 "
        "Mobile/15E148 Safari/604.1"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Cache-Control": "no-cache",
}


class ListingsProjectAnalyzer:
    def __init__(self, config: Config):
        self.config = config
        self.session = self._make_session()
        self.crawl_log: list[dict[str, Any]] = []
        self.errors: list[dict[str, Any]] = []

    @staticmethod
    def _make_session() -> requests.Session:
        retry = Retry(
            total=4,
            connect=4,
            read=4,
            status=4,
            backoff_factor=1.0,
            status_forcelist=(429, 500, 502, 503, 504),
            allowed_methods=frozenset({"GET", "HEAD"}),
            respect_retry_after_header=True,
            raise_on_status=False,
        )
        adapter = HTTPAdapter(max_retries=retry, pool_connections=10, pool_maxsize=10)
        session = requests.Session()
        session.headers.update(HEADERS)
        session.mount("https://", adapter)
        session.mount("http://", adapter)
        return session

    def pause(self) -> None:
        time.sleep(random.uniform(self.config.delay_min, self.config.delay_max))

    def fetch(
        self, url: str, *, allow_redirects: bool = True, purpose: str = "request"
    ) -> requests.Response | None:
        started = time.monotonic()
        try:
            response = self.session.get(
                url,
                timeout=self.config.request_timeout,
                allow_redirects=allow_redirects,
            )
            self.crawl_log.append(
                {
                    "Timestamp UTC": utc_now(),
                    "Purpose": purpose,
                    "Requested URL": url,
                    "Status": response.status_code,
                    "Final URL": response.url,
                    "Bytes": len(response.content),
                    "Seconds": round(time.monotonic() - started, 2),
                }
            )
            return response
        except requests.RequestException as exc:
            self.errors.append(
                {
                    "Stage": purpose,
                    "URL": url,
                    "Error Type": type(exc).__name__,
                    "Error": str(exc),
                }
            )
            return None

    def clean_url(self, href: str | None) -> str | None:
        if not href:
            return None
        full = urljoin(self.config.base, href.strip())
        parts = urlsplit(full)
        allowed_hosts = {
            urlsplit(self.config.base).netloc.lower(),
            urlsplit(self.config.base).netloc.lower().removeprefix("www."),
        }
        if parts.scheme not in {"http", "https"}:
            return None
        if parts.netloc.lower().removeprefix("www.") not in {
            host.removeprefix("www.") for host in allowed_hosts
        }:
            return None
        path = re.sub(r"/+", "/", parts.path).rstrip("/")
        if not LISTING_PATH_RE.match(path):
            return None
        canonical_host = urlsplit(self.config.base).netloc
        return urlunsplit(("https", canonical_host, path, "", ""))

    def page_url(self, page: int) -> str:
        if page == 1:
            return self.config.browse_url
        parts = urlsplit(self.config.browse_url)
        query = dict(parse_qsl(parts.query))
        query["page"] = str(page)
        return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(query), ""))

    def extract_card_record(self, anchor: Any, url: str, page: int) -> dict[str, Any]:
        card = enclosing_card(anchor)
        card_text = normalize_space(card.get_text(" ", strip=True)) or ""
        title = first_text(
            card,
            [
                "h4 a[href*='/listings/']",
                "h3 a[href*='/listings/']",
                "h2 a[href*='/listings/']",
                "h5 a[href*='/listings/']",
                "h4",
                "h3",
                "h2",
                "h5",
            ],
        )
        if not title:
            anchor_text = normalize_space(anchor.get_text(" ", strip=True))
            if anchor_text and anchor_text.casefold() not in {
                "see more",
                "read the full listing and contact details",
            }:
                title = anchor_text

        price_match = PRICE_RE.search(card_text)
        date_match = DATE_RANGE_RE.search(card_text)
        slug = slug_from_url(url)
        base_slug, suffix = split_uuid_suffix(slug)
        return {
            "Title": title,
            "Price": normalize_space(price_match.group(0)) if price_match else None,
            "Start": date_match.group(1) if date_match else None,
            "End": date_match.group(2) if date_match else None,
            "URL (real)": url,
            "Slug": slug,
            "Base Slug": base_slug,
            "Suffix": suffix,
            "Has Suffix": suffix is not None,
            "Source Page": page,
            "Method": "html",
        }

    def crawl_html(self) -> dict[str, dict[str, Any]]:
        records: dict[str, dict[str, Any]] = {}
        consecutive_empty = 0
        previous_page_urls: set[str] | None = None

        for page in range(1, self.config.max_pages + 1):
            response = self.fetch(self.page_url(page), purpose=f"browse page {page}")
            if response is None or response.status_code != 200:
                break

            soup = BeautifulSoup(response.text, "lxml")
            page_records: dict[str, dict[str, Any]] = {}
            for anchor in soup.select("a[href]"):
                listing_url = self.clean_url(anchor.get("href"))
                if listing_url:
                    merge_record(
                        page_records,
                        self.extract_card_record(anchor, listing_url, page),
                    )

            page_set = set(page_records)
            before = len(records)
            for record in page_records.values():
                merge_record(records, record)
            print(
                f"page {page:2}: {len(page_records):2} listing URLs, "
                f"+{len(records) - before:2} new, {len(records):3} total"
            )

            consecutive_empty = consecutive_empty + 1 if not page_set else 0
            if previous_page_urls is not None and page_set == previous_page_urls:
                print("Stopping: page repeated the previous page exactly.")
                break
            previous_page_urls = page_set

            if consecutive_empty >= self.config.stop_after_empty_pages:
                print("Stopping after repeated empty pages.")
                break
            if page > 1 and not find_next_href(soup, page + 1):
                print("Stopping: no next-page link found.")
                break
            self.pause()
        return records

    def crawl_embedded_data(self) -> dict[str, dict[str, Any]]:
        response = self.fetch(self.config.browse_url, purpose="embedded-data fallback")
        if response is None or response.status_code != 200:
            return {}
        patterns = [
            r'["\'](/listings/[^"\'?#<>\s]+)["\']',
            r'https?://(?:www\.)?listingsproject\.com(/listings/[^"\'?#<>\s]+)',
        ]
        paths: set[str] = set()
        for pattern in patterns:
            paths.update(re.findall(pattern, response.text, flags=re.I))
        return self.records_from_urls(paths, method="embedded data")

    def crawl_sitemaps(self) -> dict[str, dict[str, Any]]:
        records: dict[str, dict[str, Any]] = {}
        queue = [
            urljoin(self.config.base, "/sitemap.xml"),
            urljoin(self.config.base, "/sitemap_index.xml"),
        ]
        visited: set[str] = set()
        while queue and len(visited) < self.config.max_sitemaps:
            sitemap_url = queue.pop(0)
            if sitemap_url in visited:
                continue
            visited.add(sitemap_url)
            response = self.fetch(sitemap_url, purpose="sitemap fallback")
            if response is None or response.status_code != 200:
                continue
            xml = BeautifulSoup(response.text, "xml")
            for node in xml.find_all("loc"):
                location = normalize_space(node.get_text())
                if not location:
                    continue
                if urlsplit(location).path.lower().endswith(".xml"):
                    if location not in visited:
                        queue.append(location)
                    continue
                listing_url = self.clean_url(location)
                if listing_url:
                    merge_record(records, make_url_only_record(listing_url, "sitemap"))
            self.pause()
        return records

    def records_from_urls(
        self, urls: Iterable[str], *, method: str
    ) -> dict[str, dict[str, Any]]:
        records: dict[str, dict[str, Any]] = {}
        for raw_url in urls:
            listing_url = self.clean_url(raw_url)
            if listing_url:
                merge_record(records, make_url_only_record(listing_url, method))
        return records

    def extract_page_identity(self, html: str, final_url: str) -> dict[str, Any]:
        soup = BeautifulSoup(html, "lxml")
        canonical_tag = soup.select_one('link[rel="canonical"][href]')
        og_tag = soup.select_one('meta[property="og:url"][content]')
        page_title = normalize_space(soup.title.get_text(" ", strip=True)) if soup.title else None
        if page_title:
            page_title = re.sub(
                r"\s*\|\s*Listings Project\s*$", "", page_title, flags=re.I
            )
        return {
            "Resolved URL": self.clean_url(final_url) or final_url,
            "Canonical URL": self.clean_url(canonical_tag.get("href")) if canonical_tag else None,
            "OG URL": self.clean_url(og_tag.get("content")) if og_tag else None,
            "Page Heading": first_text(soup, ["main h1", "main h2", "h1", "h2"]),
            "Page Title": page_title,
        }

    def verify_listing_pages(self, df: pd.DataFrame) -> pd.DataFrame:
        checks: list[dict[str, Any]] = []
        for index, row in df.iterrows():
            response = self.fetch(row["URL (real)"], purpose="listing verification")
            info = {
                "Resolved URL": None,
                "Canonical URL": None,
                "OG URL": None,
                "Page Heading": None,
                "Page Title": None,
                "Listing Status": None,
            }
            if response is not None:
                info["Listing Status"] = response.status_code
                if response.status_code == 200:
                    info.update(self.extract_page_identity(response.text, response.url))
            checks.append(info)
            if (index + 1) % 20 == 0 or index + 1 == len(df):
                print(f"verified {index + 1}/{len(df)}")
            self.pause()
        return pd.concat([df.reset_index(drop=True), pd.DataFrame(checks)], axis=1)

    def test_base_urls(self, suffixed: pd.DataFrame) -> pd.DataFrame:
        results: list[dict[str, Any]] = []
        for _, row in suffixed.iterrows():
            base_url = urljoin(self.config.base, f"/listings/{row['Base Slug']}")
            response = self.fetch(base_url, purpose="base URL test")
            result = {
                "Current Listing Title": row["Title"],
                "Current Suffixed URL": row["URL (real)"],
                "Base URL Tested": base_url,
                "HTTP Status": None,
                "Redirected": None,
                "Final URL": None,
                "Canonical URL": None,
                "OG URL": None,
                "Base Page Heading": None,
                "Base Page Title": None,
                "Verdict": None,
                "Interpretation": None,
            }
            if response is None:
                result.update(Verdict="request error", Interpretation="No conclusion.")
                results.append(result)
                continue

            result["HTTP Status"] = response.status_code
            result["Redirected"] = bool(response.history)
            result["Final URL"] = response.url
            identity: dict[str, Any] = {}
            if response.status_code == 200:
                identity = self.extract_page_identity(response.text, response.url)
                result.update(
                    {
                        "Canonical URL": identity.get("Canonical URL"),
                        "OG URL": identity.get("OG URL"),
                        "Base Page Heading": identity.get("Page Heading"),
                        "Base Page Title": identity.get("Page Title"),
                    }
                )

            candidates = [response.url, identity.get("Canonical URL"), identity.get("OG URL")]
            points_to_current = any(
                same_clean_url(self, value, row["URL (real)"])
                for value in candidates
                if value
            )
            resolved = self.clean_url(response.url)
            resolves_elsewhere = bool(
                response.status_code == 200
                and resolved
                and not same_clean_url(self, resolved, row["URL (real)"])
            )
            returned_title = normalize_space(identity.get("Page Heading") or identity.get("Page Title"))
            current_title = normalize_space(row["Title"])
            title_differs = bool(
                current_title
                and returned_title
                and current_title.casefold() != returned_title.casefold()
            )

            if response.status_code == 200 and points_to_current:
                verdict = "base aliases or redirects to current suffixed listing"
                interpretation = "The suffix can currently be omitted for this listing."
            elif response.status_code == 200 and resolves_elsewhere and title_differs:
                verdict = "base resolves to a different listing"
                interpretation = "Current namespace collision confirmed."
            elif response.status_code == 200 and resolves_elsewhere:
                verdict = "base resolves elsewhere; identity uncertain"
                interpretation = "Possible collision, but page identity was not decisive."
            elif response.status_code in {404, 410}:
                verdict = f"base unavailable ({response.status_code})"
                interpretation = (
                    "No current collision is visible; a historical collision remains possible."
                )
            elif response.status_code == 200:
                verdict = "base returns 200; identity unresolved"
                interpretation = "Inspect the URL and page-identity columns."
            else:
                verdict = f"HTTP {response.status_code}"
                interpretation = "No reliable collision conclusion."
            result.update(Verdict=verdict, Interpretation=interpretation)
            results.append(result)
            self.pause()
        return pd.DataFrame(results)

    def run(self) -> dict[str, Any]:
        print("PART A — crawl current listing URLs")
        records = self.crawl_html()
        method = "html"
        if not records:
            print("Trying embedded page data.")
            records = self.crawl_embedded_data()
            method = "embedded data"
        if not records:
            print("Trying sitemap discovery.")
            records = self.crawl_sitemaps()
            method = "sitemap"
        if not records:
            raise RuntimeError(
                "No listing URLs were recovered. The browse URL or site structure may have changed."
            )

        df = pd.DataFrame(records.values()).sort_values(
            by=["Source Page", "Title", "URL (real)"], na_position="last"
        ).reset_index(drop=True)
        if self.config.verify_listing_pages:
            df = self.verify_listing_pages(df)

        df["Expected Slug (approx.)"] = df["Title"].apply(slugify_title)
        df["Base Matches Approx. Title Slug"] = (
            df["Expected Slug (approx.)"].notna()
            & (df["Base Slug"] == df["Expected Slug (approx.)"])
        )
        df["Slug Comparison"] = df.apply(classify_slug, axis=1)
        suffixed = df[df["Has Suffix"]].copy()
        base_tests = self.test_base_urls(suffixed)
        predictor_summary = build_predictor_summary(df)

        summary = {
            "run_time_utc": utc_now(),
            "browse_url": self.config.browse_url,
            "primary_method": method,
            "unique_listing_urls": int(len(df)),
            "listings_with_titles": int(df["Title"].notna().sum()),
            "suffixed_listings": int(len(suffixed)),
            "suffixed_percentage": float(len(suffixed) / len(df)) if len(df) else None,
            "base_urls_tested": int(len(base_tests)),
            "recorded_request_errors": int(len(self.errors)),
        }
        self.save_workbook(df, suffixed, base_tests, predictor_summary, summary)
        Path(self.config.summary_json).parent.mkdir(parents=True, exist_ok=True)
        Path(self.config.summary_json).write_text(
            json.dumps({"config": asdict(self.config), "summary": summary}, indent=2),
            encoding="utf-8",
        )
        print(f"Saved: {self.config.output}")
        print(f"Saved: {self.config.summary_json}")
        return {"data": df, "summary": summary}

    def save_workbook(
        self,
        df: pd.DataFrame,
        suffixed: pd.DataFrame,
        base_tests: pd.DataFrame,
        predictor_summary: pd.DataFrame,
        summary: dict[str, Any],
    ) -> None:
        output = Path(self.config.output)
        output.parent.mkdir(parents=True, exist_ok=True)
        summary_df = pd.DataFrame(
            [{"Metric": key.replace("_", " ").title(), "Value": value} for key, value in summary.items()]
        )
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            summary_df.to_excel(writer, sheet_name="Run Summary", index=False)
            df.to_excel(writer, sheet_name="All Listings", index=False)
            suffixed.to_excel(writer, sheet_name="Suffixed Only", index=False)
            base_tests.to_excel(writer, sheet_name="Base URL Tests", index=False)
            predictor_summary.to_excel(writer, sheet_name="Predictor Summary")
            pd.DataFrame(self.crawl_log).to_excel(writer, sheet_name="Crawl Log", index=False)
            pd.DataFrame(self.errors).to_excel(writer, sheet_name="Errors", index=False)
            format_workbook(writer.book)


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def normalize_space(value: Any) -> str | None:
    if value is None:
        return None
    text = " ".join(str(value).split()).strip()
    return text or None


def slug_from_url(url: str) -> str:
    return urlsplit(url).path.rstrip("/").rsplit("/", 1)[-1]


def split_uuid_suffix(slug: str) -> tuple[str, str | None]:
    match = UUID_RE.search(str(slug))
    if not match:
        return str(slug), None
    return str(slug)[: match.start()].rstrip("-"), match.group(1).lower()


def slugify_title(title: Any) -> str | None:
    if title is None or pd.isna(title):
        return None
    text = unicodedata.normalize("NFKD", str(title))
    text = text.encode("ascii", "ignore").decode("ascii").lower()
    return re.sub(r"[^a-z0-9]+", "-", text).strip("-")


def first_text(node: Any, selectors: list[str]) -> str | None:
    for selector in selectors:
        found = node.select_one(selector)
        if found:
            value = normalize_space(found.get_text(" ", strip=True))
            if value:
                return value
    return None


def enclosing_card(anchor: Any) -> Any:
    for parent in anchor.parents:
        if not getattr(parent, "name", None) or parent.name in {"body", "html"}:
            break
        text = normalize_space(parent.get_text(" ", strip=True)) or ""
        if parent.select_one("h2, h3, h4, h5") and (
            PRICE_RE.search(text) or DATE_RANGE_RE.search(text)
        ):
            return parent
    return anchor.parent or anchor


def find_next_href(soup: BeautifulSoup, expected_page: int) -> str | None:
    for anchor in soup.select("a[href]"):
        text = normalize_space(anchor.get_text(" ", strip=True)) or ""
        rel_value = anchor.get("rel", [])
        rel = " ".join(rel_value) if isinstance(rel_value, list) else str(rel_value)
        href = anchor.get("href", "")
        if (
            "next" in rel.casefold()
            or text in {">", "›", "→", "Next", "next"}
            or re.search(rf"(?:\?|&)page={expected_page}(?:&|$)", href)
        ):
            return href
    return None


def make_url_only_record(url: str, method: str) -> dict[str, Any]:
    slug = slug_from_url(url)
    base_slug, suffix = split_uuid_suffix(slug)
    return {
        "Title": None,
        "Price": None,
        "Start": None,
        "End": None,
        "URL (real)": url,
        "Slug": slug,
        "Base Slug": base_slug,
        "Suffix": suffix,
        "Has Suffix": suffix is not None,
        "Source Page": None,
        "Method": method,
    }


def record_quality(record: dict[str, Any]) -> int:
    return sum(
        [
            4 if record.get("Title") else 0,
            2 if record.get("Price") else 0,
            1 if record.get("Start") else 0,
            1 if record.get("End") else 0,
        ]
    )


def merge_record(store: dict[str, dict[str, Any]], record: dict[str, Any]) -> None:
    url = record["URL (real)"]
    current = store.get(url)
    if current is None or record_quality(record) > record_quality(current):
        store[url] = record


def same_clean_url(analyzer: ListingsProjectAnalyzer, first: str, second: str) -> bool:
    clean_first = analyzer.clean_url(first)
    clean_second = analyzer.clean_url(second)
    return bool(clean_first and clean_second and clean_first == clean_second)


def classify_slug(row: pd.Series) -> str:
    has_title = pd.notna(row["Expected Slug (approx.)"])
    if not has_title:
        return "title unavailable"
    has_suffix = bool(row["Has Suffix"])
    matches = bool(row["Base Matches Approx. Title Slug"])
    if has_suffix and matches:
        return "suffixed; base matches approximate title slug"
    if has_suffix:
        return "suffixed; base differs from approximate title slug"
    if matches:
        return "unsuffixed; slug matches approximate title slug"
    return "unsuffixed; slug differs from approximate title slug"


def title_features(title: Any) -> dict[str, Any]:
    text = "" if title is None or pd.isna(title) else str(title)
    approximate_slug = slugify_title(text) or ""
    return {
        "Title Characters": len(text),
        "Title Words": len(text.split()),
        "Approx. Slug Characters": len(approximate_slug),
        "Contains Digit": bool(re.search(r"\d", text)),
        "Contains Punctuation": bool(re.search(r"[|()—–:!?,;&/+]", text)),
        "Contains Non-ASCII": any(ord(character) > 127 for character in text),
        "All Caps Word Count": sum(
            word.isupper() for word in re.findall(r"\b[A-Za-z]{2,}\b", text)
        ),
    }


def build_predictor_summary(df: pd.DataFrame) -> pd.DataFrame:
    feature_df = df["Title"].apply(lambda value: pd.Series(title_features(value)))
    feature_df["Has Suffix"] = df["Has Suffix"].astype(bool)
    feature_df["Title Available"] = df["Title"].notna()
    usable = feature_df[feature_df["Title Available"]].drop(columns="Title Available")
    if len(usable) and usable["Has Suffix"].nunique() == 2:
        summary = (
            usable.groupby("Has Suffix")
            .mean(numeric_only=True)
            .T.rename(columns={False: "Unsuffixed Mean", True: "Suffixed Mean"})
        )
        summary["Absolute Difference"] = (
            summary["Suffixed Mean"] - summary["Unsuffixed Mean"]
        ).abs()
        return summary
    return pd.DataFrame(
        {"Note": ["Comparison requires titles and listings in both suffix classes."]}
    )


def format_workbook(workbook: Any) -> None:
    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column_cells in sheet.columns:
            values = [str(cell.value) for cell in column_cells if cell.value is not None]
            maximum = max((len(value) for value in values), default=10)
            sheet.column_dimensions[column_cells[0].column_letter].width = min(
                max(maximum + 2, 12), 60
            )


def mount_drive_if_colab() -> None:
    try:
        from google.colab import drive  # type: ignore
    except ImportError:
        return
    drive.mount("/content/drive")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", default=Config.output)
    parser.add_argument("--summary-json", default=Config.summary_json)
    parser.add_argument("--max-pages", type=int, default=Config.max_pages)
    parser.add_argument("--verify-listings", action="store_true")
    parser.add_argument("--no-drive-mount", action="store_true")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if not args.no_drive_mount:
        mount_drive_if_colab()
    config = Config(
        output=args.output,
        summary_json=args.summary_json,
        max_pages=args.max_pages,
        verify_listing_pages=args.verify_listings,
    )
    result = ListingsProjectAnalyzer(config).run()
    try:
        from IPython.display import display

        display(result["data"].head(10))
    except ImportError:
        print(result["data"].head(10).to_string(index=False))


if __name__ == "__main__":
    main()
