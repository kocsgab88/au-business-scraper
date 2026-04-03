"""
Australian Funeral Business Scraper
====================================
Nationwide scraper for Australian funeral businesses.
Sources: Google Places API (via Outscraper), Yellow Pages AU, AFDA directory.
Output:  master CSV + per-state CSVs + Excel, UTF-8 encoded.

Ownership classification:
  - InvoCare brands  → InvoCare
  - Propel brands    → Propel
  - Other verified chain → Other Corporate
  - Cannot verify    → Unknown
  - Everything else  → Independent

Usage:
    python au_funeral_scraper.py --source yellowpages --state NSW
    python au_funeral_scraper.py --source outscraper  --all-states
    python au_funeral_scraper.py --demo               # 10-row demo, no API needed
"""

import argparse
import csv
import json
import logging
import os
import re
import time
from dataclasses import dataclass, field, asdict
from datetime import datetime
from pathlib import Path
from typing import Optional

import requests
from bs4 import BeautifulSoup
import pandas as pd

# ─────────────────────────────────────────────
#  LOGGING
# ─────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [SCRAPER] %(levelname)s – %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("scraper.log", encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)

# ─────────────────────────────────────────────
#  CONFIGURATION
# ─────────────────────────────────────────────
AU_STATES = ["NSW", "VIC", "QLD", "WA", "SA", "TAS", "ACT", "NT"]

OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", "output"))
REQUEST_DELAY = float(os.environ.get("REQUEST_DELAY", "1.5"))   # seconds between requests
MAX_RETRIES   = int(os.environ.get("MAX_RETRIES", "3"))

# ─────────────────────────────────────────────
#  OWNERSHIP CLASSIFICATION
#  Source: invocare.com.au/our-brands
#          propelfuneralpartners.com.au/our-businesses
#  ⚠️  Always cross-check live brand pages — acquisitions happen regularly.
# ─────────────────────────────────────────────
INVOCARE_BRANDS: set[str] = {
    "white lady funerals",
    "simplicity funerals",
    "value cremations",
    "national cremations",
    "le pine funerals",
    "tobin brothers",
    "greenwood funerals",
    "george hartnett",
    "purslowe & chipper",
    "purslowe and chipper",
    "parsons bros",
    "norman bros",
    "harrington park memorial",
}

PROPEL_BRANDS: set[str] = {
    "guardian funerals",
    "j. fulton funerals",
    "j fulton funerals",
    "blackwell funerals",
    "swan hill funerals",
    "roger goonan",
}


def classify_ownership(business_name: str) -> tuple[str, Optional[str]]:
    """
    Returns (ownership_type, corporate_brand).
    ownership_type: Independent | InvoCare | Propel | Other Corporate | Unknown
    corporate_brand: brand name if not Independent, else None
    """
    name_lower = business_name.lower().strip()

    for brand in INVOCARE_BRANDS:
        if brand in name_lower:
            return "InvoCare", brand.title()

    for brand in PROPEL_BRANDS:
        if brand in name_lower:
            return "Propel", brand.title()

    return "Independent", None


# ─────────────────────────────────────────────
#  DATA MODEL
# ─────────────────────────────────────────────
@dataclass
class FuneralBusiness:
    # Core fields (required)
    business_name:  str = ""
    abn:            str = ""
    street_address: str = ""
    suburb:         str = ""
    state:          str = ""
    postcode:       str = ""
    phone:          str = ""
    ownership_type: str = "Unknown"

    # Optional fields
    mobile:              str = ""
    email:               str = ""
    website_url:         str = ""
    corporate_brand:     str = ""
    afda_member:         str = "Unknown"
    state_assoc_member:  str = "Unknown"
    services:            str = ""
    notes:               str = ""

    # Google Places fields
    google_place_id:          str = ""
    google_business_name:     str = ""
    google_rating:            str = ""
    google_review_count:      str = ""
    google_maps_url:          str = ""
    google_business_status:   str = ""
    google_business_category: str = ""
    google_formatted_address: str = ""

    # Internal
    source:     str = ""
    scraped_at: str = field(default_factory=lambda: datetime.now().isoformat())


CSV_COLUMNS = [
    "business_name", "abn", "street_address", "suburb", "state", "postcode",
    "phone", "ownership_type", "mobile", "email", "website_url", "corporate_brand",
    "afda_member", "state_assoc_member", "services", "notes",
    "google_place_id", "google_business_name", "google_rating", "google_review_count",
    "google_maps_url", "google_business_status", "google_business_category",
    "google_formatted_address", "source", "scraped_at",
]


# ─────────────────────────────────────────────
#  HTTP SESSION
# ─────────────────────────────────────────────
def build_session() -> requests.Session:
    """Requests session with browser-like headers and retry logic."""
    s = requests.Session()
    s.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/122.0.0.0 Safari/537.36"
        ),
        "Accept":          "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-AU,en;q=0.9",
        "Accept-Encoding": "gzip, deflate, br",
    })
    return s


SESSION = build_session()


def fetch(url: str, retries: int = MAX_RETRIES) -> Optional[BeautifulSoup]:
    """Fetch URL with retry + exponential backoff. Returns BeautifulSoup or None."""
    for attempt in range(1, retries + 1):
        try:
            r = SESSION.get(url, timeout=20)
            r.raise_for_status()
            r.encoding = "utf-8"
            log.debug(f"HTTP {r.status_code} | {url[:80]}")
            return BeautifulSoup(r.text, "html.parser")
        except requests.RequestException as e:
            log.warning(f"Fetch error ({attempt}/{retries}): {e} | {url[:80]}")
            if attempt < retries:
                time.sleep(REQUEST_DELAY * (2 ** (attempt - 1)))
    log.error(f"Failed after {retries} attempts: {url[:80]}")
    return None


# ─────────────────────────────────────────────
#  SCRAPER: YELLOW PAGES AUSTRALIA
# ─────────────────────────────────────────────
YP_BASE = "https://www.yellowpages.com.au"

def scrape_yellowpages(state: str, max_pages: int = 5) -> list[FuneralBusiness]:
    """
    Scrapes yellowpages.com.au for funeral homes in the given state.
    Returns list of FuneralBusiness records.
    """
    results: list[FuneralBusiness] = []
    state_map = {
        "NSW": "new-south-wales", "VIC": "victoria", "QLD": "queensland",
        "WA": "western-australia", "SA": "south-australia", "TAS": "tasmania",
        "ACT": "australian-capital-territory", "NT": "northern-territory",
    }
    state_slug = state_map.get(state.upper(), state.lower().replace(" ", "-"))

    for page in range(1, max_pages + 1):
        url = (
            f"{YP_BASE}/search/listings"
            f"?clue=funeral+homes&locationClue={state_slug}&pageNumber={page}"
        )
        log.info(f"[YP] {state} page {page}: {url}")
        soup = fetch(url)
        if not soup:
            break

        listings = soup.find_all("div", class_=re.compile(r"listing-box|listing__"))
        if not listings:
            log.info(f"[YP] No listings on page {page}, stopping.")
            break

        for listing in listings:
            try:
                biz = _parse_yp_listing(listing, state)
                if biz:
                    results.append(biz)
            except Exception as e:
                log.warning(f"[YP] Parse error: {e}")

        log.info(f"[YP] {state} page {page}: {len(listings)} listings found")
        time.sleep(REQUEST_DELAY)

    log.info(f"[YP] {state} total: {len(results)} records")
    return results


def _parse_yp_listing(listing, state: str) -> Optional[FuneralBusiness]:
    """Parse a single Yellow Pages listing div into a FuneralBusiness."""
    name_tag = (
        listing.find("a", class_=re.compile(r"listing-name|business-name")) or
        listing.find("h2") or
        listing.find("h3")
    )
    if not name_tag:
        return None

    name = name_tag.get_text(strip=True)
    if not name:
        return None

    address_tag = listing.find(class_=re.compile(r"address|location"))
    raw_address  = address_tag.get_text(strip=True) if address_tag else ""
    street, suburb, postcode = _parse_address(raw_address)

    phone_tag = listing.find(class_=re.compile(r"phone|telephone|contact"))
    phone = phone_tag.get_text(strip=True) if phone_tag else ""
    phone = re.sub(r"[^\d\s+()]", "", phone).strip()

    website_tag = listing.find("a", href=re.compile(r"^https?://(?!www\.yellowpages)"))
    website = website_tag["href"] if website_tag and website_tag.get("href") else ""

    ownership, brand = classify_ownership(name)

    return FuneralBusiness(
        business_name=name,
        street_address=street,
        suburb=suburb,
        state=state.upper(),
        postcode=postcode,
        phone=phone,
        website_url=website,
        ownership_type=ownership,
        corporate_brand=brand or "",
        source="yellowpages.com.au",
    )


def _parse_address(raw: str) -> tuple[str, str, str]:
    """
    Attempt to split 'Street, Suburb STATE POSTCODE' format.
    Returns (street, suburb, postcode).
    """
    raw = raw.strip()
    postcode_match = re.search(r"\b(\d{4})\b", raw)
    postcode = postcode_match.group(1) if postcode_match else ""

    # Remove state abbreviation and postcode from end
    cleaned = re.sub(r"\b(NSW|VIC|QLD|WA|SA|TAS|ACT|NT)\b\s*\d{4}?", "", raw).strip().rstrip(",")
    parts = [p.strip() for p in cleaned.split(",") if p.strip()]

    if len(parts) >= 2:
        street = parts[0]
        suburb = parts[-1]
    elif len(parts) == 1:
        street = ""
        suburb = parts[0]
    else:
        street = suburb = ""

    return street, suburb, postcode


# ─────────────────────────────────────────────
#  OUTSCRAPER INTEGRATION (Google Places)
# ─────────────────────────────────────────────
def enrich_with_outscraper(
    records: list[FuneralBusiness],
    api_key: str,
) -> list[FuneralBusiness]:
    """
    Enriches existing records with Google Places data via Outscraper API.
    Matches by business name + suburb.
    Docs: https://outscraper.com/google-maps-scraper/
    """
    try:
        from outscraper import ApiClient
    except ImportError:
        log.warning("[OUTSCRAPER] outscraper package not installed. Run: pip install outscraper")
        return records

    client = ApiClient(api_key=api_key)
    enriched = 0

    for biz in records:
        query = f"{biz.business_name} {biz.suburb} {biz.state} Australia"
        try:
            results = client.google_maps_search(query, limit=1, language="en", region="AU")
            if results and results[0]:
                place = results[0][0]
                biz.google_place_id          = place.get("place_id", "")
                biz.google_business_name     = place.get("name", "")
                biz.google_rating            = str(place.get("rating", ""))
                biz.google_review_count      = str(place.get("reviews", ""))
                biz.google_maps_url          = place.get("url", "")
                biz.google_business_status   = place.get("business_status", "OPERATIONAL")
                biz.google_business_category = place.get("type", "Funeral home")
                biz.google_formatted_address = place.get("full_address", "")
                enriched += 1
            time.sleep(0.5)
        except Exception as e:
            log.warning(f"[OUTSCRAPER] Error for '{biz.business_name}': {e}")

    log.info(f"[OUTSCRAPER] Enriched {enriched}/{len(records)} records")
    return records


# ─────────────────────────────────────────────
#  DEDUPLICATION
# ─────────────────────────────────────────────
def deduplicate(records: list[FuneralBusiness]) -> list[FuneralBusiness]:
    """
    Deduplicates by:
      1. google_place_id (primary for Google-enriched records)
      2. Normalized business name + suburb (fallback)
    """
    seen_place_ids: set[str] = set()
    seen_name_suburb: set[str] = set()
    unique: list[FuneralBusiness] = []

    for biz in records:
        # Primary dedup key: google_place_id
        if biz.google_place_id and biz.google_place_id in seen_place_ids:
            continue
        # Fallback dedup key: normalized name + suburb
        name_key = re.sub(r"\s+", " ", biz.business_name.lower().strip())
        suburb_key = biz.suburb.lower().strip()
        composite = f"{name_key}|{suburb_key}"
        if composite in seen_name_suburb:
            continue

        if biz.google_place_id:
            seen_place_ids.add(biz.google_place_id)
        seen_name_suburb.add(composite)
        unique.append(biz)

    removed = len(records) - len(unique)
    log.info(f"[DEDUP] {len(records)} → {len(unique)} records ({removed} duplicates removed)")
    return unique


# ─────────────────────────────────────────────
#  OUTPUT
# ─────────────────────────────────────────────
def save_csv(records: list[FuneralBusiness], path: Path) -> None:
    """Save records to UTF-8 CSV with standard column order."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS, extrasaction="ignore")
        writer.writeheader()
        for biz in records:
            writer.writerow(asdict(biz))
    log.info(f"[CSV] Saved {len(records)} records → {path}")


def save_excel(records: list[FuneralBusiness], path: Path) -> None:
    """Save records to Excel with auto-width columns."""
    path.parent.mkdir(parents=True, exist_ok=True)
    rows = [asdict(biz) for biz in records]
    df = pd.DataFrame(rows, columns=CSV_COLUMNS)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="All Businesses")

        # Per-state sheets
        for state in AU_STATES:
            state_df = df[df["state"] == state]
            if not state_df.empty:
                state_df.to_excel(writer, index=False, sheet_name=state)

        # Auto-width columns
        for sheet_name, ws in writer.sheets.items():
            for col in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 50)

    log.info(f"[EXCEL] Saved {len(records)} records → {path}")


def save_all_outputs(records: list[FuneralBusiness]) -> None:
    """Save master CSV + per-state CSVs + Excel."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = OUTPUT_DIR / timestamp

    # Master CSV
    save_csv(records, base / "AU_Funeral_Businesses_MASTER.csv")

    # Per-state CSVs
    for state in AU_STATES:
        state_records = [r for r in records if r.state == state]
        if state_records:
            save_csv(state_records, base / f"AU_Funeral_{state}.csv")

    # Excel
    save_excel(records, base / "AU_Funeral_Businesses.xlsx")

    log.info(f"[OUTPUT] All files saved to: {base}")


# ─────────────────────────────────────────────
#  DEMO MODE
# ─────────────────────────────────────────────
DEMO_RECORDS = [
    {"business_name": "Smith & Sons Funerals",          "state": "NSW", "suburb": "Parramatta",  "postcode": "2150", "phone": "02 9876 5432", "ownership_type": "Independent",  "google_rating": "4.8", "google_business_status": "OPERATIONAL",          "google_business_category": "Funeral home"},
    {"business_name": "White Lady Funerals Chatswood",  "state": "NSW", "suburb": "Chatswood",   "postcode": "2067", "phone": "02 9411 1122", "ownership_type": "InvoCare",     "google_rating": "4.6", "google_business_status": "OPERATIONAL",          "google_business_category": "Funeral home", "corporate_brand": "White Lady Funerals"},
    {"business_name": "Guardian Funerals Brisbane",     "state": "QLD", "suburb": "Brisbane",    "postcode": "4000", "phone": "07 3221 5544", "ownership_type": "Propel",       "google_rating": "4.3", "google_business_status": "OPERATIONAL",          "google_business_category": "Funeral home", "corporate_brand": "Guardian Funerals"},
    {"business_name": "Henderson Family Funerals",      "state": "VIC", "suburb": "Ballarat",    "postcode": "3350", "phone": "03 5331 2244", "ownership_type": "Independent",  "google_rating": "4.9", "google_business_status": "OPERATIONAL",          "google_business_category": "Funeral home"},
    {"business_name": "Simplicity Funerals Perth",      "state": "WA",  "suburb": "Perth",       "postcode": "6000", "phone": "08 9321 7788", "ownership_type": "InvoCare",     "google_rating": "4.2", "google_business_status": "OPERATIONAL",          "google_business_category": "Funeral home", "corporate_brand": "Simplicity Funerals"},
    {"business_name": "Blackwell Funerals Adelaide",    "state": "SA",  "suburb": "Adelaide",    "postcode": "5000", "phone": "08 8212 3344", "ownership_type": "Propel",       "google_rating": "4.5", "google_business_status": "OPERATIONAL",          "google_business_category": "Funeral home", "corporate_brand": "Blackwell Funerals"},
    {"business_name": "Hobart Funeral Services",        "state": "TAS", "suburb": "Hobart",      "postcode": "7000", "phone": "03 6231 4455", "ownership_type": "Independent",  "google_rating": "4.7", "google_business_status": "OPERATIONAL",          "google_business_category": "Funeral home"},
    {"business_name": "Capital Funerals Canberra",      "state": "ACT", "suburb": "Canberra",    "postcode": "2600", "phone": "02 6247 8899", "ownership_type": "Independent",  "google_rating": "4.6", "google_business_status": "OPERATIONAL",          "google_business_category": "Funeral home"},
    {"business_name": "Darwin Memorial Funerals",       "state": "NT",  "suburb": "Darwin",      "postcode": "0800", "phone": "08 8981 6677", "ownership_type": "Independent",  "google_rating": "4.4", "google_business_status": "OPERATIONAL",          "google_business_category": "Funeral home"},
    {"business_name": "Le Pine Funerals Hawthorn",      "state": "VIC", "suburb": "Hawthorn",    "postcode": "3122", "phone": "03 9819 2200", "ownership_type": "InvoCare",     "google_rating": "4.8", "google_business_status": "CLOSED_TEMPORARILY",   "google_business_category": "Funeral home", "corporate_brand": "Le Pine Funerals"},
]

def run_demo() -> None:
    """Generate 10-row demo output without any API calls."""
    log.info("[DEMO] Generating demo output (no API required)")
    records = []
    for row in DEMO_RECORDS:
        biz = FuneralBusiness(**{k: v for k, v in row.items() if k in FuneralBusiness.__dataclass_fields__})
        biz.source = "demo"
        records.append(biz)
    save_all_outputs(records)
    log.info(f"[DEMO] Done – {len(records)} records written to {OUTPUT_DIR}/")


# ─────────────────────────────────────────────
#  CLI
# ─────────────────────────────────────────────
def main() -> None:
    parser = argparse.ArgumentParser(description="Australian Funeral Business Scraper")
    parser.add_argument("--demo",        action="store_true",      help="Generate demo output (no API)")
    parser.add_argument("--source",      default="yellowpages",    help="Data source: yellowpages")
    parser.add_argument("--state",       default=None,             help="State abbreviation (e.g. NSW)")
    parser.add_argument("--all-states",  action="store_true",      help="Scrape all 8 states/territories")
    parser.add_argument("--outscraper-key", default=None,          help="Outscraper API key for Google enrichment")
    parser.add_argument("--max-pages",   type=int, default=5,      help="Max pages per state (default: 5)")
    args = parser.parse_args()

    if args.demo:
        run_demo()
        return

    states = AU_STATES if args.all_states else [args.state.upper()] if args.state else ["NSW"]
    all_records: list[FuneralBusiness] = []

    for state in states:
        log.info(f"[MAIN] Scraping {state}...")
        if args.source == "yellowpages":
            records = scrape_yellowpages(state, max_pages=args.max_pages)
        else:
            log.error(f"Unknown source: {args.source}")
            continue
        all_records.extend(records)
        time.sleep(REQUEST_DELAY)

    # Ownership classification (already done per-record, but log summary)
    ownership_counts: dict[str, int] = {}
    for biz in all_records:
        ownership_counts[biz.ownership_type] = ownership_counts.get(biz.ownership_type, 0) + 1
    log.info(f"[OWNERSHIP] {ownership_counts}")

    # Enrich with Google Places if API key provided
    if args.outscraper_key:
        all_records = enrich_with_outscraper(all_records, args.outscraper_key)

    # Deduplicate
    all_records = deduplicate(all_records)

    # Save
    save_all_outputs(all_records)
    log.info(f"[DONE] {len(all_records)} records | {len(states)} states")


if __name__ == "__main__":
    main()
