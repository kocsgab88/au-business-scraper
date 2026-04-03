# 🇦🇺 Australian Business Directory Scraper

Production-ready Python scraper for building verified Australian business databases with Google Places enrichment and ownership classification.

**Built for:** [eziFunerals](https://ezifunerals.com.au) — Australia's independent funeral planning platform.

---

## Features

- **Multi-source scraping** — Yellow Pages AU, AFDA directory, state associations
- **Google Places enrichment** — ratings, review counts, business status, place IDs via [Outscraper](https://outscraper.com)
- **Ownership classification** — InvoCare / Propel / Independent / Other Corporate / Unknown
- **Deduplication** — by ABN (primary) and Google Place ID + normalized name
- **Per-state output** — 8 separate CSVs (NSW, VIC, QLD, WA, SA, TAS, ACT, NT) + master CSV + Excel
- **Bot-detection handling** — browser-like headers, request throttling, exponential backoff retry
- **Demo mode** — 10-row sample output without any API calls

---

## Output Format

Each record contains 26 fields matching the eziFunerals Import Template:

| Field | Example | Notes |
|-------|---------|-------|
| `business_name` | Smith & Sons Funerals | Trading name |
| `state` | NSW | Abbreviation |
| `postcode` | 2150 | 4 digits |
| `phone` | 02 9876 5432 | With area code |
| `ownership_type` | Independent | See classification below |
| `google_place_id` | ChIJN1t_tDeu... | Primary dedup key |
| `google_rating` | 4.8 | 0.0–5.0 |
| `google_business_status` | OPERATIONAL | OPERATIONAL / CLOSED_TEMPORARILY / CLOSED_PERMANENTLY |
| `corporate_brand` | White Lady Funerals | Only if not Independent |

Full field reference: see [`field_reference.md`](field_reference.md)

---

## Ownership Classification

Cross-referenced against live brand pages (updated regularly):

| Type | Source |
|------|--------|
| **InvoCare** | [invocare.com.au/our-brands](https://www.invocare.com.au/our-brands) |
| **Propel** | [propelfuneralpartners.com.au/our-businesses](https://www.propelfuneralpartners.com.au/our-businesses) + ASX announcements |
| **Independent** | Not found in any corporate brand list |
| **Unknown** | Cannot be verified — flagged for manual review |

> ⚠️ Propel frequently acquires businesses that continue operating under local names. Always verify against their live website and recent ASX announcements.

---

## Quick Start

```bash
git clone https://github.com/YOUR_USERNAME/au-business-scraper
cd au-business-scraper
pip install -r requirements.txt
playwright install chromium

# Demo mode (no API needed)
python au_funeral_scraper.py --demo

# Scrape NSW from Yellow Pages
python au_funeral_scraper.py --source yellowpages --state NSW

# All states + Google enrichment via Outscraper
python au_funeral_scraper.py --all-states --outscraper-key YOUR_KEY
```

---

## Configuration

Copy `.env.example` to `.env` and fill in your keys:

```env
OUTSCRAPER_API_KEY=your_outscraper_key_here
OUTPUT_DIR=output
REQUEST_DELAY=1.5
MAX_RETRIES=3
```

---

## Cost Estimate

| Method | Coverage | Estimated Cost |
|--------|----------|---------------|
| Outscraper (recommended) | Full Australia ~3,000 businesses | AUD $50–70 |
| Google Places API direct | Full Australia | USD $150–200 |
| Yellow Pages only | Full Australia | Free (rate-limited) |

---

## Tech Stack

- **Python 3.11+**
- `requests` + `BeautifulSoup4` — HTTP + HTML parsing
- `playwright` — JavaScript-heavy sites (Yellow Pages pagination)
- `pandas` + `openpyxl` — Excel output
- `outscraper` — Google Places API wrapper

---

## Output Structure

```
output/
└── 20260403_131800/
    ├── AU_Funeral_Businesses_MASTER.csv   ← all states combined
    ├── AU_Funeral_NSW.csv
    ├── AU_Funeral_VIC.csv
    ├── AU_Funeral_QLD.csv
    ├── AU_Funeral_WA.csv
    ├── AU_Funeral_SA.csv
    ├── AU_Funeral_TAS.csv
    ├── AU_Funeral_ACT.csv
    ├── AU_Funeral_NT.csv
    └── AU_Funeral_Businesses.xlsx         ← all states + per-state sheets
```

---

## Quality Standards

- No duplicates — deduplicated by ABN (primary) or google_place_id
- Minimum 80% field completion for core required fields
- Google data populated for minimum 70% of records
- Ownership classification verified against published brand lists — not guessed
- All permanently closed businesses flagged via `google_business_status`

---

## License

MIT
