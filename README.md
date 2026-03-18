# 🍽️ Google Maps Restaurant Scraper

A Python-based web scraping toolkit that automatically collects restaurant data from Google Maps and exports it to a professionally formatted Excel workbook.

Two scripts are included:

| Script | Purpose |
|---|---|
| `google_maps_scraper.py` | General-purpose scraper — accepts any search query |
| `dhaka_zones_scraper.py` | Dhaka-specific scraper — hardcoded for Banani, Baridhara, and Gulshan |

---

## 📋 Table of Contents

- [Features](#-features)
- [Requirements](#-requirements)
- [Installation](#-installation)
- [Usage](#-usage)
  - [General Scraper](#1-general-scraper)
  - [Dhaka Zones Scraper](#2-dhaka-zones-scraper)
- [Configuration](#️-configuration)
- [Output](#-output)
- [Data Fields](#-data-fields)
- [Project Structure](#-project-structure)
- [Troubleshooting](#-troubleshooting)
- [Limitations](#-limitations)
- [License](#-license)

---

## ✨ Features

- 🔍 **Automated Google Maps search** — types your query and navigates results just like a real user
- 📜 **Auto-scrolling** — scrolls the results panel until the target count is reached or the list ends
- 🏪 **Detailed extraction** — captures name, address, phone, website, email, rating, reviews, cuisine, price level, hours, and Plus Code
- 📊 **Excel export** — outputs a formatted `.xlsx` file with colour-coded headers, alternating row fills, frozen panes, and auto-filters
- 🗂️ **Multi-sheet workbook** *(Dhaka script)* — one sheet per zone plus a combined "All Restaurants" sheet and a Summary tab
- 🛡️ **Graceful error handling** — missing fields default to `"N/A"` without crashing the script
- ⚙️ **Easy configuration** — all key settings live in a single `CONFIG` dictionary at the top of each file

---

## 📦 Requirements

- **Python** 3.9 or higher
- **Google Chrome** installed on your machine
- The following Python packages:

```
selenium
webdriver-manager
openpyxl
```

> `webdriver-manager` automatically downloads the correct ChromeDriver version for your installed Chrome — no manual setup needed.

---

## 🔧 Installation

**Step 1 — Clone or download this repository**

```bash
git clone https://github.com/your-username/gmaps-restaurant-scraper.git
cd gmaps-restaurant-scraper
```

**Step 2 — (Recommended) Create a virtual environment**

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# macOS / Linux
source venv/bin/activate
```

**Step 3 — Install dependencies**

```bash
pip install selenium webdriver-manager openpyxl
```

---

## 🚀 Usage

### 1. General Scraper

**File:** `google_maps_scraper.py`

Use this when you want to search for restaurants in **any city or location**.

Open the file and edit the `CONFIG` block near the top:

```python
CONFIG = {
    "search_query": "restaurants in New York",  # ← change this
    "max_results":  50,
    "output_file":  "restaurants_data.xlsx",
    ...
}
```

Then run:

```bash
python google_maps_scraper.py
```

---

### 2. Dhaka Zones Scraper

**File:** `dhaka_zones_scraper.py`

Pre-configured to scrape **Banani**, **Baridhara**, and **Gulshan** in Dhaka. No query editing required.

```bash
python dhaka_zones_scraper.py
```

The script will run three separate searches automatically and merge everything into one workbook.

---

## ⚙️ Configuration

Both scripts share the same `CONFIG` dictionary pattern. Below is a full reference:

| Key | Default | Description |
|---|---|---|
| `search_query` | `"restaurants in New York"` | *(General script only)* The Google Maps search string |
| `max_results` / `max_results_per_zone` | `50` / `40` | Maximum listings to collect. Set to `None` to collect everything |
| `output_file` | `"restaurants_data.xlsx"` / `"dhaka_restaurants.xlsx"` | Name of the output Excel file |
| `scroll_pause` | `2` | Seconds to pause between scroll attempts. Increase on slow connections |
| `detail_load_wait` | `3` | Seconds to wait for each detail page to fully load |
| `headless` | `True` | `True` = invisible browser. Set `False` to watch the browser in real time |

**Example — scrape 100 restaurants with the browser visible:**

```python
CONFIG = {
    "search_query": "restaurants in London",
    "max_results":  100,
    "output_file":  "london_restaurants.xlsx",
    "scroll_pause": 2.5,
    "detail_load_wait": 4,
    "headless": False,   # watch the browser
}
```

---

## 📁 Output

### General Scraper — `restaurants_data.xlsx`

| Sheet | Contents |
|---|---|
| `Restaurants` | All scraped records |
| `Summary` | Quick stats — total count, coverage per field |

### Dhaka Zones Scraper — `dhaka_restaurants.xlsx`

| Sheet | Accent Colour | Contents |
|---|---|---|
| `All Restaurants` | Dark slate | Every record from all three zones |
| `Banani` | Deep blue | Banani zone only |
| `Baridhara` | Deep green | Baridhara zone only |
| `Gulshan` | Deep purple | Gulshan zone only |
| `Summary` | — | Count per zone + field coverage stats |

Each data sheet includes:
- 📌 **Frozen header row** — stays visible while scrolling
- 🔽 **Auto-filter dropdowns** — filter by any column instantly
- 🎨 **Alternating row colours** — easier to read at a glance
- 📐 **Pre-set column widths** — no manual resizing needed

---

## 🗃️ Data Fields

Every scraped restaurant record contains the following fields:

| Field | Description | Notes |
|---|---|---|
| `Zone` | Zone label (Banani / Baridhara / Gulshan) | Dhaka script only |
| `Name` | Restaurant name | — |
| `Address` | Full street address | — |
| `Phone` | Phone number | — |
| `Website` | Official website URL | — |
| `Email` | Email address | Only available if a `mailto:` link appears on the Maps page |
| `Rating` | Star rating (e.g. `4.3`) | — |
| `Reviews` | Total number of reviews | — |
| `Category` | Primary category (e.g. `Restaurant`) | — |
| `Cuisine` | Cuisine or food type (e.g. `Italian, Pizza`) | — |
| `Price Level` | Price indicator (e.g. `$$`) | — |
| `Hours` | Opening hours | — |
| `Plus Code` | Google Plus Code location | — |
| `Google Maps URL` | Direct link to the Maps listing | — |
| `Scraped At` | Timestamp of when the record was collected | — |

> Fields not available for a given restaurant are filled with `"N/A"` — the script never crashes on missing data.

---

## 🗂️ Project Structure

```
gmaps-restaurant-scraper/
│
├── google_maps_scraper.py    # General-purpose scraper
├── dhaka_zones_scraper.py    # Dhaka zones scraper (Banani, Baridhara, Gulshan)
├── README.md                 # This file
│
└── output/                   # Generated after running the scripts
    ├── restaurants_data.xlsx
    └── dhaka_restaurants.xlsx
```

---

## 🛠️ Troubleshooting

**Chrome version mismatch**
> `webdriver-manager` handles this automatically. If you see a driver error, try upgrading: `pip install --upgrade webdriver-manager`

**No data collected / empty Excel**
> Set `"headless": False` in CONFIG to watch the browser. Google Maps may have changed its layout or is showing a CAPTCHA. Try running again after a few minutes.

**Script stops early**
> Google Maps may have detected automation. Try increasing `scroll_pause` to `3` or `4` seconds and setting `max_results` to a lower number.

**Fields showing "N/A" unexpectedly**
> Google Maps doesn't guarantee all fields are present for every listing. Some restaurants simply don't have a website, phone number, or hours listed. This is expected behaviour.

**`ModuleNotFoundError`**
> Make sure you've installed all dependencies: `pip install selenium webdriver-manager openpyxl`

---

## ⚠️ Limitations

- **Email addresses** are rarely exposed directly on Google Maps pages. The script captures them only when a `mailto:` link is explicitly present on the listing.
- **Google Maps layout** changes periodically. If the script stops extracting certain fields correctly, the CSS selectors inside `parse_detail_page()` may need updating. Run with `headless: False` to inspect the live page.
- **Rate limiting** — scraping large numbers of listings quickly may trigger Google's bot detection. Use reasonable `max_results` values and consider adding extra delay via `scroll_pause` and `detail_load_wait`.
- This tool is intended for **personal research and analysis** only. Please review [Google Maps Terms of Service](https://maps.google.com/help/terms_maps/) before large-scale use.

---

## 📄 License

This project is released for personal and educational use. Attribution appreciated but not required.
