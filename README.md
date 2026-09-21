# Automated Flight Deal Finder & Price Intelligence System

A Python project for automated round-trip flight monitoring, historical fare analysis, price intelligence, and deal alerts.

The system collects flight observations through browser automation, stores historical results, evaluates price/quality trade-offs, and uses historical data to support price forecasting and low-fare detection.

## Features

- Automated round-trip flight searches with Selenium
- Configurable departure/return date windows
- Route-specific historical observations
- Price and itinerary-duration filtering
- Quality/price scoring
- Historical fare analysis
- Linear-regression-based price forecasting
- Low-price detection using accumulated observations
- Excel reports and summary sheets
- Price charts
- Optional Pushover notifications
- Configurable headless browser mode
- Desktop search UI with airport lookup and date selection
- Connectivity, CAPTCHA, stale-element, and browser-session handling

## Architecture

```text
Search Configuration / UI
          ↓
Selenium Flight Collection
          ↓
Filtering & Normalization
          ↓
Historical Observations
          ↓
Price Analysis / Forecasting
          ↓
Excel + Charts + Deal Alerts
```

## Tech Stack

- Python
- Selenium
- NumPy
- scikit-learn
- Matplotlib
- OpenPyXL
- Tkinter / tkcalendar
- Pushover API

## Project Structure

```text
automated-flight-deal-finder/
├── flight_deal_finder.py
├── flight_utils.py
├── notifications.py
├── search_ui.py
├── requirements.txt
├── .env.example
├── .gitignore
└── README.md
```

## Setup

Create and activate a virtual environment:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

For optional Pushover notifications, configure environment variables using `.env.example` as a reference. Never commit real API credentials.

Run the main monitoring script:

```bash
python flight_deal_finder.py
```

## Notes

The project relies on browser automation and external website structure, so selectors and scraping behavior may require maintenance when third-party interfaces change. Automated access should always respect the relevant website's terms and applicable usage restrictions.

The forecasting component is intended as a decision-support experiment based on collected historical observations; it does not guarantee future airfare movements.

## Security

Credentials are not included in this public version. Notification credentials are loaded from environment variables.

## Disclaimer

This repository is an independent educational and portfolio project. It is not affiliated with or endorsed by Kayak or any airline, booking platform, or travel provider.
