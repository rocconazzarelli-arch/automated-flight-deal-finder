"""
KAYAKBOT_LOOP_ADVANCED.py

Roundtrip advanced version:
- output separati per rotta
- riuso file Excel/PNG esistenti
- salvataggio di tutte le osservazioni utili
- filtri (durata max, diretti opzionali)
- score qualità/prezzo
- foglio Summary Excel
- daily summary via Pushover
- grafico PNG più leggibile
- gestione robusta di StaleElementReferenceException
- finestra di ricerca date configurabile
"""

# ============================================================
# 1) IMPORT
# ============================================================

import os
import io
import re
import time
import random
import datetime
import mimetypes
import http.client
import socket
import tempfile
from pathlib import Path
from typing import Optional, Tuple, List, Dict
from sklearn.linear_model import LinearRegression
import numpy as np

import matplotlib.pyplot as plt

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException,
    WebDriverException,
    InvalidSessionIdException,
)

from openpyxl import Workbook, load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.label import DataLabelList

import notifications as invia


# ============================================================
# 2) CONFIG
# ============================================================

PUSHOVER_TOKEN = os.getenv("PUSHOVER_TOKEN", "")
PUSHOVER_USER = os.getenv("PUSHOVER_USER", "")
BOT_DISPLAY_NAME = "BOTFLIGHTS RT"

EXCEL_SHEET_NAME = "Flights"
CHART_TITLE = "Prezzo voli roundtrip"

MAX_TOTAL_DURATION_MIN = None   # es. 900 = max 15h totali A/R
DIRECT_ONLY = False
SAVE_ALL_OBSERVATIONS = True

PRICE_WEIGHT = 1.0
DURATION_WEIGHT = 0.18

DAILY_SUMMARY_HOUR = 21

# Finestra di ricerca intorno alle date scelte
OUTBOUND_DELTA_DAYS = 5
RETURN_DELTA_DAYS = 5

# Gap minimo tra andata e ritorno per evitare date incoerenti
MIN_STAY_DAYS = 1

CAPTCHA_WAIT_TIMEOUT_SEC = 45
CAPTCHA_POLL_INTERVAL_SEC = 5
LOW_PRICE_MIN_HISTORY_POINTS = 12
LOW_PRICE_ALERT_COOLDOWN_CHECKS = 250
CONNECTIVITY_CHECK_TIMEOUT_SEC = 3
CONNECTIVITY_WAIT_POLL_SEC = 10
LOOP_SLEEP_MIN_SEC = 45
LOOP_SLEEP_MAX_SEC = 120
CAPTCHA_ROUTE_COOLDOWN_SEC = 15 * 60
ERROR_PAUSE_SEC = 5 * 60
DRIVER_RESTART_PAUSE_SEC = 60
STARTUP_NOTIFICATION_SENT = False
CONNECTION_STATUS = "online"
LAST_CONNECTION_LOSS_AT: Optional[datetime.datetime] = None
LAST_CONNECTION_RESTORE_AT: Optional[datetime.datetime] = None
HEADLESS_MODE = os.getenv("BOTFLIGHTS_HEADLESS", "auto").lower()

# Tratte configurate manualmente per esecuzione da terminale/server.
# Formato: partenza IATA, destinazione IATA, data andata, data ritorno, prezzo soglia.
MANUAL_ROUTES = [
    ("GRU", "MAO", "2026-07-04", "2026-07-13", 250),
    ("GRU", "SSA", "2026-06-18", "2026-06-21", 200),
    ("GRU", "GUA", "2026-07-10", "2026-07-22", 350),
]


# ============================================================
# 3) PUSHOVER
# ============================================================

def _multipart_encode(fields: dict, files: dict) -> tuple[bytes, str]:
    boundary = "----BOTFLIGHTS_BOUNDARY_" + str(int(time.time() * 1000))
    crlf = "\r\n"
    body = io.BytesIO()

    for name, value in fields.items():
        body.write(f"--{boundary}{crlf}".encode())
        body.write(f'Content-Disposition: form-data; name="{name}"{crlf}{crlf}'.encode())
        body.write(str(value).encode())
        body.write(crlf.encode())

    for name, (filename, content, content_type) in files.items():
        body.write(f"--{boundary}{crlf}".encode())
        body.write(f'Content-Disposition: form-data; name="{name}"; filename="{filename}"{crlf}'.encode())
        body.write(f"Content-Type: {content_type}{crlf}{crlf}".encode())
        body.write(content)
        body.write(crlf.encode())

    body.write(f"--{boundary}--{crlf}".encode())
    return body.getvalue(), boundary

def format_eur(value) -> str:
    try:
        v = int(round(float(value)))
        return f"{v}€"
    except Exception:
        return "N/A"

def _percentile(sorted_values, q: float):
    if not sorted_values:
        return None
    if len(sorted_values) == 1:
        return float(sorted_values[0])

    q = max(0.0, min(1.0, float(q)))
    pos = (len(sorted_values) - 1) * q
    low = int(pos)
    high = min(low + 1, len(sorted_values) - 1)
    frac = pos - low
    return float(sorted_values[low] + (sorted_values[high] - sorted_values[low]) * frac)


def build_ml_dataset(rows):
    """
    Costruisce un dataset pulito per la predizione:
    - usa i minimi di round per evitare duplicati interni al singolo scrape
    - scarta prezzi nulli / non numerici
    - scarta prezzi assurdi
    - scarta date troppo vicine/passate o troppo lontane
    - raggruppa per days_to_flight usando il prezzo mediano
    """
    grouped = {}
    samples = []

    now = datetime.datetime.now()

    for r in extract_round_min_points(rows):
        price = r.get("price")
        ga = r.get("ga")

        if price is None or not ga:
            continue

        try:
            price = float(price)
            dt = datetime.datetime.strptime(str(ga), "%Y-%m-%d")
            days_to_flight = (dt - now).days
        except Exception:
            continue

        # filtri sensati
        if not (1 <= days_to_flight <= 365):
            continue

        if not (30 <= price <= 5000):
            continue

        grouped.setdefault(days_to_flight, []).append(price)
        samples.append({
            "days_to_flight": int(days_to_flight),
            "price": float(price),
            "saved_at": r.get("saved_at"),
            "ga": str(ga),
        })

    if not grouped:
        return np.array([]), np.array([]), []

    X = []
    y = []

    for days in sorted(grouped.keys()):
        prices = sorted(grouped[days])
        n = len(prices)

        # mediana robusta
        if n % 2 == 1:
            median_price = prices[n // 2]
        else:
            median_price = (prices[n // 2 - 1] + prices[n // 2]) / 2

        X.append([days])
        y.append(median_price)

    samples.sort(key=lambda item: item.get("saved_at") or datetime.datetime.min)

    return np.array(X, dtype=float), np.array(y, dtype=float), samples

def predict_price(rows):
    X, y, samples = build_ml_dataset(rows)

    # pochi dati = niente forecast
    if len(X) < 8:
        return None

    # usa solo gli ultimi punti più rilevanti, se sono tanti
    if len(X) > 120:
        X = X[-120:]
        y = y[-120:]

    model = LinearRegression()
    sample_weight = np.linspace(0.7, 1.3, len(X))
    model.fit(X, y, sample_weight=sample_weight)

    future_days = np.array([[7], [14], [30]], dtype=float)
    linear_preds = model.predict(future_days)

    hist_min = float(np.min(y))
    hist_max = float(np.max(y))
    hist_avg = float(np.mean(y))
    sorted_prices = sorted(float(v) for v in y)
    recent_prices = [float(item["price"]) for item in samples[-12:]] if samples else []
    recent_anchor = float(np.median(recent_prices)) if recent_prices else float(np.median(y))

    current_sample = samples[-1] if samples else None
    current_price = float(current_sample["price"]) if current_sample else float(recent_anchor)
    current_days = int(current_sample["days_to_flight"]) if current_sample else int(X[-1][0])

    blended_preds = []
    for idx, target_days in enumerate((7, 14, 30)):
        trend_component = float(linear_preds[idx])
        distance = abs(current_days - target_days)
        blend_ratio = min(0.75, distance / 30.0)
        blended = recent_anchor * (1.0 - blend_ratio) + trend_component * blend_ratio
        blended_preds.append(blended)

    # clamp forte: niente numeri assurdi
    p10 = _percentile(sorted_prices, 0.10) or hist_min
    p90 = _percentile(sorted_prices, 0.90) or hist_max
    lower_bound = max(30.0, hist_min * 0.80, p10 * 0.85)
    upper_bound = min(5000.0, hist_max * 1.20, p90 * 1.15)

    preds = np.clip(np.array(blended_preds, dtype=float), lower_bound, upper_bound)
    trend_delta_30 = preds[2] - current_price
    recent_volatility = float(np.std(recent_prices)) if len(recent_prices) >= 2 else 0.0
    confidence = "Alta" if len(samples) >= 25 and recent_volatility <= 35 else (
        "Media" if len(samples) >= 12 and recent_volatility <= 70 else "Bassa"
    )

    return {
        "7d": int(round(float(preds[0]))),
        "14d": int(round(float(preds[1]))),
        "30d": int(round(float(preds[2]))),
        "avg": int(round(hist_avg)),
        "min": int(round(hist_min)),
        "max": int(round(hist_max)),
        "points": int(len(y)),
        "current_est": int(round(current_price)),
        "current_days_to_flight": int(current_days),
        "recent_anchor": int(round(recent_anchor)),
        "trend_30d": int(round(float(trend_delta_30))),
        "volatility": int(round(recent_volatility)),
        "confidence": confidence,
    }


def build_ml_prediction_message(route_key_name: str, prediction: dict) -> str:
    trend_30d = prediction["trend_30d"]

    if trend_30d <= -40:
        signal = "📉 Probabile discesa"
    elif trend_30d >= 40:
        signal = "📈 Probabile rialzo"
    else:
        signal = "➡️ Fascia stabile"

    return (
        f"🧠 ML Forecast {route_key_name}\n"
        f"Round tratta: {prediction.get('rounds', 'N/A')}\n"
        f"Controlli tratta: {prediction.get('checks', 'N/A')}\n"
        f"Prezzo attuale stimato: {format_eur(prediction['current_est'])}\n"
        f"Baseline recente: {format_eur(prediction['recent_anchor'])}\n"
        f"Giorni al volo attuali: {prediction['current_days_to_flight']}\n"
        f"Media storica: {format_eur(prediction['avg'])}\n"
        f"Range storico: {format_eur(prediction['min'])} - {format_eur(prediction['max'])}\n"
        f"Punti modello: {prediction['points']}\n"
        f"Volatilità recente: {format_eur(prediction['volatility'])}\n"
        f"Affidabilità: {prediction['confidence']}\n\n"
        f"Stima 7g: {format_eur(prediction['7d'])}\n"
        f"Stima 14g: {format_eur(prediction['14d'])}\n"
        f"Stima 30g: {format_eur(prediction['30d'])}\n\n"
        f"Segnale: {signal}"
    )

def pushover_send(
    message: str,
    attachment_path: Optional[str] = None,
    title: str = "BOTFLIGHTS",
    priority: int = 0,
    sound: Optional[str] = None,
) -> None:
    if not PUSHOVER_TOKEN or not PUSHOVER_USER:
        print("⚠️ Pushover non configurato: mancano PUSHOVER_TOKEN o PUSHOVER_USER.")
        return

    title = str(title or "Alert")
    if BOT_DISPLAY_NAME not in title:
        title = f"{BOT_DISPLAY_NAME} · {title}"

    message = str(message)
    if BOT_DISPLAY_NAME not in message.splitlines()[0:1]:
        message = f"{BOT_DISPLAY_NAME}\n{message}"

    fields = {
        "token": PUSHOVER_TOKEN,
        "user": PUSHOVER_USER,
        "message": message,
        "title": title,
        "priority": str(priority),
    }
    if sound:
        fields["sound"] = sound

    files = {}
    if attachment_path:
        p = Path(attachment_path)
        if p.exists() and p.is_file():
            ctype = mimetypes.guess_type(str(p))[0] or "application/octet-stream"
            files["attachment"] = (p.name, p.read_bytes(), ctype)

    body, boundary = _multipart_encode(fields, files)
    headers = {"Content-Type": f"multipart/form-data; boundary={boundary}"}

    conn = http.client.HTTPSConnection("api.pushover.net", 443, timeout=30)
    try:
        conn.request("POST", "/1/messages.json", body=body, headers=headers)
        resp = conn.getresponse()
        data = resp.read()
        print("Pushover:", resp.status, data[:200])
    finally:
        conn.close()


def notify_startup_once() -> None:
    global STARTUP_NOTIFICATION_SENT
    if STARTUP_NOTIFICATION_SENT:
        return

    STARTUP_NOTIFICATION_SENT = True
    try:
        pushover_send("🚀 KAYAK BOT (ROUNDTRIP ADVANCED) AVVIATO 🚀", title="BOTFLIGHTS RT")
    except Exception as exc:
        print(f"⚠️ Notifica avvio non inviata: {exc}")


# ============================================================
# 4) UTILS
# ============================================================

def to_iata(s: str) -> str:
    s = (s or "").strip()
    if "(" in s and ")" in s:
        return s.split("(")[-1].split(")")[0].strip().upper()

    parts = [p.strip() for p in s.split("-") if p.strip()]
    if len(parts) >= 2:
        candidate = parts[1].upper()
        if len(candidate) == 3 and candidate.isalpha():
            return candidate

    return parts[0].upper() if parts else s.upper()


def build_kayak_url(departure: str, destination: str, dep_date: str, ret_date: str) -> str:
    o = to_iata(departure)
    d = to_iata(destination)
    return f"https://www.kayak.it/flights/{o}-{d}/{dep_date}/{ret_date}?sort=bestflight_a"


def parse_price_to_int(raw: str) -> Optional[int]:
    if not raw:
        return None
    digits = "".join(ch for ch in raw if ch.isdigit())
    return int(digits) if digits else None


def shift_date(date_str: str, delta_days: int) -> str:
    base = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
    newd = base + datetime.timedelta(days=delta_days)
    return newd.strftime("%Y-%m-%d")


def clean_text(x: str) -> str:
    return (x or "").strip()


def weekday_short(date_iso: str) -> str:
    dt = datetime.datetime.strptime(date_iso, "%Y-%m-%d").date()
    return dt.strftime("%a")


def route_key(dep_iata: str, dst_iata: str) -> str:
    dep = (dep_iata or "").strip().upper()
    dst = (dst_iata or "").strip().upper()
    return f"{dep}-{dst}"


def parse_duration_minutes(text: str) -> Optional[int]:
    if not text:
        return None

    t = text.lower().strip().replace(" ", "")

    try:
        days = 0
        hours = 0
        minutes = 0

        m_days = re.search(r"(\d+)d", t)
        if m_days:
            days = int(m_days.group(1))

        m_hours = re.search(r"(\d+)h", t)
        if m_hours:
            hours = int(m_hours.group(1))

        m_minutes = re.search(r"(\d+)min", t)
        if m_minutes:
            minutes = int(m_minutes.group(1))

        total = days * 24 * 60 + hours * 60 + minutes
        return total if total > 0 else None
    except Exception:
        return None


def compute_quality_score(price: Optional[int], total_duration_min: Optional[int]) -> Optional[float]:
    if price is None:
        return None
    dur = total_duration_min if total_duration_min is not None else 0
    return PRICE_WEIGHT * float(price) + DURATION_WEIGHT * float(dur)


def is_direct_flight(duration_text: str, company_text: str = "") -> bool:
    blob = f"{duration_text or ''}".lower()

    positive_terms = [
        "nonstop", "non-stop", "diretto", "direct",
        "senza scali", "0 scali", "0 stop", "0 stops"
    ]
    if any(term in blob for term in positive_terms):
        return True

    negative_terms = [
        "1 stop", "2 stops", "3 stops",
        "1 scalo", "2 scali", "3 scali",
        "stop", "stops", "scalo", "scali"
    ]
    if any(term in blob for term in negative_terms):
        return False

    return True


def should_keep_flight(item: Dict) -> bool:
    total_duration = item.get("total_duration_min")

    if MAX_TOTAL_DURATION_MIN is not None and total_duration is not None and total_duration > MAX_TOTAL_DURATION_MIN:
        return False

    if DIRECT_ONLY and not item.get("is_direct", False):
        return False

    return True


def safe_float(x, default=None):
    try:
        return float(x)
    except Exception:
        return default


def safe_int(x, default=None):
    try:
        return int(float(x))
    except Exception:
        return default


def parse_saved_at_dt(x: str):
    try:
        return datetime.datetime.strptime(str(x), "%Y-%m-%d %H:%M:%S")
    except Exception:
        return None


def bool_from_any(x) -> bool:
    return str(x).strip().lower() in {"true", "1", "yes", "y"}


def extract_round_min_points(rows: List[dict]) -> List[dict]:
    """
    Restituisce un punto per ogni round (ROUND_ID), scegliendo
    il prezzo minimo del round. In caso di pari prezzo, sceglie
    la riga con score migliore.
    """
    grouped: Dict[str, dict] = {}

    for r in rows:
        round_id = str(r.get("ROUND_ID") or r.get("SAVED_AT") or "").strip()
        saved_at = parse_saved_at_dt(r.get("SAVED_AT"))
        price = safe_float(r.get("PT"))
        score = safe_float(r.get("SCORE"), float("inf"))

        if not round_id or saved_at is None or price is None:
            continue

        candidate = {
            "round_id": round_id,
            "saved_at": saved_at,
            "price": price,
            "ga": str(r.get("GA", "")),
            "gr": str(r.get("GR", "")),
            "ga_label": str(r.get("GA_LABEL", "")),
            "gr_label": str(r.get("GR_LABEL", "")),
            "com": str(r.get("COM", "")),
            "score": score,
            "is_direct": bool_from_any(r.get("IS_DIRECT")),
            "total_duration_min": safe_int(r.get("TOTAL_DURATION_MIN")),
        }

        current = grouped.get(round_id)
        if current is None:
            grouped[round_id] = candidate
        else:
            if candidate["price"] < current["price"]:
                grouped[round_id] = candidate
            elif candidate["price"] == current["price"] and candidate["score"] < current["score"]:
                grouped[round_id] = candidate

    return sorted(grouped.values(), key=lambda x: x["saved_at"])


# ============================================================
# 5) EXCEL LOADING / SUMMARY
# ============================================================

def load_existing_rows(excel_path: str) -> List[dict]:
    headers = [
        "SEARCHED_AT", "SAVED_AT", "ROUTE",
        "GA", "GA_LABEL", "GA_DOW",
        "GR", "GR_LABEL", "GR_DOW",
        "OPA", "OAA", "OPR", "OAR",
        "DA", "DR", "DA_MIN", "DR_MIN", "TOTAL_DURATION_MIN",
        "PT", "COM", "IS_DIRECT", "SCORE", "IS_NEW_MIN", "STATUS",
        "ROUND_ID", "ROUND_MIN_FOR_GA"
    ]

    p = Path(excel_path)
    if not p.exists():
        return []

    try:
        wb = load_workbook(excel_path, data_only=True)
        if EXCEL_SHEET_NAME not in wb.sheetnames:
            return []

        ws = wb[EXCEL_SHEET_NAME]
        rows = []
        for r in ws.iter_rows(min_row=2, values_only=True):
            if not r or not r[0]:
                continue
            row_dict = {h: (r[i] if i < len(r) else None) for i, h in enumerate(headers)}
            rows.append(row_dict)
        return rows
    except Exception:
        return []


def summarize_existing_data(rows: List[dict]) -> Dict:
    round_points = extract_round_min_points(rows)

    saved_prices = []
    direct_saved_count = 0
    scored = []

    for row in rows:
        pt = safe_float(row.get("PT"))
        if pt is not None:
            saved_prices.append(pt)

        if bool_from_any(row.get("IS_DIRECT")):
            direct_saved_count += 1

        score_val = safe_float(row.get("SCORE"))
        if score_val is not None:
            scored.append((score_val, row))

    round_prices = [p["price"] for p in round_points]
    direct_round_count = sum(1 for p in round_points if p["is_direct"])

    stats = {
        "saved_observations": len(rows),
        "round_count": len(round_points),

        "min_price_saved": min(saved_prices) if saved_prices else None,
        "max_price_saved": max(saved_prices) if saved_prices else None,
        "avg_price_saved": (sum(saved_prices) / len(saved_prices)) if saved_prices else None,

        "min_round_price": min(round_prices) if round_prices else None,
        "max_round_price": max(round_prices) if round_prices else None,
        "avg_round_price": (sum(round_prices) / len(round_prices)) if round_prices else None,
        "last_round_price": round_prices[-1] if round_prices else None,

        "direct_saved_count": direct_saved_count,
        "direct_round_count": direct_round_count,

        "best_score_row": min(scored, key=lambda x: x[0])[1] if scored else None,
        "round_points": round_points,
    }
    return stats


def detect_low_price_signal(history_rows: List[dict], candidate_price: Optional[float]) -> Optional[Dict]:
    if candidate_price is None:
        return None

    round_points = extract_round_min_points(history_rows)
    history_prices = sorted(float(p["price"]) for p in round_points if p.get("price") is not None)
    if len(history_prices) < LOW_PRICE_MIN_HISTORY_POINTS:
        return None

    hist_min = min(history_prices)
    hist_avg = sum(history_prices) / len(history_prices)
    p10 = _percentile(history_prices, 0.10)
    p25 = _percentile(history_prices, 0.25)
    if p10 is None or p25 is None:
        return None

    candidate_price = float(candidate_price)
    avg_discount = ((hist_avg - candidate_price) / hist_avg) if hist_avg > 0 else 0.0

    if candidate_price > p10:
        return None

    if avg_discount < 0.06 and candidate_price > hist_min:
        return None

    severity = "strong" if candidate_price <= min(hist_min, p10 * 0.97) else "medium"
    return {
        "candidate_price": int(round(candidate_price)),
        "hist_min": int(round(hist_min)),
        "hist_avg": int(round(hist_avg)),
        "p10": int(round(p10)),
        "p25": int(round(p25)),
        "history_points": len(history_prices),
        "avg_discount_pct": int(round(avg_discount * 100)),
        "severity": severity,
    }


def build_low_price_alert_message(route_key_name: str, dep: str, ret: str, signal: Dict) -> str:
    label = "🔥 Prezzo eccezionale" if signal["severity"] == "strong" else "💡 Prezzo molto interessante"
    return (
        f"{label} {route_key_name}\n"
        f"Andata: {dep}\n"
        f"Ritorno: {ret}\n"
        f"Prezzo round: {format_eur(signal['candidate_price'])}\n"
        f"Min storico round: {format_eur(signal['hist_min'])}\n"
        f"Media storica round: {format_eur(signal['hist_avg'])}\n"
        f"P10 storico: {format_eur(signal['p10'])}\n"
        f"P25 storico: {format_eur(signal['p25'])}\n"
        f"Vantaggio vs media: {signal['avg_discount_pct']}%\n"
        f"Round storici usati: {signal['history_points']}"
    )


# ============================================================
# 6) SELENIUM HELPERS
# ============================================================

def accept_cookies_best_effort(driver):
    try:
        for txt in ("Accetta", "Accetta tutto", "Accept", "Accept all"):
            btns = driver.find_elements(By.XPATH, f"//button[contains(., '{txt}')]")
            if btns:
                btns[0].click()
                return
    except Exception:
        pass


def _has_visible_elements(driver, by, selector: str) -> bool:
    try:
        elements = driver.find_elements(by, selector)
    except Exception:
        return False

    for el in elements:
        try:
            if el.is_displayed():
                return True
        except Exception:
            continue
    return False


def is_captcha_page(driver) -> bool:
    """
    Rileva solo challenge interattive da risolvere manualmente:
    - Arkose/FunCaptcha
    - GeeTest / slider captcha
    - hCaptcha/reCAPTCHA challenge frame visibile

    Evita match generici su testo pagina per ridurre i falsi positivi.
    """
    iframe_checks = [
        "iframe[src*='arkoselabs']",
        "iframe[src*='funcaptcha']",
        "iframe[src*='geetest']",
        "iframe[src*='captcha-delivery']",
        "iframe[src*='bframe'][title*='challenge']",
        "iframe[src*='hcaptcha.com'][title*='challenge']",
        "iframe[title*='hCaptcha challenge']",
        "iframe[title*='recaptcha challenge']",
    ]

    widget_checks = [
        "[class*='arkose']",
        "[id*='arkose']",
        "[class*='funcaptcha']",
        "[id*='funcaptcha']",
        "[class*='geetest']",
        "[id*='geetest']",
        "[class*='slider']",
        "[id*='slider']",
        "[class*='captcha-slider']",
        "[class*='puzzle']",
        "[id*='puzzle']",
        "canvas[aria-label*='captcha']",
        "canvas[class*='captcha']",
    ]

    xpath_checks = [
        "//*[contains(@class, 'geetest') and (contains(@class, 'panel') or contains(@class, 'holder') or contains(@class, 'wrap'))]",
        "//*[contains(@class, 'slider') and (contains(@class, 'captcha') or contains(@class, 'verify'))]",
        "//*[contains(@class, 'puzzle') and (self::div or self::section or self::canvas)]",
        "//*[contains(@id, 'captcha') and (contains(@class, 'slider') or contains(@class, 'puzzle'))]",
    ]

    for selector in iframe_checks:
        if _has_visible_elements(driver, By.CSS_SELECTOR, selector):
            return True

    for selector in widget_checks:
        if _has_visible_elements(driver, By.CSS_SELECTOR, selector):
            return True

    for selector in xpath_checks:
        if _has_visible_elements(driver, By.XPATH, selector):
            return True

    return False


def wait_until_captcha_solved(driver, timeout: int = CAPTCHA_WAIT_TIMEOUT_SEC) -> bool:
    started = time.time()
    while time.time() - started < timeout:
        if not is_captcha_page(driver):
            return True
        time.sleep(CAPTCHA_POLL_INTERVAL_SEC)
    return False


def save_captcha_screenshot(driver, route_name: str) -> Optional[str]:
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_route = re.sub(r"[^A-Za-z0-9_-]+", "_", route_name or "route")
    screenshot_path = Path("/tmp") / f"captcha_{safe_route}_{ts}.png"

    try:
        ok = driver.save_screenshot(str(screenshot_path))
        if ok and screenshot_path.exists():
            return str(screenshot_path)
    except Exception:
        pass
    return None


def save_timeout_debug_snapshot(driver, route_name: str, dep: str, ret: str) -> None:
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_route = re.sub(r"[^A-Za-z0-9_-]+", "_", route_name or "route")
    base = Path("/tmp") / f"kayak_timeout_{safe_route}_{dep}_{ret}_{ts}"
    png_path = base.with_suffix(".png")
    html_path = base.with_suffix(".html")

    try:
        driver.save_screenshot(str(png_path))
    except Exception as exc:
        print(f"⚠️ Screenshot timeout non salvato: {exc}")

    try:
        html_path.write_text(driver.page_source or "", encoding="utf-8", errors="ignore")
    except Exception as exc:
        print(f"⚠️ HTML timeout non salvato: {exc}")

    try:
        body_text = driver.find_element(By.TAG_NAME, "body").text
    except Exception:
        body_text = ""

    title = ""
    current_url = ""
    try:
        title = driver.title
        current_url = driver.current_url
    except Exception:
        pass

    snippet = re.sub(r"\s+", " ", body_text).strip()[:500]
    print(f"🧪 Debug timeout salvato: {png_path} | {html_path}")
    print(f"🧪 Page title: {title}")
    print(f"🧪 Current URL: {current_url}")
    if snippet:
        print(f"🧪 Body snippet: {snippet}")


def handle_captcha_if_present(driver, route_name: str) -> bool:
    if not is_captcha_page(driver):
        return False

    print(f"⚠️ CAPTCHA puzzle rilevato su {route_name}. Risolvilo nel browser: il bot resta in attesa.")
    screenshot_path = save_captcha_screenshot(driver, route_name)
    try:
        pushover_send(
            f"⚠️ CAPTCHA puzzle rilevato su {route_name}. Risoluzione in corso... il bot poi riparte automaticamente.",
            attachment_path=screenshot_path,
            title=f"Captcha {route_name}",
        )
    except Exception:
        pass

    solved = wait_until_captcha_solved(driver, timeout=CAPTCHA_WAIT_TIMEOUT_SEC)
    if solved:
        print(f"✅ CAPTCHA risolto su {route_name}. Riprendo automaticamente il controllo della pagina.")
        try:
            pushover_send(
                f"✅ CAPTCHA risolto su {route_name}. Il bot ha ripreso automaticamente il controllo della pagina.",
                title=f"Captcha OK {route_name}",
            )
        except Exception:
            pass
    else:
        print(f"⏱️ CAPTCHA  {CAPTCHA_WAIT_TIMEOUT_SEC}s su {route_name}.")
    return True


def collect_valid_items_best_effort(driver) -> List[Dict]:
    try:
        items = collect_results_roundtrip(driver)
    except Exception:
        return []
    return [it for it in items if it.get("prezzo_int") is not None]


def log_round_status(route_name: str, dep: str, ret: str, status: str, detail: str = "") -> None:
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    msg = f"[{ts}] [{route_name}] [{status}] {dep} -> {ret}"
    if detail:
        msg += f" | {detail}"
    print(msg)


def has_internet_connection() -> bool:
    targets = [
        ("1.1.1.1", 53),
        ("8.8.8.8", 53),
        ("www.kayak.it", 443),
    ]

    for host, port in targets:
        try:
            with socket.create_connection((host, port), timeout=CONNECTIVITY_CHECK_TIMEOUT_SEC):
                return True
        except OSError:
            continue
    return False


def _connection_push_section(title: str, rows: List[Tuple[str, str]]) -> List[str]:
    lines = [f"▸ {title}"]
    for label, value in rows:
        lines.append(f"  {label:<8} {value}")
    return lines


def notify_connection_lost_if_needed(route_name: str, dep: str, ret: str) -> None:
    global CONNECTION_STATUS, LAST_CONNECTION_LOSS_AT

    if CONNECTION_STATUS == "offline":
        return

    CONNECTION_STATUS = "offline"
    LAST_CONNECTION_LOSS_AT = datetime.datetime.now()

    message = "\n".join([
        "🔴 BOTFLIGHTS · Connection lost",
        "━━━━━━━━━━━━━━━━",
        *_connection_push_section("Status", [
            ("State", "Offline"),
            ("Time", LAST_CONNECTION_LOSS_AT.strftime("%Y-%m-%d %H:%M:%S")),
            ("Route", route_name),
            ("Dates", f"{dep} → {ret}"),
            ("Action", "Search paused"),
        ]),
    ])

    try:
        pushover_send(message, title="BOTFLIGHTS offline", priority=1, sound="siren")
    except Exception as exc:
        print(f"⚠️ Notifica perdita connessione non inviata: {exc}")


def notify_connection_restored_if_needed(route_name: str, dep: str, ret: str) -> None:
    global CONNECTION_STATUS, LAST_CONNECTION_LOSS_AT, LAST_CONNECTION_RESTORE_AT

    if CONNECTION_STATUS != "offline":
        return

    LAST_CONNECTION_RESTORE_AT = datetime.datetime.now()
    offline_for = "N/A"
    if LAST_CONNECTION_LOSS_AT is not None:
        offline_for = str(LAST_CONNECTION_RESTORE_AT - LAST_CONNECTION_LOSS_AT).split(".")[0]

    CONNECTION_STATUS = "online"

    message = "\n".join([
        "🟢 BOTFLIGHTS · Connection restored",
        "━━━━━━━━━━━━━━━━",
        *_connection_push_section("Status", [
            ("State", "Online"),
            ("Time", LAST_CONNECTION_RESTORE_AT.strftime("%Y-%m-%d %H:%M:%S")),
            ("Offline", offline_for),
            ("Route", route_name),
            ("Dates", f"{dep} → {ret}"),
            ("Action", "Search resumed"),
        ]),
    ])

    try:
        pushover_send(message, title="BOTFLIGHTS online", priority=0, sound="magic")
    except Exception as exc:
        print(f"⚠️ Notifica ripristino connessione non inviata: {exc}")


def wait_until_connection_restored(route_name: str, dep: str, ret: str) -> None:
    print("🌐 Connessione assente. Il bot entra in attesa finché la rete non torna disponibile.")
    log_round_status(route_name, dep, ret, "NO_CONNECTION", "connessione assente, attendo ripristino rete")
    notify_connection_lost_if_needed(route_name, dep, ret)

    while True:
        if has_internet_connection():
            print("✅ Connessione ripristinata. Riprendo automaticamente il bot.")
            log_round_status(route_name, dep, ret, "CONNECTION_OK", "connessione ripristinata, riprendo il loop")
            notify_connection_restored_if_needed(route_name, dep, ret)
            return

        time.sleep(CONNECTIVITY_WAIT_POLL_SEC)


def is_connectivity_error(exc: Exception) -> bool:
    text = str(exc).lower()
    patterns = [
        "err_internet_disconnected",
        "err_name_not_resolved",
        "err_network_changed",
        "err_connection_timed_out",
        "err_timed_out",
        "dns",
        "name resolution",
        "connection refused",
        "connection reset",
        "net::",
        "disconnected",
    ]
    return any(p in text for p in patterns)


def put_route_in_captcha_cooldown(route_data: Dict, route_name: str, reason: str) -> None:
    until = time.monotonic() + CAPTCHA_ROUTE_COOLDOWN_SEC
    route_data[route_name]["cooldown_until"] = until
    route_data[route_name]["captcha_count"] = int(route_data[route_name].get("captcha_count", 0)) + 1
    minutes = int(round(CAPTCHA_ROUTE_COOLDOWN_SEC / 60))
    print(f"🧊 {route_name} in cooldown {minutes} min dopo captcha ({reason}).")


def wait_for_results(driver, wait: WebDriverWait):
    candidates = [
        (By.CSS_SELECTOR, "div.nrc6"),
        (By.CSS_SELECTOR, "div.Flights-Results-FlightResultItem"),
        (By.CSS_SELECTOR, "div.f8F1-price-text"),
        (By.CSS_SELECTOR, "div.e2GB-price-text"),
    ]

    last = None
    for how, sel in candidates:
        try:
            wait.until(EC.presence_of_all_elements_located((how, sel)))
            elems = driver.find_elements(how, sel)
            if elems:
                return
        except Exception as e:
            last = e

    raise TimeoutException("Risultati non trovati / pagina bloccata.") from last


def _first_text(el, selectors: List[Tuple[str, str]]) -> str:
    for how, sel in selectors:
        try:
            found = el.find_elements(By.CSS_SELECTOR, sel) if how == "css" else el.find_elements(By.XPATH, sel)
            for f in found:
                try:
                    t = clean_text(f.text)
                    if t:
                        return t
                except StaleElementReferenceException:
                    continue
        except StaleElementReferenceException:
            continue
        except Exception:
            continue
    return ""


def _all_texts(el, selectors: List[Tuple[str, str]]) -> List[str]:
    out: List[str] = []
    for how, sel in selectors:
        try:
            found = el.find_elements(By.CSS_SELECTOR, sel) if how == "css" else el.find_elements(By.XPATH, sel)
            for f in found:
                try:
                    t = clean_text(f.text)
                    if t:
                        out.append(t)
                except StaleElementReferenceException:
                    continue
        except StaleElementReferenceException:
            continue
        except Exception:
            continue
    return out


# ============================================================
# 7) PARSING RESULTS ROUNDTRIP
# ============================================================

def collect_results_roundtrip(driver):
    results = []

    card_selectors = [
        ("css", "div.nrc6"),
        ("css", "div.Flights-Results-FlightResultItem"),
        ("xpath", "//div[contains(@class,'nrc6')]"),
    ]

    cards = []
    for how, sel in card_selectors:
        try:
            cards = driver.find_elements(By.CSS_SELECTOR, sel) if how == "css" else driver.find_elements(By.XPATH, sel)
            if cards:
                break
        except Exception:
            continue

    if not cards:
        price_selectors = [
            ("css", "div.f8F1-price-text"),
            ("css", "div.e2GB-price-text"),
            ("xpath", "//*[contains(@class,'price-text')]"),
        ]
        prezzi = _all_texts(driver, price_selectors)
        for raw in prezzi:
            p = parse_price_to_int(raw)
            item = {
                "prezzo_raw": raw,
                "prezzo_int": p,
                "opa": "", "oaa": "", "opr": "", "oar": "",
                "da": "", "dr": "",
                "da_min": None, "dr_min": None, "total_duration_min": None,
                "com": "",
                "is_direct": True,
                "score": compute_quality_score(p, None),
            }
            if should_keep_flight(item):
                results.append(item)
        return results

    for idx in range(len(cards)):
        try:
            fresh_cards = []
            for how, sel in card_selectors:
                try:
                    fresh_cards = driver.find_elements(By.CSS_SELECTOR, sel) if how == "css" else driver.find_elements(By.XPATH, sel)
                    if fresh_cards:
                        break
                except Exception:
                    continue

            if idx >= len(fresh_cards):
                continue

            c = fresh_cards[idx]

            prezzo_raw = _first_text(c, [
                ("css", "div.f8F1-price-text"),
                ("css", "div.e2GB-price-text"),
                ("xpath", ".//*[contains(@class,'price-text')]"),
                ("xpath", ".//*[contains(., '€') and (contains(., '0') or contains(., '1') or contains(., '2') or contains(., '3') or contains(., '4') or contains(., '5') or contains(., '6') or contains(., '7') or contains(., '8') or contains(., '9'))]"),
            ])
            prezzo_int = parse_price_to_int(prezzo_raw)

            time_elements = c.find_elements(
                By.XPATH,
                ".//*[contains(@class,'e2Sc-time') or contains(@class,'vmXl')]//span"
            )

            filtered_times = []
            seen_times = set()
            for el in time_elements:
                try:
                    t = el.text.strip().replace("+1", "").strip()
                    if len(t) == 5 and t[2] == ":" and t.replace(":", "").isdigit():
                        if t not in seen_times:
                            seen_times.add(t)
                            filtered_times.append(t)
                except StaleElementReferenceException:
                    continue

            opa = filtered_times[0] if len(filtered_times) > 0 else ""
            oaa = filtered_times[1] if len(filtered_times) > 1 else ""
            opr = filtered_times[2] if len(filtered_times) > 2 else ""
            oar = filtered_times[3] if len(filtered_times) > 3 else ""

            durations = _all_texts(c, [
                ("css", "div.vmXl.vmXl-mod-variant-default"),
                ("xpath", ".//div[contains(@class,'vmXl') and contains(@class,'variant-default')]"),
                ("xpath", ".//div[contains(@class,'vmXl')][contains(.,'min')]"),
            ])

            filtered_dur = []
            seen_dur = set()
            for d in durations:
                dd = d.strip()
                if "min" in dd and any(ch.isdigit() for ch in dd):
                    if dd not in seen_dur and len(dd) <= 25:
                        seen_dur.add(dd)
                        filtered_dur.append(dd)

            da = filtered_dur[0] if len(filtered_dur) > 0 else ""
            dr = filtered_dur[1] if len(filtered_dur) > 1 else ""

            da_min = parse_duration_minutes(da)
            dr_min = parse_duration_minutes(dr)

            total_duration_min = None
            if da_min is not None or dr_min is not None:
                total_duration_min = (da_min or 0) + (dr_min or 0)

            com = _first_text(c, [
                ("css", "div.J0g6-operator-text"),
                ("xpath", ".//div[contains(@class,'J0g6-operator-text')]"),
                ("xpath", ".//*[contains(@class,'operator-text')]"),
            ])

            is_direct_outbound = is_direct_flight(da, com)
            is_direct_return = is_direct_flight(dr, com)
            is_direct = is_direct_outbound and is_direct_return

            score = compute_quality_score(prezzo_int, total_duration_min)

            item = {
                "prezzo_raw": prezzo_raw,
                "prezzo_int": prezzo_int,
                "opa": opa, "oaa": oaa, "opr": opr, "oar": oar,
                "da": da, "dr": dr,
                "da_min": da_min,
                "dr_min": dr_min,
                "total_duration_min": total_duration_min,
                "com": com,
                "is_direct": is_direct,
                "score": score,
            }

            if should_keep_flight(item):
                results.append(item)

        except StaleElementReferenceException:
            print(f"⚠️ Card {idx} diventata stale, la salto.")
            continue
        except Exception as e:
            print(f"⚠️ Errore parsing card {idx}: {e}")
            continue

    return results


# ============================================================
# 8) CHART PNG
# ============================================================

def render_price_chart_png(rows: List[dict], png_path: str, title: str) -> None:
    if not rows:
        return

    import matplotlib as mpl
    import matplotlib.dates as mdates
    from matplotlib.ticker import FuncFormatter, MaxNLocator

    points = extract_round_min_points(rows)
    if len(points) < 2:
        return

    mpl.rcParams.update({
        "font.family": "DejaVu Sans",
        "font.size": 11,
        "axes.titlesize": 20,
        "axes.labelsize": 12,
        "xtick.labelsize": 10,
        "ytick.labelsize": 10,
        "figure.dpi": 180,
        "savefig.dpi": 240,
    })

    xs = [p["saved_at"] for p in points]
    ys = [p["price"] for p in points]

    smooth_ys = []
    window = 5
    for i in range(len(ys)):
        start = max(0, i - window + 1)
        smooth_ys.append(sum(ys[start:i+1]) / len(ys[start:i+1]))

    running_best = []
    best_so_far = float("inf")
    for y in ys:
        best_so_far = min(best_so_far, y)
        running_best.append(best_so_far)

    min_idx = min(range(len(ys)), key=lambda i: ys[i])
    last_idx = len(ys) - 1

    min_p = points[min_idx]
    last_p = points[last_idx]

    fig, ax = plt.subplots(figsize=(15.5, 8.2))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    # Colori eleganti
    main_color = "#2F6DB2"
    best_color = "#E67E22"
    min_marker_color = "#C0392B"
    last_marker_color = "#1F7A4D"
    fill_color = "#DCEAF7"

    # Area tra minimo storico e serie reale
    ax.fill_between(
        xs,
        running_best,
        ys,
        where=[y >= rb for y, rb in zip(ys, running_best)],
        interpolate=True,
        alpha=0.18,
        color=fill_color,
        zorder=1
    )

    # Linea prezzo round
    ax.plot(
        xs, ys,
        color=main_color,
        linewidth=2.6,
        alpha=0.95,
        solid_capstyle="round",
        label="Minimo per round",
        zorder=3
    )


    # Linea trend (smooth)
    ax.plot(        
        xs, smooth_ys,
        linewidth=2.0,
        alpha=0.35,
        color="#1B3A57",
        label="Trend breve",
        zorder=2.5
    )

    # Marker discreti
    ax.scatter(
        xs, ys,
        s=22,
        color=main_color,
        alpha=0.65,
        zorder=4
    )

    # Minimo storico progressivo
    ax.plot(
        xs, running_best,
        linestyle=(0, (4, 3)),
        linewidth=1.8,
        color=best_color,
        alpha=0.95,
        label="Minimo storico",
        zorder=2
    )

    # Evidenzia min assoluto e ultimo
    ax.scatter(
        [xs[min_idx]], [ys[min_idx]],
        s=150,
        color=min_marker_color,
        edgecolors="white",
        linewidths=1.4,
        marker="v",
        zorder=6
    )

    ax.scatter(
        [xs[last_idx]], [ys[last_idx]],
        s=145,
        color=last_marker_color,
        edgecolors="white",
        linewidths=1.4,
        marker="o",
        zorder=6
    )

    # Annotazione minimo assoluto
    ax.annotate(
        (
            f"Min assoluto\n"
            f"{int(ys[min_idx])} €\n"
            f"{min_p['ga']} / {min_p['gr']}"
        ),
        xy=(xs[min_idx], ys[min_idx]),
        xytext=(-10, -58),
        textcoords="offset points",
        ha="center",
        va="top",
        fontsize=10,
        bbox=dict(
            boxstyle="round,pad=0.4",
            fc="white",
            ec="#D9D9D9",
            alpha=0.97
        ),
        arrowprops=dict(
            arrowstyle="-",
            color="#A0A0A0",
            lw=0.8,
            alpha=0.8
        )
    )

    # Annotazione ultimo round
    ax.annotate(
        (
            f"Ultimo round\n"
            f"{int(ys[last_idx])} €\n"
            f"{last_p['ga']} / {last_p['gr']}"
        ),
        xy=(xs[last_idx], ys[last_idx]),
        xytext=(18, 10),
        textcoords="offset points",
        ha="left",
        va="bottom",
        fontsize=10,
        bbox=dict(
            boxstyle="round,pad=0.4",
            fc="white",
            ec="#D9D9D9",
            alpha=0.97
        ),
        arrowprops=dict(
            arrowstyle="-",
            color="#A0A0A0",
            lw=0.8,
            alpha=0.8
        )
    )

    # Titolo + sottotitolo
    route_txt = f"{min_p['ga'].split()[0]} → {min_p['gr'].split()[0]}" if min_p.get("ga") and min_p.get("gr") else ""
    subtitle = (
        f"Osservazioni aggregate per round · "
        f"{len(points)} round osservati"
    )

    ax.set_title(title, pad=22, weight="semibold")
    ax.text(
        0.5, 1.02,
        subtitle,
        transform=ax.transAxes,
        ha="center",
        va="bottom",
        fontsize=10.5,
        color="#666666"
    )

    ax.set_xlabel("Momento della ricerca", labelpad=10)
    ax.set_ylabel("Prezzo minimo round (€)", labelpad=10)

    # Formattazione assi
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, pos: f"{int(v)} €"))
    ax.yaxis.set_major_locator(MaxNLocator(nbins=7))

    total_span = (xs[-1] - xs[0]).total_seconds() / 3600.0
    if total_span <= 48:
        formatter = mdates.DateFormatter("%d/%m\n%H:%M")
    else:
        formatter = mdates.DateFormatter("%d/%m")

    locator = mdates.AutoDateLocator(minticks=5, maxticks=9)
    ax.xaxis.set_major_locator(locator)
    ax.xaxis.set_major_formatter(formatter)

    # Griglia elegante
    ax.grid(axis="y", alpha=0.18, linewidth=0.8)
    ax.grid(axis="x", alpha=0.05, linewidth=0.6)

    # Spines pulite
    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)

    ax.spines["left"].set_alpha(0.18)
    ax.spines["bottom"].set_alpha(0.18)
    ax.tick_params(axis="both", which="both", length=0, pad=8)

    # Limiti asse Y con padding ragionevole
    y_min = min(ys)
    y_max = max(ys)
    spread = y_max - y_min
    padding = max(10, spread * 0.20 if spread > 0 else 18)
    ax.set_ylim(y_min - padding, y_max + padding)

    # Box statistiche elegante
    avg_price = round(sum(ys) / len(ys), 1)
    stats_txt = (
        f"Round osservati: {len(points)}\n"
        f"Min: {int(min(ys))} €\n"
        f"Max: {int(max(ys))} €\n"
        f"Media: {avg_price} €\n"
        f"Ultimo: {int(ys[-1])} €"
    )

    ax.text(
        0.015, 0.98, stats_txt,
        transform=ax.transAxes,
        ha="left",
        va="top",
        fontsize=10,
        color="#2F2F2F",
        bbox=dict(
            boxstyle="round,pad=0.45",
            fc="white",
            ec="#DDDDDD",
            alpha=0.98
        )
    )

    # Legenda discreta
    ax.legend(
        frameon=False,
        loc="upper right",
        fontsize=10.5
    )

    fig.tight_layout()
    fig.savefig(png_path, bbox_inches="tight", facecolor="white")
    plt.close(fig)


# ============================================================
# 9) EXCEL + OPENPYXL CHART + SUMMARY
# ============================================================

def _ensure_workbook(path: str, headers: List[str]) -> tuple:
    p = Path(path)
    if p.exists():
        wb = load_workbook(p)
        ws = wb[EXCEL_SHEET_NAME] if EXCEL_SHEET_NAME in wb.sheetnames else wb.create_sheet(EXCEL_SHEET_NAME)
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = EXCEL_SHEET_NAME

    if ws.max_row == 1 and ws.cell(1, 1).value is None:
        for c, h in enumerate(headers, start=1):
            ws.cell(1, c, h)

    return wb, ws


def _clear_existing_charts(ws):
    try:
        ws._charts = []
    except Exception:
        pass


def update_excel_and_chart(excel_path: str, rows: List[dict], chart_png_path: str, route_name: str) -> None:
    headers = [
        "SEARCHED_AT", "SAVED_AT", "ROUTE",
        "GA", "GA_LABEL", "GA_DOW",
        "GR", "GR_LABEL", "GR_DOW",
        "OPA", "OAA", "OPR", "OAR",
        "DA", "DR", "DA_MIN", "DR_MIN", "TOTAL_DURATION_MIN",
        "PT", "COM", "IS_DIRECT", "SCORE", "IS_NEW_MIN", "STATUS",
        "ROUND_ID", "ROUND_MIN_FOR_GA"
    ]

    wb, ws = _ensure_workbook(excel_path, headers)

    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    sortable_rows = list(rows)

    def _row_sort_key(rr):
        saved = str(rr.get("SAVED_AT", "1900-01-01 00:00:00"))
        ga = str(rr.get("GA", "1900-01-01"))
        gr = str(rr.get("GR", "1900-01-01"))
        pt = safe_float(rr.get("PT"), 999999)
        return (saved, ga, gr, pt)

    sortable_rows.sort(key=_row_sort_key)

    for row in sortable_rows:
        ws.append([row.get(h, "") for h in headers])

    last_row = ws.max_row
    _clear_existing_charts(ws)

    if last_row >= 3:
        data_ref = Reference(ws, min_col=19, min_row=1, max_col=19, max_row=last_row)
        cats_ref = Reference(ws, min_col=2, min_row=2, max_row=last_row)

        chart = LineChart()
        chart.style = 2
        chart.legend = None
        chart.title = f"Andamento prezzi RT {route_name}"
        chart.y_axis.title = "Prezzo (€)"
        chart.x_axis.title = "Data osservazione"
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = False
        ws.add_chart(chart, "Z2")

    if "Summary" in wb.sheetnames:
        ws_summary = wb["Summary"]
        wb.remove(ws_summary)
    ws_summary = wb.create_sheet("Summary")

    stats = summarize_existing_data(sortable_rows)
    best_score_row = stats.get("best_score_row")

    ws_summary.append(["Metric", "Value"])
    ws_summary.append(["Route", route_name])
    ws_summary.append(["Saved observations", stats.get("saved_observations")])
    ws_summary.append(["Rounds observed", stats.get("round_count")])
    ws_summary.append(["Historical min round price", stats.get("min_round_price")])
    ws_summary.append(["Historical max round price", stats.get("max_round_price")])
    ws_summary.append(["Average round min price", stats.get("avg_round_price")])
    ws_summary.append(["Last round min price", stats.get("last_round_price")])
    ws_summary.append(["Direct saved observations", stats.get("direct_saved_count")])
    ws_summary.append(["Direct round minima", stats.get("direct_round_count")])

    if best_score_row:
        ws_summary.append([
            "Best score flight",
            f"{best_score_row.get('GA_LABEL')} | {best_score_row.get('GR_LABEL')} | "
            f"{best_score_row.get('PT')} € | {best_score_row.get('COM')} | "
            f"Score={round(safe_float(best_score_row.get('SCORE'), 0), 2)}"
        ])

    tmp_excel_path = excel_path + ".tmp.xlsx"
    wb.save(tmp_excel_path)
    os.replace(tmp_excel_path, excel_path)  

    render_price_chart_png(
        rows=sortable_rows,
        png_path=chart_png_path,
        title=f"Andamento prezzi ROUNDTRIP {route_name}"
    )


# ============================================================
# 10) BUILD ROUND ROWS
# ============================================================

def build_roundtrip_rows_for_round(
    valid_items: List[Dict],
    rk: str,
    dep: str,
    ret: str,
    searched_at: str,
    current_global_min: float
) -> Tuple[List[Dict], List[Dict], float]:
    """
    Costruisce tutte le righe da salvare per un round di ricerca.
    Ritorna:
    - rows_to_append
    - new_global_min_rows
    - updated_global_min
    """
    if not valid_items:
        return [], [], current_global_min

    ga_dow = weekday_short(dep)
    gr_dow = weekday_short(ret)
    ga_label = f"{dep} ({ga_dow})"
    gr_label = f"{ret} ({gr_dow})"
    round_id = searched_at

    round_prices = [it["prezzo_int"] for it in valid_items if it.get("prezzo_int") is not None]
    if not round_prices:
        return [], [], current_global_min

    round_min_price = min(round_prices)

    rows_to_append = []
    new_global_min_rows = []
    running_min = current_global_min

    for item in valid_items:
        p = item.get("prezzo_int")
        if p is None:
            continue

        is_round_min = (p == round_min_price)
        is_new_global_min = p < running_min

        if is_new_global_min:
            running_min = p

        row = {
            "SEARCHED_AT": searched_at,
            "SAVED_AT": searched_at,
            "ROUTE": rk,
            "GA": dep,
            "GA_LABEL": ga_label,
            "GA_DOW": ga_dow,
            "GR": ret,
            "GR_LABEL": gr_label,
            "GR_DOW": gr_dow,
            "OPA": item.get("opa", ""),
            "OAA": item.get("oaa", ""),
            "OPR": item.get("opr", ""),
            "OAR": item.get("oar", ""),
            "DA": item.get("da", ""),
            "DR": item.get("dr", ""),
            "DA_MIN": item.get("da_min"),
            "DR_MIN": item.get("dr_min"),
            "TOTAL_DURATION_MIN": item.get("total_duration_min"),
            "PT": p,
            "COM": item.get("com", ""),
            "IS_DIRECT": item.get("is_direct", False),
            "SCORE": item.get("score"),
            "IS_NEW_MIN": is_new_global_min,
            "STATUS": "ROUND_MIN" if is_round_min else "OBS",
            "ROUND_ID": round_id,
            "ROUND_MIN_FOR_GA": dep,
        }

        if SAVE_ALL_OBSERVATIONS or is_round_min:
            rows_to_append.append(row)

        if is_new_global_min:
            new_global_min_rows.append(row)

    return rows_to_append, new_global_min_rows, running_min


# ============================================================
# 11) DRIVER FACTORY
# ============================================================

def create_driver():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--disable-infobars")
    chrome_options.add_argument("--lang=it-IT")
    chrome_options.add_argument("--window-size=1440,1100")
    chrome_options.add_argument("--no-first-run")
    chrome_options.add_argument("--no-default-browser-check")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-background-networking")
    chrome_options.add_argument("--remote-debugging-address=127.0.0.1")
    chrome_options.add_argument("--remote-debugging-port=0")

    chrome_binary = os.getenv("CHROME_BINARY", "").strip()
    if chrome_binary:
        chrome_options.binary_location = chrome_binary

    no_display = not os.getenv("DISPLAY")
    use_headless = HEADLESS_MODE in {"1", "true", "yes", "headless"} or (
        HEADLESS_MODE == "auto" and no_display
    )

    if use_headless:
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-setuid-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--disable-software-rasterizer")
        chrome_options.add_argument("--disable-features=VizDisplayCompositor")
        chrome_profile = tempfile.mkdtemp(prefix="botflights_chrome_")
        chrome_options.add_argument(f"--user-data-dir={chrome_profile}")
    else:
        chrome_options.add_argument("--start-maximized")

    driver = webdriver.Chrome(service=Service(), options=chrome_options)
    driver.set_page_load_timeout(60)
    return driver


def restart_driver(driver=None):
    if driver is not None:
        try:
            driver.quit()
        except Exception:
            pass

    driver = create_driver()
    wait = WebDriverWait(driver, 25)
    return driver, wait


def is_invalid_session_error(exc: Exception) -> bool:
    msg = str(exc).lower()
    return isinstance(exc, InvalidSessionIdException) or "invalid session id" in msg


def pause_bot_after_error(route_name: str, dep: str, ret: str, reason: str, seconds: int = ERROR_PAUSE_SEC) -> None:
    seconds = max(1, int(seconds))
    log_round_status(route_name, dep, ret, "PAUSE", f"{reason}; riprovo tra {seconds}s")
    time.sleep(seconds)


# ============================================================
# 12) BEST-10-FUNCTION
# ============================================================


def get_top_10_flights(rows):
    valid = [r for r in rows if r.get("PT") is not None]
    return sorted(valid, key=lambda x: x["PT"])[:10]


def load_manual_routes() -> List[Tuple[str, str, str, str, int]]:
    routes = []

    for idx, route in enumerate(MANUAL_ROUTES[:3], start=1):
        if len(route) != 5:
            raise ValueError(f"MANUAL_ROUTES[{idx}] deve avere 5 valori.")

        departure, destination, dep_date, ret_date, threshold = route
        dep_iata = to_iata(departure)
        dst_iata = to_iata(destination)

        if len(dep_iata) != 3 or len(dst_iata) != 3:
            raise ValueError(f"Tratta {idx} non valida: usa codici IATA da 3 lettere.")

        dep_dt = datetime.datetime.strptime(dep_date, "%Y-%m-%d").date()
        ret_dt = datetime.datetime.strptime(ret_date, "%Y-%m-%d").date()
        if ret_dt <= dep_dt:
            raise ValueError(f"Tratta {idx} non valida: ritorno non successivo all'andata.")

        routes.append((dep_iata, dst_iata, dep_date, ret_date, int(threshold)))

    return routes


# ============================================================
# 13) MAIN
# ============================================================
def main():

    # ============================================================
    # ROTTE CONFIGURATE DA CODICE
    # ============================================================

    routes = load_manual_routes()

    if not routes:
        print("❌ Nessuna rotta configurata in MANUAL_ROUTES.")
        return

    print("✈️ Tratte configurate da codice:")
    for dep, dst, dep_date, ret_date, threshold in routes:
        print(f"   {dep}-{dst} | {dep_date} -> {ret_date} | soglia {threshold}€")

    # ============================================================
    # 📁 CARTELLA OUTPUT
    # ============================================================

    cartella = str(Path.home() / "Desktop" / "BOTFLIGHTS_LOCAL")
    os.makedirs(cartella, exist_ok=True)

    route_data = {}

    # ============================================================
    # 🔥 SETUP PER OGNI ROTTA
    # ============================================================

    for route in routes:

        selected_departure, selected_destination, selected_departure_date, selected_return_date, selected_threshold = route

        dep_iata = to_iata(selected_departure)
        dst_iata = to_iata(selected_destination)
        rk = route_key(dep_iata, dst_iata)

        excel_path = os.path.join(cartella, f"Dati voli ROUNDTRIP {rk}.xlsx")
        chart_png_path = os.path.join(cartella, f"price_chart_roundtrip_{rk}.png")

        existing_rows = load_existing_rows(excel_path)
        stats = summarize_existing_data(existing_rows)

        existing_min = stats.get("min_round_price")
        prezzo_min = float(existing_min) if existing_min else float(selected_threshold)
        current_round_count = int(stats.get("round_count") or 0)

        base_dep_dt = datetime.datetime.strptime(selected_departure_date, "%Y-%m-%d").date()
        base_ret_dt = datetime.datetime.strptime(selected_return_date, "%Y-%m-%d").date()
        base_trip_days = max((base_ret_dt - base_dep_dt).days, MIN_STAY_DAYS)

        route_data[rk] = {
            "dep": selected_departure,
            "dst": selected_destination,
            "dep_date": selected_departure_date,
            "ret_date": selected_return_date,
            "base_dep_dt": base_dep_dt,
            "base_ret_dt": base_ret_dt,
            "trip_days": base_trip_days,
            "threshold": selected_threshold,
            "excel": excel_path,
            "png": chart_png_path,
            "rows": existing_rows,
            "prezzo_min": prezzo_min,
            "checks": 0,
            "last_ml_prediction_round_count": current_round_count - (current_round_count % 200),
            "last_low_price_alert_check": -LOW_PRICE_ALERT_COOLDOWN_CHECKS,
            "last_low_price_alert_value": None,
            "cooldown_until": 0.0,
            "captcha_count": 0,
        }

    # ============================================================
    # 🔥 DRIVER (UNO SOLO → più efficiente)
    # ============================================================

    driver, wait = restart_driver()

    route_keys = list(route_data.keys())
    conta = 0

    notify_startup_once()


    # ============================================================
    # 🔥 LOOP MULTI ROUTE
    # ============================================================

    try:
        while True:

            try:
                # 🔥 SCEGLIE UNA ROTTA RANDOM TRA LE 3
                now_mono = time.monotonic()
                available_route_keys = [
                    key for key in route_keys
                    if now_mono >= float(route_data[key].get("cooldown_until", 0.0))
                ]
                if not available_route_keys:
                    next_available = min(float(route_data[key].get("cooldown_until", 0.0)) for key in route_keys)
                    sleep_for = max(5, int(next_available - now_mono))
                    print(f"🧊 Tutte le tratte sono in cooldown captcha. Attendo {sleep_for}s.")
                    time.sleep(sleep_for)
                    continue

                rk = random.choice(available_route_keys)
                data = route_data[rk]

                selected_departure = data["dep"]
                selected_destination = data["dst"]
                base_dep_dt = data["base_dep_dt"]
                base_ret_dt = data["base_ret_dt"]
                base_trip_days = data["trip_days"]

                excel_path = data["excel"]
                chart_png_path = data["png"]
                existing_rows = data["rows"]
                prezzo_min = data["prezzo_min"]
                route_checks = int(data.get("checks", 0)) + 1
                data["checks"] = route_checks

                # ============================================================
                # 🔥 RANDOM DATE SEARCH
                # ============================================================

                if conta == 0:
                    dep = data["dep_date"]
                    ret = data["ret_date"]
                else:
                    dep_delta = random.randint(-OUTBOUND_DELTA_DAYS, OUTBOUND_DELTA_DAYS)
                    ret_delta = random.randint(-RETURN_DELTA_DAYS, RETURN_DELTA_DAYS)

                    dep_dt = base_dep_dt + datetime.timedelta(days=dep_delta)
                    ret_dt = base_ret_dt + datetime.timedelta(days=ret_delta)

                    if ret_dt <= dep_dt:
                        ret_dt = dep_dt + datetime.timedelta(days=base_trip_days)

                    dep = dep_dt.strftime("%Y-%m-%d")
                    ret = ret_dt.strftime("%Y-%m-%d")

                url = build_kayak_url(selected_departure, selected_destination, dep, ret)
                print(f"[{conta}] {rk} → {url}")

                if not has_internet_connection():
                    wait_until_connection_restored(rk, dep, ret)

                try:
                    driver.get(url)
                except WebDriverException as e:
                    if is_invalid_session_error(e):
                        log_round_status(rk, dep, ret, "DRIVER_RESTART", "sessione Selenium non valida, riavvio ChromeDriver")
                        driver, wait = restart_driver(driver)
                        pause_bot_after_error(rk, dep, ret, "sessione Selenium persa", DRIVER_RESTART_PAUSE_SEC)
                        conta += 1
                        continue
                    if is_connectivity_error(e) or not has_internet_connection():
                        wait_until_connection_restored(rk, dep, ret)
                        conta += 1
                        continue
                    log_round_status(rk, dep, ret, "WEBDRIVER_ERROR", str(e).splitlines()[0])
                    driver, wait = restart_driver(driver)
                    pause_bot_after_error(rk, dep, ret, "errore Selenium recuperabile")
                    conta += 1
                    continue
                time.sleep(3)

                accept_cookies_best_effort(driver)
                valid_items = []
                if is_captcha_page(driver):
                    valid_items = collect_valid_items_best_effort(driver)
                    if valid_items:
                        log_round_status(
                            rk, dep, ret, "CAPTCHA_WITH_RESULTS",
                            f"captcha presente ma risultati leggibili: salvo {len(valid_items)} voli"
                        )
                        put_route_in_captcha_cooldown(route_data, rk, "risultati leggibili con captcha")
                    else:
                        if handle_captcha_if_present(driver, rk):
                            if is_captcha_page(driver):
                                log_round_status(rk, dep, ret, "CAPTCHA_TIMEOUT", "captcha non risolto entro il timeout")
                                put_route_in_captcha_cooldown(route_data, rk, "timeout captcha iniziale")
                                conta += 1
                                continue
                            log_round_status(rk, dep, ret, "CAPTCHA_OK", "captcha risolto, riprendo il controllo")

                if not valid_items:
                    try:
                        wait_for_results(driver, wait)
                        time.sleep(2)
                    except TimeoutException:
                        if not has_internet_connection():
                            wait_until_connection_restored(rk, dep, ret)
                            conta += 1
                            continue
                        if is_captcha_page(driver):
                            valid_items = collect_valid_items_best_effort(driver)
                            if valid_items:
                                log_round_status(
                                    rk, dep, ret, "CAPTCHA_WITH_RESULTS",
                                    f"captcha presente dopo timeout, ma risultati leggibili: salvo {len(valid_items)} voli"
                                )
                                put_route_in_captcha_cooldown(route_data, rk, "risultati leggibili dopo captcha")
                            elif handle_captcha_if_present(driver, rk):
                                if is_captcha_page(driver):
                                    log_round_status(rk, dep, ret, "CAPTCHA_TIMEOUT", "blocco captcha durante il caricamento risultati")
                                    put_route_in_captcha_cooldown(route_data, rk, "timeout captcha caricamento risultati")
                                    conta += 1
                                    continue
                                accept_cookies_best_effort(driver)
                                try:
                                    wait_for_results(driver, wait)
                                    time.sleep(2)
                                except TimeoutException:
                                    log_round_status(rk, dep, ret, "TIMEOUT", "risultati non trovati dopo la risoluzione captcha")
                                    conta += 1
                                    continue
                        else:
                            log_round_status(rk, dep, ret, "TIMEOUT", "risultati non trovati / pagina bloccata")
                            conta += 1
                            continue
                    except WebDriverException as e:
                        if is_invalid_session_error(e):
                            log_round_status(rk, dep, ret, "DRIVER_RESTART", "sessione Selenium persa durante il caricamento risultati")
                            driver, wait = restart_driver(driver)
                            pause_bot_after_error(rk, dep, ret, "sessione Selenium persa durante il caricamento", DRIVER_RESTART_PAUSE_SEC)
                            conta += 1
                            continue
                        if is_connectivity_error(e) or not has_internet_connection():
                            wait_until_connection_restored(rk, dep, ret)
                            conta += 1
                            continue
                        log_round_status(rk, dep, ret, "WEBDRIVER_ERROR", str(e).splitlines()[0])
                        driver, wait = restart_driver(driver)
                        pause_bot_after_error(rk, dep, ret, "errore Selenium durante il caricamento risultati")
                        conta += 1
                        continue

                if not valid_items:
                    try:
                        valid_items = collect_valid_items_best_effort(driver)
                    except Exception as e:
                        log_round_status(rk, dep, ret, "ERROR", f"errore parsing: {e}")
                        pause_bot_after_error(rk, dep, ret, "errore parsing risultati", DRIVER_RESTART_PAUSE_SEC)
                        conta += 1
                        continue

                observation_added = False
                new_min_rows = []
                low_price_signal = None

                if valid_items:
                    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    current_round_min = min(it["prezzo_int"] for it in valid_items if it.get("prezzo_int") is not None)
                    low_price_signal = detect_low_price_signal(existing_rows, current_round_min)

                    rows_to_append, round_new_min_rows, prezzo_min = build_roundtrip_rows_for_round(
                        valid_items,
                        rk,
                        dep,
                        ret,
                        now,
                        prezzo_min,
                    )

                    if rows_to_append:
                        existing_rows.extend(rows_to_append)
                        observation_added = True

                    if round_new_min_rows:
                        new_min_rows.extend(round_new_min_rows)
                else:
                    log_round_status(rk, dep, ret, "EMPTY", "nessun volo valido estratto")

                # ============================================================
                # 🔥 UPDATE + ML
                # ============================================================

                if observation_added:

                    update_excel_and_chart(excel_path, existing_rows, chart_png_path, rk)
                    detail = f"osservazioni={len(valid_items)}"
                    if low_price_signal:
                        detail += f", round_min={low_price_signal['candidate_price']}€"
                    log_round_status(rk, dep, ret, "OK", detail)

                current_round_count = len(extract_round_min_points(existing_rows))
                last_ml_prediction_round_count = int(data.get("last_ml_prediction_round_count", 0))
                if (
                    current_round_count > 0
                    and current_round_count % 200 == 0
                    and current_round_count > last_ml_prediction_round_count
                ):
                    prediction = predict_price(existing_rows)
                    if prediction:
                        prediction["checks"] = route_checks
                        prediction["rounds"] = current_round_count
                        msg = build_ml_prediction_message(rk, prediction)
                        pushover_send(msg, title=f"ML {rk}")
                        data["last_ml_prediction_round_count"] = current_round_count

                if low_price_signal:
                    last_alert_check = int(data.get("last_low_price_alert_check", -LOW_PRICE_ALERT_COOLDOWN_CHECKS))
                    last_alert_value = data.get("last_low_price_alert_value")
                    is_better_price = last_alert_value is None or low_price_signal["candidate_price"] < float(last_alert_value)
                    cooldown_elapsed = (route_checks - last_alert_check) >= LOW_PRICE_ALERT_COOLDOWN_CHECKS
                    if is_better_price or cooldown_elapsed:
                        msg = build_low_price_alert_message(rk, dep, ret, low_price_signal)
                        pushover_send(msg, attachment_path=chart_png_path, title=f"Deal {rk}")
                        data["last_low_price_alert_check"] = route_checks
                        data["last_low_price_alert_value"] = low_price_signal["candidate_price"]

                # ============================================================
                # 🔥 NUOVI MINIMI
                # ============================================================

                for row in new_min_rows:
                    msg = (
                        f"🔥 NEW MIN {rk}\n"
                        f"{row['GA_LABEL']} → {row['GR_LABEL']}\n"
                        f"{row['PT']}€ | {row['COM']}"
                    )
                    pushover_send(msg, attachment_path=chart_png_path)

                # ============================================================
                # 🔥 UPDATE DATI
                # ============================================================

                route_data[rk]["rows"] = existing_rows
                route_data[rk]["prezzo_min"] = prezzo_min

                conta += 1
                time.sleep(random.randint(LOOP_SLEEP_MIN_SEC, LOOP_SLEEP_MAX_SEC))

            except KeyboardInterrupt:
                raise
            except Exception as e:
                err_rk = locals().get("rk", "UNKNOWN")
                err_dep = locals().get("dep", "N/A")
                err_ret = locals().get("ret", "N/A")
                log_round_status(err_rk, err_dep, err_ret, "UNEXPECTED_ERROR", str(e).splitlines()[0])
                try:
                    driver, wait = restart_driver(driver)
                except Exception as restart_error:
                    log_round_status(err_rk, err_dep, err_ret, "DRIVER_RESTART_FAILED", str(restart_error).splitlines()[0])
                pause_bot_after_error(err_rk, err_dep, err_ret, "errore imprevisto recuperabile")
                conta += 1
                continue

    finally:
        try:
            driver.quit()
        except:
            pass


if __name__ == "__main__":
    main()
