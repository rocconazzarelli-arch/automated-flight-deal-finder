import csv
import urllib.request
from pathlib import Path
import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
import locale

from flight_utils import show_alert

# =========================
# 🌍 FIX LOCALE (CRITICO)
# =========================
try:
    locale.setlocale(locale.LC_ALL, "en_US.UTF-8")
except:
    pass

# =========================
# 🎨 THEME
# =========================
BG = "#0f172a"
PANEL = "#1e293b"
FG = "#f8fafc"
ACCENT = "#2563eb"

# =========================
# ✈ Airports Data
# =========================
AIRPORTS_URL = "https://raw.githubusercontent.com/jpatokal/openflights/master/data/airports.dat"

CACHE_DIR = Path.home() / ".botflights_cache"
CACHE_DIR.mkdir(parents=True, exist_ok=True)
AIRPORTS_CACHE = CACHE_DIR / "airports.dat"


def _download_airports_if_needed():
    if AIRPORTS_CACHE.exists() and AIRPORTS_CACHE.stat().st_size > 1000:
        return
    urllib.request.urlretrieve(AIRPORTS_URL, AIRPORTS_CACHE)


def _load_airports():
    _download_airports_if_needed()
    airports = []
    with open(AIRPORTS_CACHE, "r", encoding="utf-8", errors="ignore") as f:
        reader = csv.reader(f)
        for row in reader:
            if len(row) < 5:
                continue
            name, city, country, iata = row[1], row[2], row[3], row[4]
            if not iata or iata == r"\N":
                continue
            airports.append({
                "city": city,
                "country": country,
                "name": name,
                "iata": iata
            })
    return airports


def _format_option(a):
    return f"{a['city']} ({a['iata']}) — {a['name']}, {a['country']}"


_AIRPORTS = _load_airports()

# =========================
# 🖥 UI
# =========================
def create_flight_search_ui():
    result = None

    root = tk.Tk()
    root.title("Flight Search")
    root.configure(bg=BG)
    root.resizable(False, False)

    # 🎯 FIX tema ttk (importante)
    style = ttk.Style()
    style.theme_use("default")

    def close_panel():
        try:
            root.quit()
        except tk.TclError:
            pass
        try:
            root.destroy()
        except tk.TclError:
            pass

    def on_close():
        nonlocal result
        result = None
        close_panel()

    root.protocol("WM_DELETE_WINDOW", on_close)

    W, H = 900, 720
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (W // 2)
    y = (root.winfo_screenheight() // 2) - (H // 2)
    root.geometry(f"{W}x{H}+{x}+{y}")

    outer = ttk.Frame(root)
    outer.pack(fill="both", expand=True)

    panel = ttk.Frame(outer, padding=40)
    panel.pack(padx=60, pady=20, fill="both", expand=True)

    row = 0

    ttk.Label(panel, text="Departure").grid(row=row, column=0)
    row += 1
    dep_entry = ttk.Entry(panel, width=50)
    dep_entry.grid(row=row, column=0)
    row += 1
    departure_combobox = ttk.Combobox(panel, width=70, state="readonly")
    departure_combobox.grid(row=row, column=0)
    row += 1

    ttk.Label(panel, text="Destination").grid(row=row, column=0)
    row += 1
    dst_entry = ttk.Entry(panel, width=50)
    dst_entry.grid(row=row, column=0)
    row += 1
    destination_combobox = ttk.Combobox(panel, width=70, state="readonly")
    destination_combobox.grid(row=row, column=0)
    row += 1

    # =========================
    # 📅 DATE PICKERS (FIXATI)
    # =========================
    ttk.Label(panel, text="Departure Date").grid(row=row, column=0)
    row += 1

    dep_date = DateEntry(
        panel,
        date_pattern="yyyy-mm-dd",
        locale="en_US",
        showweeknumbers=False
    )
    dep_date.grid(row=row, column=0)
    row += 1

    ttk.Label(panel, text="Return Date").grid(row=row, column=0)
    row += 1

    ret_date = DateEntry(
        panel,
        date_pattern="yyyy-mm-dd",
        locale="en_US",
        showweeknumbers=False
    )
    ret_date.grid(row=row, column=0)
    row += 1

    ttk.Label(panel, text="Price Threshold (€)").grid(row=row, column=0)
    row += 1
    threshold_spinbox = tk.Spinbox(panel, from_=0, to=9999)
    threshold_spinbox.grid(row=row, column=0)
    row += 1

    # =========================
    # 🔎 FILTRO AEROPORTI
    # =========================
    def build_options(text):
        t = (text or "").lower().strip()
        if not t:
            return []
        return [_format_option(a) for a in _AIRPORTS if t in _format_option(a).lower()][:200]

    dep_entry.bind("<KeyRelease>", lambda e: departure_combobox.configure(values=build_options(dep_entry.get())))
    dst_entry.bind("<KeyRelease>", lambda e: destination_combobox.configure(values=build_options(dst_entry.get())))

    # =========================
    # ✅ VALIDAZIONE + SUBMIT
    # =========================
    def on_confirm():
        nonlocal result

        dep_val = departure_combobox.get()
        dst_val = destination_combobox.get()

        if not dep_val or not dst_val:
            show_alert("Select both departure and destination.")
            return

        try:
            d1_obj = dep_date.get_date()
            d2_obj = ret_date.get_date()

            if d2_obj <= d1_obj:
                show_alert("Return date must be after departure date.")
                return

            d1 = d1_obj.strftime("%Y-%m-%d")
            d2 = d2_obj.strftime("%Y-%m-%d")

            thr = int(threshold_spinbox.get())
        except Exception:
            show_alert("Invalid input values.")
            return

        result = (dep_val, dst_val, d1, d2, thr)

        close_panel()

    ttk.Button(panel, text="Search", command=on_confirm).grid(row=row, column=0)

    root.mainloop()

    return result
