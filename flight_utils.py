# funzioni_utili_ricerca_voli.py
import random
import time
from datetime import datetime, timedelta
import calendar
import tkinter as tk
from tkinter import messagebox
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# ---------- Date helpers ----------
def converti_data_testuale(selected_date: str) -> str:
    d = datetime.strptime(selected_date, "%Y-%m-%d")
    return d.strftime("%d %B %Y")


def giorno_settimana(data_str: str) -> str:
    d = datetime.strptime(data_str, "%Y-%m-%d")
    return d.strftime("%A")


def numero_a_nome_mese(numero_mese: int) -> str:
    try:
        return calendar.month_name[numero_mese]
    except IndexError:
        return "Mese non valido"


def is_first_date_greater(date1: str, date2: str) -> bool:
    try:
        d1 = datetime.strptime(date1, "%Y-%m-%d")
        d2 = datetime.strptime(date2, "%Y-%m-%d")
        return d1 > d2
    except ValueError:
        return False


def aggiungi_giorni(data_str: str, ngiorni: int):
    d = datetime.strptime(data_str, "%Y-%m-%d")
    future = d + timedelta(days=ngiorni)
    return future.strftime("%Y-%m-%d"), future.strftime("%A")


def find_year(month: int, day: int, weekday3: str):
    current_year = datetime.now().year
    weekdays_map = {'Mon': 0, 'Tue': 1, 'Wed': 2, 'Thu': 3, 'Fri': 4, 'Sat': 5, 'Sun': 6}
    if weekday3 not in weekdays_map:
        raise ValueError("weekday deve essere tipo 'Mon', 'Tue', ...")
    target = weekdays_map[weekday3]

    for year in range(current_year - 1, current_year + 2):
        try:
            d = datetime(year, month, day)
            if d.weekday() == target:
                return year
        except ValueError:
            continue
    return None


# ---------- UI alert ----------
def show_alert(message: str):
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Errore", message)
    root.destroy()


# ---------- Anti-bot movement (light) ----------
def move_random(driver, moves: int = 4):
    chain = ActionChains(driver)
    for _ in range(moves):
        x = random.randint(10, 120)
        y = random.randint(10, 80)
        chain.move_by_offset(x, y)
        time.sleep(random.uniform(0.3, 1.0))
    try:
        chain.perform()
    except Exception:
        pass


# ---------- SAFE “book” hook ----------
def open_booking_provider_best_effort(driver, timeout: int = 20):
    """
    Fa SOLO un click sul link “Prenota su …” se esiste e poi si ferma.
    NON compila dati personali e NON procede al pagamento.
    """
    wait = WebDriverWait(driver, timeout)
    # prova vari testi: Kayak cambia provider
    xpaths = [
        "//a[contains(., 'Prenota')]",
        "//a[contains(., 'Book')]",
        "//button[contains(., 'Prenota')]",
        "//button[contains(., 'Book')]",
    ]
    for xp in xpaths:
        try:
            el = wait.until(EC.element_to_be_clickable((By.XPATH, xp)))
            el.click()
            return True
        except Exception:
            continue
    return False