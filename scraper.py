"""
=============================================================================
SCRAPER BAZY DANYCH — SZKIELET SKRYPTU
=============================================================================
Cel:    Automatyczne pobieranie danych organizacji/firm z Internetu
        i zapis do pliku XLSX zgodnego z formatką firmy.

Użycie: python scraper.py
        python scraper.py --reset    ← usuwa postęp i zaczyna od nowa

Wznowienie po awarii: uruchom ponownie — skrypt czyta progress.json
                      i kontynuuje od miejsca zatrzymania.

Rekonfiguracja pod nowe zadanie: zmień wyłącznie blok CONFIG poniżej.
=============================================================================
"""

import argparse
import json
import logging
import random
import re
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

import requests
import pdfplumber
from bs4 import BeautifulSoup
from io import BytesIO
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from tqdm import tqdm


# =============================================================================
# KONFIGURACJA — zmień tylko tę sekcję przy nowym zadaniu
# =============================================================================
CONFIG = {
    # --- IDENTYFIKACJA ZADANIA ---
    "task_name": "organizacje_polonijne",
    "output_file": "output/baza_organizacje_polonijne.xlsx",
    "rejected_file": "output/odrzucone.xlsx",
    "log_file": "scraper_log.txt",
    "progress_file": "progress.json",

    # --- ETAP 1: ZBIERANIE NAZW OFERENTÓW Z WYNIKÓW KONKURSÓW MSZ/KPRM ---
    # Plik pośredni — tylko nazwa organizacji, projekt, kwota i status.
    # Adres korespondencyjny, KRS, telefon, e-mail itd. uzupełnia dopiero
    # kolejny etap (wzbogacanie przez API KRS/REGON), którego jeszcze nie ma.
    "seed_file": "output/oferenci_konkursy_polonijne.xlsx",

    # Każdy adres poniżej zweryfikowano ręcznie przez realne pobranie strony
    # (patrz historia rozmowy — nazwa i podział konkursu zmienia się rok do
    # roku, więc NIE da się zbudować jednego wzorca adresu zależnego od roku).
    "konkursy_polonijne": [
        {
            "url": "https://www.gov.pl/web/dyplomacja/wyniki-konkursu-polonia-i-polacy-za-granica-2026",
            "rok": 2026,
            "nazwa_konkursu": "Polonia i Polacy za Granicą 2026",
            "archiwum_wayback": False,
        },
        {
            "url": "https://www.gov.pl/web/dyplomacja/ogloszenie-wynikow-konkursu-wsparcie-srodowisk-polonijnych-w-obszarze-aktywizacji-postaw-obywatelskich",
            "rok": 2025,
            "nazwa_konkursu": "Wsparcie środowisk polonijnych w obszarze aktywizacji postaw obywatelskich 2025",
            "archiwum_wayback": False,
        },
        {
            "url": "https://www.gov.pl/web/dyplomacja/wyniki-konkursu-infrastruktura-polonijna-2025",
            "rok": 2025,
            "nazwa_konkursu": "Infrastruktura Polonijna 2025",
            "archiwum_wayback": False,
        },
        {
            "url": "https://www.gov.pl/web/dyplomacja/wyniki-konkursu-polonia-i-polacy-za-granica-2024--media-i-struktury2",
            "rok": 2024,
            "nazwa_konkursu": "Polonia i Polacy za Granicą 2024 – Media i Struktury",
            "archiwum_wayback": False,
        },
        {
            "url": "https://www.gov.pl/web/dyplomacja/ogloszenie-wynikow-konkursu-polonia-i-polacy-za-granica-2024---wydarzenia-i-inicjatywy-polonijne",
            "rok": 2024,
            "nazwa_konkursu": "Polonia i Polacy za Granicą 2024 – Wydarzenia i inicjatywy polonijne",
            "archiwum_wayback": False,
        },
        {
            "url": "https://www.gov.pl/web/dyplomacja/ogloszenie-wynikow-konkursu-polonia-i-polacy-za-granica-2024---regranting",
            "rok": 2024,
            "nazwa_konkursu": "Polonia i Polacy za Granicą 2024 – Regranting",
            "archiwum_wayback": False,
        },
        {
            # Strona oryginalna KPRM już nie istnieje (kompetencje przejęło MSZ
            # w połowie 2024 roku) — dane pobieramy z Wayback Machine.
            "url": "https://web.archive.org/web/20240201084241/https://www.gov.pl/web/polonia/wyniki-konkursu-polonia-i-polacy-za-granica-2023",
            "rok": 2023,
            "nazwa_konkursu": "Polonia i Polacy za Granicą 2023",
            "archiwum_wayback": True,
        },
    ],

    # Załączniki o tytułach zawierających te frazy pomijamy — to zbiorcze
    # listy "wszystkich zgłoszonych ofert", które nie mówią nic o wyniku
    # (dublują się z listą przyznanych + listą odrzuconych/nieprzyjętych).
    "zalaczniki_pomijane_frazy": ["zgłoszon", "wpłynęł"],

    # --- ŹRÓDŁA DANYCH (w kolejności priorytetu) ---
    # Uzupełnij przed uruchomieniem po konsultacji z użytkownikiem.
    # Każde źródło: url (punkt startowy), priority (1=najwyższy), type, notes
    "sources": [
        {
            "id": "ngo_pl",
            "url": "https://baza.ngo.pl/szukaj/organizacje?query=polonijna",
            "priority": 1,
            "type": "authoritative",
            "requires_api_key": False,
            "requires_login": False,
            "rate_limit": "1 req/s",
            "notes": "Główna autorytatywna baza NGO w Polsce",
        },
        # Dodaj kolejne źródła według potrzeb:
        # {
        #     "id": "gov_portal",
        #     "url": "https://...",
        #     "priority": 2,
        #     "type": "supplementary",
        #     "requires_api_key": False,
        #     "requires_login": False,
        #     "rate_limit": "2 req/s",
        #     "notes": "Portal rządowy — uzupełniający",
        # },
    ],

    # --- FILTROWANIE ---
    "filter_country": "Polska",           # siedziba musi być w Polsce
    "filter_active_months": 12,           # aktywna w ciągu ostatnich N miesięcy
    "filter_exclude_sole_trader": False,  # True = wyklucz jednoosobowe działalności

    # --- LIMITY DANYCH ---
    "max_phones_org": 3,                  # maks. telefonów do organizacji
    "max_emails_org": 3,                  # maks. e-maili do organizacji

    # --- SIEĆ ---
    "delay_min": 1.0,                     # minimalne opóźnienie między req (sekundy)
    "delay_max": 3.0,                     # maksymalne opóźnienie między req (sekundy)
    "retry_delay": 5.0,                   # czas przed ponowną próbą po błędzie
    "request_timeout": 15,                # timeout jednego żądania HTTP (sekundy)

    # --- WYJŚCIE ---
    "encoding": "utf-8-sig",             # UTF-8 z BOM = poprawne polskie znaki w Excel
    "date_format": "%d.%m.%Y",           # format daty w kolumnie date_collected
}

# Nagłówek HTTP — symulujemy normalną przeglądarkę
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "pl-PL,pl;q=0.9,en-US;q=0.8,en;q=0.7",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

# Kolumny pliku wyjściowego (bez Lp. — użytkownik wypełnia ręcznie)
OUTPUT_COLUMNS = [
    "Kategoria",
    "Branża / Typ",
    "Nazwa",
    "Adres korespondencyjny",
    "Województwo",
    "Numer telefonu",
    "Adres e-mail",
    "Strona WWW",
    "Profil w mediach społecznościowych",
    "Osoba kontaktowa",
    "Numer telefonu do osoby kontaktowej",
    "Adres e-mail do osoby kontaktowej",
    "Data pozyskania informacji",
    "Krótka charakterystyka podmiotu",
    # Kolumny dodatkowe (poza formatką firmy):
    "URL źródła",
    "KRS",
    "REGON",
    "NIP",
]

# Kolumny pliku pośredniego z etapu 1 (zbieranie oferentów z konkursów).
# To NIE jest jeszcze formatka — brakuje adresu, KRS, kontaktu itd.
SEED_COLUMNS = [
    "Rok",
    "Nazwa konkursu",
    "Lista źródłowa (status)",
    "Oferent",
    "Nazwa projektu",
    "Kwota",
    "URL artykułu",
    "URL załącznika PDF",
]


# =============================================================================
# LOGOWANIE
# =============================================================================

def setup_logging():
    """Konfiguruje logging do pliku i konsoli."""
    logger = logging.getLogger("scraper")
    logger.setLevel(logging.DEBUG)

    # Formatka: zdarzenie | status | czas
    class PipeFormatter(logging.Formatter):
        def format(self, record):
            status = "ERROR" if record.levelno >= logging.ERROR else (
                "INFO" if record.levelno == logging.INFO else "SUCCESS"
            )
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            return f"{record.getMessage()} | {status} | {ts}"

    # Plik logów (dopisywanie, nie nadpisywanie)
    fh = logging.FileHandler(CONFIG["log_file"], mode="a", encoding="utf-8")
    fh.setFormatter(PipeFormatter())
    logger.addHandler(fh)

    # Konsola (czytelniejszy format)
    ch = logging.StreamHandler()
    ch.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
    logger.addHandler(ch)

    return logger


log = setup_logging()


def log_event(event_name: str, success: bool, details: str = ""):
    """Zapisuje zdarzenie do logu w formacie: zdarzenie | SUCCESS/ERROR | czas"""
    msg = event_name if not details else f"{event_name} — {details}"
    if success:
        log.info(msg)
    else:
        log.error(msg)


# =============================================================================
# POSTĘP I WZNOWIENIE
# =============================================================================

def load_progress() -> dict:
    """Wczytuje stan postępu z pliku JSON."""
    p = Path(CONFIG["progress_file"])
    if p.exists():
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    return {"completed_ids": [], "last_source": None, "total_scraped": 0}


def save_progress(progress: dict):
    """Zapisuje aktualny stan postępu."""
    with open(CONFIG["progress_file"], "w", encoding="utf-8") as f:
        json.dump(progress, f, ensure_ascii=False, indent=2)


def reset_progress():
    """Usuwa plik postępu — wymusza start od początku."""
    p = Path(CONFIG["progress_file"])
    if p.exists():
        p.unlink()
        log_event("RESET_POSTEPU", True)


# =============================================================================
# SIEĆ — pobieranie stron
# =============================================================================

def fetch_page(url: str, session: requests.Session) -> BeautifulSoup | None:
    """
    Pobiera stronę i zwraca BeautifulSoup lub None po nieudanych próbach.
    Dwie próby: pierwsza → czeka retry_delay → druga → None.
    """
    for attempt in range(1, 3):
        try:
            resp = session.get(url, headers=HEADERS, timeout=CONFIG["request_timeout"])
            resp.raise_for_status()
            resp.encoding = "utf-8"
            log_event(f"POBIERANIE {url[:60]}", True)
            return BeautifulSoup(resp.text, "lxml")

        except requests.RequestException as e:
            log_event(f"POBIERANIE {url[:60]} (próba {attempt})", False, str(e))
            if attempt < 2:
                time.sleep(CONFIG["retry_delay"])

    return None  # obie próby nieudane


def polite_delay():
    """Losowe opóźnienie między żądaniami — szanujemy serwer."""
    time.sleep(random.uniform(CONFIG["delay_min"], CONFIG["delay_max"]))


# =============================================================================
# PARSOWANIE — dostosuj do każdego źródła
# =============================================================================

def get_entity_urls_from_source(source: dict, session: requests.Session) -> list[str]:
    """
    Pobiera listę URL-i do poszczególnych rekordów ze źródła.

    !! DOSTOSUJ TĘ FUNKCJĘ DO KAŻDEGO ŹRÓDŁA !!

    Przykład dla baza.ngo.pl:
      - pobierz stronę listową
      - znajdź linki do profili organizacji
      - obsłuż paginację (kolejne strony)
    """
    urls = []
    page_url = source["url"]

    while page_url:
        soup = fetch_page(page_url, session)
        if soup is None:
            log_event(f"LISTA_{source['id'].upper()}", False, f"Nie można pobrać: {page_url}")
            break

        # ---- DOSTOSUJ: selektory CSS/XPath dla konkretnego źródła ----
        # Przykład (zastąp prawdziwymi selektorami):
        # links = soup.select("a.organization-link")
        # for link in links:
        #     href = link.get("href", "")
        #     if href:
        #         urls.append(href if href.startswith("http") else BASE_URL + href)

        # Przykład paginacji:
        # next_btn = soup.select_one("a.next-page")
        # page_url = next_btn["href"] if next_btn else None

        # Tymczasowy placeholder — usuń po implementacji:
        log.warning(f"get_entity_urls_from_source: zaimplementuj parsowanie dla źródła '{source['id']}'")
        page_url = None  # usuń gdy paginacja będzie gotowa
        polite_delay()

    log_event(f"LISTA_{source['id'].upper()}", True, f"Znaleziono {len(urls)} URL-i")
    return urls


def parse_entity_page(url: str, soup: BeautifulSoup) -> dict:
    """
    Parsuje stronę jednej organizacji/firmy i zwraca słownik danych.

    !! DOSTOSUJ TĘ FUNKCJĘ DO STRUKTURY STRONY ŹRÓDŁOWEJ !!

    Zwraca dict z kluczami odpowiadającymi OUTPUT_COLUMNS.
    Jeśli danych brak → wartość "none" (lub "brak" / "nie ustalono" zgodnie ze spec.).
    """
    today = datetime.now().strftime(CONFIG["date_format"])

    # ---- DOSTOSUJ: selektory dla konkretnej strony ----
    # Poniżej schemat — zastąp rzeczywistymi selektorami:

    record = {
        "Kategoria":             extract_text(soup, "span.category") or "organizacja pozarządowa",
        "Branża / Typ":          extract_text(soup, "span.industry") or "none",
        "Nazwa":                 extract_text(soup, "h1.org-name") or "none",
        "Adres korespondencyjny": extract_address(soup),
        "Województwo":           extract_text(soup, "span.voivodeship") or "none",
        "Numer telefonu":        extract_phones(soup, max_count=CONFIG["max_phones_org"]),
        "Adres e-mail":          extract_emails(soup, max_count=CONFIG["max_emails_org"]),
        "Strona WWW":            extract_website(soup),
        "Profil w mediach społecznościowych": extract_social(soup),
        "Osoba kontaktowa":      extract_text(soup, "span.contact-person") or "nie ustalono",
        "Numer telefonu do osoby kontaktowej": extract_text(soup, "span.contact-phone") or "brak",
        "Adres e-mail do osoby kontaktowej":   extract_text(soup, "span.contact-email") or "brak",
        "Data pozyskania informacji": today,
        "Krótka charakterystyka podmiotu": extract_description(soup),
        # Kolumny dodatkowe:
        "URL źródła": url,
        "KRS":   extract_registry_number(soup, "krs") or "none",
        "REGON": extract_registry_number(soup, "regon") or "none",
        "NIP":   extract_registry_number(soup, "nip") or "none",
    }

    return record


# =============================================================================
# HELPERY PARSOWANIA — dostosuj selektory CSS
# =============================================================================

def extract_text(soup: BeautifulSoup, selector: str) -> str | None:
    """Zwraca tekst pierwszego elementu pasującego do selektora CSS."""
    el = soup.select_one(selector)
    if el:
        return el.get_text(strip=True) or None
    return None


def extract_address(soup: BeautifulSoup) -> str:
    """
    Wyciąga adres korespondencyjny.
    !! DOSTOSUJ SELEKTOR !!
    """
    # Przykład:
    # parts = [
    #     extract_text(soup, "span.street"),
    #     extract_text(soup, "span.postcode"),
    #     extract_text(soup, "span.city"),
    # ]
    # return ", ".join(p for p in parts if p) or "none"
    return "none"  # zastąp implementacją


def extract_phones(soup: BeautifulSoup, max_count: int = 3) -> str:
    """
    Wyciąga telefony do organizacji (maks. max_count).
    Priorytet: sekretariat > biuro > ogólny.
    Format wynikowy: XXX XXX XXX, kolejne oddzielone \\n
    !! DOSTOSUJ SELEKTOR !!
    """
    # Przykład:
    # raw_phones = [el.get_text(strip=True) for el in soup.select("span.phone")]
    # cleaned = [normalize_phone(p) for p in raw_phones if p]
    # return "\n".join(cleaned[:max_count]) or "none"
    return "none"  # zastąp implementacją


def extract_emails(soup: BeautifulSoup, max_count: int = 3) -> str:
    """
    Wyciąga e-maile do organizacji (maks. max_count).
    Priorytet: sekretariat > biuro > ogólny.
    !! DOSTOSUJ SELEKTOR !!
    """
    # Przykład:
    # raw = [el.get_text(strip=True) for el in soup.select("a[href^='mailto:']")]
    # clean = [e.replace("mailto:", "").strip() for e in raw if "@" in e]
    # # Priorytet: sekretariat > biuro > info > reszta
    # priority = ["sekretariat", "biuro", "info", "kontakt", "office"]
    # sorted_emails = sorted(clean, key=lambda e: next(
    #     (i for i, p in enumerate(priority) if p in e.lower()), len(priority)
    # ))
    # return "\n".join(sorted_emails[:max_count]) or "none"
    return "none"  # zastąp implementacją


def extract_website(soup: BeautifulSoup) -> str:
    """
    Wyciąga URL strony WWW.
    Preferuje podstronę /kontakt lub /contact.
    !! DOSTOSUJ SELEKTOR !!
    """
    # Przykład:
    # link = soup.select_one("a.website-link")
    # if link:
    #     base = link.get("href", "").rstrip("/")
    #     # Sprawdź czy istnieje podstrona kontaktowa
    #     for suffix in ["/kontakt", "/contact", "/kontakty"]:
    #         contact_url = base + suffix
    #         # Opcjonalnie: zweryfikuj istnienie ping-iem (ostrożnie z limitami)
    #         return contact_url
    #     return base
    return "brak"  # zastąp implementacją


def extract_social(soup: BeautifulSoup) -> str:
    """
    Wyciąga URL do profilu w mediach społecznościowych.
    Priorytet: Facebook > LinkedIn > Instagram > inne.
    !! DOSTOSUJ SELEKTOR !!
    """
    # Przykład:
    # priority_domains = ["facebook.com", "linkedin.com", "instagram.com", "twitter.com"]
    # all_links = [a.get("href", "") for a in soup.select("a[href]")]
    # for domain in priority_domains:
    #     for link in all_links:
    #         if domain in link:
    #             return link
    return "brak"  # zastąp implementacją


def extract_description(soup: BeautifulSoup) -> str:
    """
    Wyciąga krótką charakterystykę (1–3 zdania).
    !! DOSTOSUJ SELEKTOR !!
    """
    # Przykład:
    # desc = extract_text(soup, "div.org-description")
    # if desc:
    #     sentences = desc.split(". ")
    #     return ". ".join(sentences[:3]).strip() + ("." if not desc.endswith(".") else "")
    return "none"  # zastąp implementacją


def extract_registry_number(soup: BeautifulSoup, reg_type: str) -> str | None:
    """
    Wyciąga numer z rejestru: krs / regon / nip.
    !! DOSTOSUJ SELEKTOR !!
    """
    # Przykład — szukamy tekstu zawierającego etykietę:
    # label_map = {"krs": "KRS", "regon": "REGON", "nip": "NIP"}
    # label = label_map.get(reg_type, "")
    # for row in soup.select("tr, div.data-row"):
    #     text = row.get_text()
    #     if label in text:
    #         numbers = re.findall(r"\d{6,10}", text)
    #         if numbers:
    #             return numbers[0]
    return None  # zastąp implementacją


def normalize_phone(raw: str) -> str:
    """Normalizuje numer telefonu do formatu XXX XXX XXX lub XX XXX XX XX."""
    digits = re.sub(r"\D", "", raw)
    # Usuń polskie prefiksy (+48, 0048)
    if digits.startswith("0048"):
        digits = digits[4:]
    elif digits.startswith("48") and len(digits) == 11:
        digits = digits[2:]
    # 9 cyfr — komórkowy lub stacjonarny bez kierunkowego: XXX XXX XXX
    if len(digits) == 9:
        return f"{digits[0:3]} {digits[3:6]} {digits[6:9]}"
    # Stacjonarny z kierunkowym: 2 cyfry kierunkowy + 7 cyfr → XX XXX XX XX
    if len(digits) == 9:
        return f"{digits[0:2]} {digits[2:5]} {digits[5:7]} {digits[7:9]}"
    # Inne długości — zwróć cyfry ze spacjami co 3
    if len(digits) >= 7:
        parts = [digits[i:i+3] for i in range(0, len(digits), 3)]
        return " ".join(parts)
    return raw  # zwróć oryginał jeśli nie da się znormalizować


def is_active(last_activity_date: datetime | None) -> bool:
    """Sprawdza czy organizacja była aktywna w ciągu ostatnich N miesięcy."""
    if last_activity_date is None:
        return True  # nie możemy zweryfikować → przyjmujemy że aktywna
    cutoff = datetime.now() - timedelta(days=30 * CONFIG["filter_active_months"])
    return last_activity_date >= cutoff


def is_based_in_poland(record: dict) -> bool:
    """Sprawdza czy siedziba organizacji jest w Polsce."""
    address = (record.get("Adres korespondencyjny") or "").lower()
    voivodeship = (record.get("Województwo") or "").lower()
    # Polskie województwa
    polish_voivodeships = {
        "dolnośląskie", "kujawsko-pomorskie", "lubelskie", "lubuskie",
        "łódzkie", "małopolskie", "mazowieckie", "opolskie", "podkarpackie",
        "podlaskie", "pomorskie", "śląskie", "świętokrzyskie",
        "warmińsko-mazurskie", "wielkopolskie", "zachodniopomorskie",
    }
    if voivodeship in polish_voivodeships:
        return True
    # Fallback: szukaj polskiego kodu pocztowego w adresie (XX-XXX)
    if re.search(r"\d{2}-\d{3}", address):
        return True
    return False


# =============================================================================
# ETAP 1: OFERENCI Z WYNIKÓW KONKURSÓW POLONIJNYCH MSZ/KPRM (dane w PDF)
# =============================================================================
# Selektory i struktura zweryfikowane przez realne pobranie stron i
# załączników (nie zgadywane). Szczegóły patrz historia rozmowy z 05.07.2026.

def pobierz_zalaczniki_pdf(article_url: str, session: requests.Session, archiwum_wayback: bool) -> list[dict]:
    """
    Pobiera stronę artykułu z wynikami konkursu i zwraca listę załączników PDF.

    Każdy element listy to słownik z kluczami: etykieta, url.
    Selektor a.file-download jest identyczny na stronach MSZ i na
    zarchiwizowanych stronach KPRM w Wayback Machine (ten sam szablon CMS gov.pl).
    """
    soup = fetch_page(article_url, session)
    if soup is None:
        log_event("ZALACZNIKI_ARTYKULU", False, f"Nie można pobrać artykułu: {article_url}")
        return []

    baza_url = "https://web.archive.org" if archiwum_wayback else "https://www.gov.pl"

    zalaczniki = []
    for a in soup.select("a.file-download"):
        href = a.get("href", "")
        if not href:
            continue
        etykieta = a.get("aria-label", "") or a.get_text(strip=True)
        etykieta = re.sub(r"\s*Rozmiar.*$", "", etykieta, flags=re.S).strip()
        etykieta = re.sub(r"^Pobierz plik\s*", "", etykieta).strip()
        if archiwum_wayback:
            # Tryb surowy "id_" — zwykły tryb odtwarzania Wayback Machine
            # potrafi zwrócić błąd 500 dla dużych plików binarnych (PDF).
            href = re.sub(r"^(/web/\d{14})/", r"\1id_/", href)
        url_pelny = href if href.startswith("http") else baza_url + href
        zalaczniki.append({"etykieta": etykieta, "url": url_pelny})

    log_event("ZALACZNIKI_ARTYKULU", True, f"{article_url} — znaleziono {len(zalaczniki)} załączników PDF")
    return zalaczniki


def _naglowek_pasuje(naglowek: str, slowa_kluczowe: list[str]) -> bool:
    """Sprawdza, czy nagłówek kolumny (po usunięciu znaków nowej linii) zawiera któreś ze słów kluczowych."""
    znorm = (naglowek or "").replace("\n", " ").strip().lower()
    return any(slowo in znorm for slowo in slowa_kluczowe)


def _znajdz_indeks_kolumny(naglowki: list[str], slowa_kluczowe: list[str]) -> int | None:
    """Zwraca indeks pierwszej kolumny pasującej do słów kluczowych, albo None."""
    for i, naglowek in enumerate(naglowki):
        if _naglowek_pasuje(naglowek, slowa_kluczowe):
            return i
    return None


def _komorka(wiersz: list, idx: int | None) -> str:
    """Bezpiecznie zwraca komórkę o danym indeksie (część wierszy w PDF bywa krótsza niż nagłówek)."""
    if idx is None or idx >= len(wiersz):
        return ""
    return (wiersz[idx] or "").replace("\n", " ").strip()


def wyciagnij_wiersze_z_pdf(zawartosc_pdf: bytes, zrodlo_opis: str) -> list[dict] | None:
    """
    Wyciąga z pliku PDF listę ofert: Oferent, Nazwa projektu, Kwota.

    Układ kolumn różni się między latami (np. 5 kolumn w 2026 roku,
    12 kolumn w 2023 roku), dlatego kolumny znajdujemy po treści nagłówka,
    a nie po stałej pozycji. Niektóre pliki mają dodatkowo wiersz tytułowy
    NAD właściwym nagłówkiem tabeli, dlatego szukamy wiersza nagłówkowego
    w całym dokumencie po dosłownej treści "Oferent", a nie zakładamy, że
    to zawsze pierwszy wiersz pierwszej strony.
    Jeśli nigdzie nie znajdziemy takiej kolumny, zwraca None i loguje
    zdarzenie — NIE zgadujemy układu.
    """
    wszystkie_wiersze = []
    with pdfplumber.open(BytesIO(zawartosc_pdf)) as pdf:
        for strona in pdf.pages:
            tabela = strona.extract_table()
            if tabela:
                wszystkie_wiersze.extend(tabela)

    if not wszystkie_wiersze:
        log_event("PARSOWANIE_PDF", False, f"{zrodlo_opis} — nie wykryto żadnej tabeli")
        return None

    idx_naglowka = next(
        (i for i, w in enumerate(wszystkie_wiersze) if _znajdz_indeks_kolumny(w, ["oferent"]) is not None),
        None,
    )
    if idx_naglowka is None:
        log_event("PARSOWANIE_PDF", False, f"{zrodlo_opis} — nie znaleziono kolumny 'Oferent'")
        return None

    naglowki = wszystkie_wiersze[idx_naglowka]
    dane = wszystkie_wiersze[idx_naglowka + 1:]

    idx_oferent = _znajdz_indeks_kolumny(naglowki, ["oferent"])
    idx_projekt = _znajdz_indeks_kolumny(naglowki, ["nazwa projektu", "tytuł oferty", "tytul oferty"])

    # Niektóre lata mają osobną kolumnę kwoty przyznanej na każdy rok
    # realizacji projektu (np. 2023: kolumny na 2023, 2024 i 2025 rok).
    # Bierzemy WSZYSTKIE takie kolumny, nie tylko pierwszą — inaczej
    # projekt sfinansowany dopiero w kolejnym roku wyglądałby jak odrzucony.
    idx_kwota_lista = [
        i for i, h in enumerate(naglowki) if _naglowek_pasuje(h, ["proponowana", "dotacj"])
    ]
    if not idx_kwota_lista:
        idx_kwota_lista = [i for i, h in enumerate(naglowki) if _naglowek_pasuje(h, ["kwota"])]

    wiersze = []
    for wiersz in dane:
        if not wiersz or not any(wiersz):
            continue
        # Pomiń ewentualne powtórzenie wiersza nagłówkowego na kolejnej stronie.
        if _komorka(wiersz, idx_oferent).lower() == "oferent":
            continue

        oferent = _komorka(wiersz, idx_oferent)
        if not oferent:
            continue
        projekt = _komorka(wiersz, idx_projekt) or "nie ustalono"

        czesci_kwoty = []
        for idx in idx_kwota_lista:
            wartosc = _komorka(wiersz, idx)
            if wartosc:
                etykieta_kolumny = (naglowki[idx] or "").replace("\n", " ").strip() if idx < len(naglowki) else ""
                czesci_kwoty.append(
                    f"{etykieta_kolumny}: {wartosc}" if len(idx_kwota_lista) > 1 else wartosc
                )
        kwota = "; ".join(czesci_kwoty) if czesci_kwoty else "nie ustalono"

        wiersze.append({
            "oferent": oferent,
            "projekt": projekt or "nie ustalono",
            "kwota": kwota or "nie ustalono",
        })

    return wiersze


def zbierz_oferentow_z_konkursow(session: requests.Session, progress: dict):
    """
    Etap 1: dla każdego skonfigurowanego artykułu z wynikami konkursu polonijnego
    pobiera załączniki PDF (pomijając zbiorcze listy "wszystkich zgłoszonych"),
    wyciąga z nich oferentów — zarówno tych, którym przyznano dotację, jak i
    odrzuconych/nieprzyjętych do dofinansowania — i zapisuje do pliku pośredniego.
    """
    seed_wb, seed_ws, seed_row = create_seed_workbook()
    przetworzone = set(progress.get("przetworzone_zalaczniki", []))

    for konkurs in CONFIG["konkursy_polonijne"]:
        zalaczniki = pobierz_zalaczniki_pdf(konkurs["url"], session, konkurs["archiwum_wayback"])
        polite_delay()

        zalaczniki_do_wziecia = [
            z for z in zalaczniki
            if not any(fraza in z["etykieta"].lower() for fraza in CONFIG["zalaczniki_pomijane_frazy"])
        ]

        for zalacznik in tqdm(zalaczniki_do_wziecia, desc=f"[{konkurs['rok']}] {konkurs['nazwa_konkursu'][:30]}", unit="zal"):
            if zalacznik["url"] in przetworzone:
                continue

            try:
                resp = session.get(zalacznik["url"], headers=HEADERS, timeout=CONFIG["request_timeout"])
                resp.raise_for_status()
            except requests.RequestException as e:
                log_event("POBIERANIE_ZALACZNIKA", False, f"{zalacznik['url']} — {e}")
                przetworzone.add(zalacznik["url"])
                progress["przetworzone_zalaczniki"] = list(przetworzone)
                save_progress(progress)
                polite_delay()
                continue

            wiersze = wyciagnij_wiersze_z_pdf(resp.content, zalacznik["url"])

            if wiersze is None:
                log_event(
                    "STRUKTURA_NIEROZPOZNANA", False,
                    f"{zalacznik['url']} ({zalacznik['etykieta']}) — wymaga ręcznej weryfikacji, pominięto",
                )
            else:
                for wiersz in wiersze:
                    seed_ws.cell(row=seed_row, column=1, value=konkurs["rok"])
                    seed_ws.cell(row=seed_row, column=2, value=konkurs["nazwa_konkursu"])
                    seed_ws.cell(row=seed_row, column=3, value=zalacznik["etykieta"])
                    seed_ws.cell(row=seed_row, column=4, value=wiersz["oferent"])
                    seed_ws.cell(row=seed_row, column=5, value=wiersz["projekt"])
                    seed_ws.cell(row=seed_row, column=6, value=wiersz["kwota"])
                    seed_ws.cell(row=seed_row, column=7, value=konkurs["url"])
                    seed_ws.cell(row=seed_row, column=8, value=zalacznik["url"])
                    seed_row += 1
                    seed_wb.save(CONFIG["seed_file"])
                    progress["total_scraped"] += 1
                    save_progress(progress)

                log_event("ZALACZNIK_PRZETWORZONY", True, f"{zalacznik['url']} — {len(wiersze)} ofert")

            przetworzone.add(zalacznik["url"])
            progress["przetworzone_zalaczniki"] = list(przetworzone)
            save_progress(progress)
            polite_delay()

    seed_wb.save(CONFIG["seed_file"])
    log_event("KONIEC_ETAPU_1", True, f"Plik pośredni: {CONFIG['seed_file']}")


def create_seed_workbook() -> tuple:
    """Tworzy nowy plik pośredni etapu 1 albo wczytuje istniejący (dopisywanie, nie nadpisywanie)."""
    seed_path = Path(CONFIG["seed_file"])
    seed_path.parent.mkdir(parents=True, exist_ok=True)

    if seed_path.exists():
        wb = load_workbook(seed_path)
        ws = wb.active
        next_row = ws.max_row + 1
        log_event("WCZYTANIE_PLIKU_POSREDNIEGO", True, f"Kontynuacja od wiersza {next_row}")
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Oferenci — konkursy polonijne"
        for col_idx, col_name in enumerate(SEED_COLUMNS, start=1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.font = Font(bold=True, name="Arial")
            cell.fill = PatternFill("solid", fgColor="D9E1F2")
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        next_row = 2
        log_event("UTWORZENIE_PLIKU_POSREDNIEGO", True, CONFIG["seed_file"])

    return wb, ws, next_row


# =============================================================================
# ZAPIS DO XLSX
# =============================================================================

def create_output_workbook() -> tuple:
    """
    Tworzy nowy plik XLSX z nagłówkami lub wczytuje istniejący.
    Zwraca (workbook, worksheet, następny_wiersz).
    """
    output_path = Path(CONFIG["output_file"])
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if output_path.exists():
        wb = load_workbook(output_path)
        ws = wb.active
        next_row = ws.max_row + 1
        log_event("WCZYTANIE_PLIKU_WYJSCIOWEGO", True, f"Kontynuacja od wiersza {next_row}")
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Baza danych"
        _write_header_row(ws)
        next_row = 2
        log_event("UTWORZENIE_PLIKU_WYJSCIOWEGO", True, CONFIG["output_file"])

    return wb, ws, next_row


def _write_header_row(ws):
    """Zapisuje wiersz nagłówkowy ze stylowaniem."""
    # Kolumna Lp. (wypełnia użytkownik)
    ws["A1"] = "Lp."
    # Pozostałe kolumny
    for col_idx, col_name in enumerate(OUTPUT_COLUMNS, start=2):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = Font(bold=True, name="Arial")
        cell.fill = PatternFill("solid", fgColor="D9E1F2")  # jasny niebieski
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    # Szerokości kolumn (przybliżone)
    column_widths = {
        "A": 6,   # Lp.
        "B": 20,  # Kategoria
        "C": 20,  # Branża / Typ
        "D": 35,  # Nazwa
        "E": 35,  # Adres
        "F": 18,  # Województwo
        "G": 30,  # Telefon
        "H": 30,  # E-mail
        "I": 40,  # Strona WWW
        "J": 35,  # Social media
        "K": 25,  # Osoba kontaktowa
        "L": 25,  # Tel. os. kontaktowej
        "M": 30,  # E-mail os. kontaktowej
        "N": 20,  # Data
        "O": 40,  # Opis
        "P": 50,  # URL źródła
        "Q": 14,  # KRS
        "R": 14,  # REGON
        "S": 14,  # NIP
    }
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width

    ws.row_dimensions[1].height = 30


def write_record(ws, row: int, record: dict):
    """Zapisuje jeden rekord do arkusza."""
    ws.cell(row=row, column=1, value="")  # Lp. — puste, wypełni użytkownik
    for col_idx, col_name in enumerate(OUTPUT_COLUMNS, start=2):
        cell = ws.cell(row=row, column=col_idx, value=record.get(col_name, "none"))
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        cell.font = Font(name="Arial", size=10)

    # Zebry — naprzemienne kolory wierszy dla czytelności
    if row % 2 == 0:
        fill = PatternFill("solid", fgColor="F2F2F2")
        for col in range(1, len(OUTPUT_COLUMNS) + 2):
            ws.cell(row=row, column=col).fill = fill


def create_rejected_workbook() -> tuple:
    """Tworzy lub wczytuje plik odrzuconych rekordów."""
    rejected_path = Path(CONFIG["rejected_file"])
    rejected_path.parent.mkdir(parents=True, exist_ok=True)

    if rejected_path.exists():
        wb = load_workbook(rejected_path)
        ws = wb.active
        next_row = ws.max_row + 1
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Odrzucone"
        # Nagłówki
        for col, header in enumerate(["URL", "Powód odrzucenia", "Czas"], start=1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True, name="Arial")
            cell.fill = PatternFill("solid", fgColor="FFB3B3")  # jasny czerwony
        ws.column_dimensions["A"].width = 60
        ws.column_dimensions["B"].width = 40
        ws.column_dimensions["C"].width = 20
        next_row = 2

    return wb, ws, next_row


def write_rejected(ws, row: int, url: str, reason: str):
    """Zapisuje odrzucony rekord do arkusza odrzuconych."""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.cell(row=row, column=1, value=url)
    ws.cell(row=row, column=2, value=reason)
    ws.cell(row=row, column=3, value=ts)


# =============================================================================
# GŁÓWNA LOGIKA SCRAPOWANIA
# =============================================================================

def scrape_source(
    source: dict,
    session: requests.Session,
    out_ws,
    out_wb,
    rej_ws,
    rej_wb,
    progress: dict,
    current_row: list,   # mutable wrapper żeby unikać nonlocal
    rej_row: list,
):
    """Przetwarza jedno źródło danych od początku do końca."""
    log_event(f"START_ZRODLA_{source['id'].upper()}", True)

    # Pobierz listę URL-i do przetworzenia
    entity_urls = get_entity_urls_from_source(source, session)

    if not entity_urls:
        log_event(f"BRAK_URLS_{source['id'].upper()}", False, "Brak URL-i do przetworzenia")
        return

    # Filtruj już przetworzone (wznowienie po awarii)
    remaining = [u for u in entity_urls if u not in progress["completed_ids"]]
    log.info(f"Źródło '{source['id']}': {len(remaining)}/{len(entity_urls)} URL-i do przetworzenia")

    for url in tqdm(remaining, desc=f"[{source['id']}]", unit="rek"):
        soup = fetch_page(url, session)

        if soup is None:
            write_rejected(rej_ws, rej_row[0], url, "Błąd pobierania strony — 2 próby nieudane")
            rej_row[0] += 1
            rej_wb.save(CONFIG["rejected_file"])
            log_event("REKORD_ODRZUCONY", False, url)
            progress["completed_ids"].append(url)
            save_progress(progress)
            polite_delay()
            continue

        # Parsuj dane
        record = parse_entity_page(url, soup)

        # Sprawdź czy siedziba jest w Polsce
        if not is_based_in_poland(record):
            write_rejected(rej_ws, rej_row[0], url, "Siedziba poza Polską")
            rej_row[0] += 1
            rej_wb.save(CONFIG["rejected_file"])
            log_event("ODRZUCONO_POZA_PL", False, url)
            progress["completed_ids"].append(url)
            save_progress(progress)
            polite_delay()
            continue

        # Zapisz rekord
        write_record(out_ws, current_row[0], record)
        current_row[0] += 1
        out_wb.save(CONFIG["output_file"])

        # Zapisz postęp
        progress["completed_ids"].append(url)
        progress["total_scraped"] += 1
        save_progress(progress)

        log_event(f"ZAPIS_REKORDU {record.get('Nazwa', url)[:40]}", True)
        polite_delay()

    log_event(f"KONIEC_ZRODLA_{source['id'].upper()}", True, f"Przetworzono {len(remaining)} URL-i")


# =============================================================================
# PUNKT WEJŚCIA
# =============================================================================

def main():
    parser = argparse.ArgumentParser(description="Scraper bazy danych organizacji")
    parser.add_argument(
        "--reset",
        action="store_true",
        help="Usuń postęp i zacznij od nowa",
    )
    parser.add_argument(
        "--zbierz-oferentow",
        action="store_true",
        help="Etap 1: zbierz nazwy oferentów z wyników konkursów polonijnych MSZ/KPRM (plik pośredni, nie formatka)",
    )
    args = parser.parse_args()

    if args.reset:
        reset_progress()
        log.info("Postęp zresetowany. Zaczynamy od początku.")

    log_event("START_SKRYPTU", True, f"Zadanie: {CONFIG['task_name']}")

    # Wczytaj postęp (wznowienie lub nowy start)
    progress = load_progress()
    progress.setdefault("przetworzone_zalaczniki", [])

    if args.zbierz_oferentow:
        session = requests.Session()
        try:
            zbierz_oferentow_z_konkursow(session, progress)
        finally:
            session.close()
        return

    # Inicjalizuj pliki wyjściowe
    out_wb, out_ws, next_out_row = create_output_workbook()
    rej_wb, rej_ws, next_rej_row = create_rejected_workbook()

    # Mutable wrappery dla liczników wierszy
    current_row = [next_out_row]
    rej_row = [next_rej_row]

    # Sesja HTTP (wielokrotne użycie połączenia)
    session = requests.Session()

    # Przetwarzaj źródła według priorytetu
    sources_sorted = sorted(CONFIG["sources"], key=lambda s: s["priority"])

    try:
        for source in sources_sorted:
            scrape_source(
                source, session,
                out_ws, out_wb,
                rej_ws, rej_wb,
                progress,
                current_row, rej_row,
            )
    except KeyboardInterrupt:
        log.info("Przerwano przez użytkownika (Ctrl+C). Postęp zapisany.")
    finally:
        # Ostatni zapis
        out_wb.save(CONFIG["output_file"])
        rej_wb.save(CONFIG["rejected_file"])
        session.close()

    # Podsumowanie
    total_scraped = progress["total_scraped"]
    total_rejected = rej_row[0] - 2  # odejmij nagłówek
    print()
    print("=" * 55)
    print(f"  ✅  Pobrano rekordów:    {total_scraped}")
    print(f"  ❌  Odrzucono rekordów:  {max(0, total_rejected)}")
    print(f"  📄  Plik wyjściowy:      {CONFIG['output_file']}")
    print(f"  📋  Odrzucone:           {CONFIG['rejected_file']}")
    print(f"  📝  Logi:                {CONFIG['log_file']}")
    print("=" * 55)

    log_event("KONIEC_SKRYPTU", True, f"Pobrano: {total_scraped}, Odrzucono: {max(0, total_rejected)}")


if __name__ == "__main__":
    main()
