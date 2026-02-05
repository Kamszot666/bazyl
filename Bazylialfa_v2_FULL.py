#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BAZYLIALFA V2.0 - KOMPLETNY SKRYPT
Zaawansowany Web Scraper do pozyskiwania danych firm
Autor: Claude & User
Data: 2026-02-05
"""

# === IMPORTY PODSTAWOWE ===
import time  # Biblioteka do obsługi czasu i opóźnień
import pandas as pd  # Biblioteka do pracy z plikami Excel i DataFrames
import os  # Biblioteka do operacji na systemie plików (ścieżki, foldery)
import logging  # Biblioteka do zaawansowanego logowania zdarzeń
import random  # Biblioteka do generowania wartości losowych
import re  # Biblioteka do wyrażeń regularnych (wyszukiwanie wzorców)
from datetime import datetime  # Biblioteka do pracy z datami i czasem
import json  # Biblioteka do pracy z formatem JSON (cache)
from pathlib import Path  # Nowoczesna obsługa ścieżek
from typing import Dict, Optional, List, Tuple  # Adnotacje typów dla lepszej czytelności

# === IMPORTY DO WEB SCRAPINGU ===
from bs4 import BeautifulSoup  # Biblioteka do parsowania HTML
from colorama import init, Fore, Style, Back  # Kolorowanie tekstu w terminalu

# === IMPORTY SELENIUM ===
from selenium import webdriver  # Główny moduł Selenium do sterowania przeglądarką
from selenium.webdriver.chrome.service import Service  # Serwis Chrome
from selenium.webdriver.common.by import By  # Lokalizowanie elementów na stronie
from selenium.webdriver.common.keys import Keys  # Symulacja naciśnięć klawiszy
from selenium.webdriver.chrome.options import Options  # Opcje konfiguracyjne Chrome
from webdriver_manager.chrome import ChromeDriverManager  # Automatyczna instalacja sterownika
from selenium.webdriver.support.ui import WebDriverWait  # Czekanie na elementy
from selenium.webdriver.support import expected_conditions as EC  # Warunki oczekiwania
from selenium.common.exceptions import TimeoutException, NoSuchElementException  # Wyjątki Selenium

# === INICJALIZACJA COLORAMA ===
init(autoreset=True)  # Automatyczne resetowanie kolorów po każdym wypisaniu

# === KONFIGURACJA GLOBALNA ===
CONFIG = {
    'FOLDER_ROBOCZY': r"c:\BAZYLIALFA",  # Główny folder roboczy projektu
    'PLIK_WEJSCIOWY': "baza_firm_final_uzupelniona.xlsx",  # Nazwa pliku wejściowego Excel
    'PLIK_WYNIKOWY': "baza_firm_final_uzupelniona_v2.xlsx",  # Nazwa pliku wynikowego
    'PLIK_LOGU': "bazylialfa_v2_log.txt",  # Nazwa pliku z logami
    'PLIK_CACHE': "bazylialfa_cache.json",  # Plik cache'owania wyników
    'CHECKPOINT_INTERVAL': 500,  # Częstotliwość zapisywania checkpointów (co ile rekordów)
    'MAX_WORKERS': 3,  # Maksymalna liczba wątków roboczych
    'REQUEST_TIMEOUT': 15,  # Timeout dla żądań HTTP (w sekundach)
    'MIN_DELAY': 2,  # Minimalne opóźnienie między żądaniami (sekundy)
    'MAX_DELAY': 5,  # Maksymalne opóźnienie między żądaniami (sekundy)
    'USER_AGENTS': [  # Lista User-Agentów do rotacji (imitacja różnych przeglądarek)
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:121.0) Gecko/20100101 Firefox/121.0',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    ]
}

# === TWORZENIE ŚCIEŻEK ===
SCIEZKA_WEJSCIE = os.path.join(CONFIG['FOLDER_ROBOCZY'], CONFIG['PLIK_WEJSCIOWY'])  # Pełna ścieżka do pliku wejściowego
SCIEZKA_WYNIK = os.path.join(CONFIG['FOLDER_ROBOCZY'], CONFIG['PLIK_WYNIKOWY'])  # Pełna ścieżka do pliku wynikowego
SCIEZKA_LOG = os.path.join(CONFIG['FOLDER_ROBOCZY'], CONFIG['PLIK_LOGU'])  # Pełna ścieżka do pliku logu
SCIEZKA_CACHE = os.path.join(CONFIG['FOLDER_ROBOCZY'], CONFIG['PLIK_CACHE'])  # Pełna ścieżka do pliku cache

# === KONFIGURACJA LOGGERA ===
class ColoredFormatter(logging.Formatter):  # Własna klasa formattera z kolorami
    """Formatter dodający kolory do logów w zależności od poziomu."""
    
    COLORS = {  # Mapowanie poziomów logowania na kolory
        'DEBUG': Fore.CYAN,  # Niebieski dla debugowania
        'INFO': Fore.GREEN,  # Zielony dla informacji
        'WARNING': Fore.YELLOW,  # Żółty dla ostrzeżeń
        'ERROR': Fore.RED,  # Czerwony dla błędów
        'CRITICAL': Fore.RED + Back.WHITE  # Czerwony tekst na białym tle dla krytycznych
    }
    
    def format(self, record):  # Metoda formatująca log
        """Formatuje wpis logu z kolorami."""
        log_color = self.COLORS.get(record.levelname, '')  # Pobierz kolor dla poziomu
        record.levelname = f"{log_color}{record.levelname}{Style.RESET_ALL}"  # Dodaj kolor do nazwy poziomu
        return super().format(record)  # Wywołaj oryginalną metodę formatowania

# Konfiguracja loggera
logger = logging.getLogger('BazylialfaV2')  # Utworzenie obiektu loggera
logger.setLevel(logging.DEBUG)  # Ustawienie minimalnego poziomu logowania

# Handler do pliku (bez kolorów)
file_handler = logging.FileHandler(SCIEZKA_LOG, mode='w', encoding='utf-8')  # Zapis do pliku
file_handler.setLevel(logging.DEBUG)  # Poziom logowania dla pliku
file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')  # Format bez kolorów
file_handler.setFormatter(file_formatter)  # Przypisanie formattera

# Handler do konsoli (z kolorami)
console_handler = logging.StreamHandler()  # Wypisywanie na konsolę
console_handler.setLevel(logging.INFO)  # Poziom logowania dla konsoli
console_formatter = ColoredFormatter('%(asctime)s - %(levelname)s - %(message)s')  # Format z kolorami
console_handler.setFormatter(console_formatter)  # Przypisanie formattera

# Dodanie handlerów do loggera
logger.addHandler(file_handler)  # Dodanie handlera pliku
logger.addHandler(console_handler)  # Dodanie handlera konsoli

# === KLASA ZARZĄDZANIA CACHE ===
class CacheManager:  # Klasa do zarządzania pamięcią podręczną
    """Zarządza cache'owaniem wyników wyszukiwania."""
    
    def __init__(self, cache_file: str):  # Konstruktor klasy
        """Inicjalizuje menedżera cache."""
        self.cache_file = cache_file  # Ścieżka do pliku cache
        self.cache = self._load_cache()  # Wczytanie cache z pliku
        logger.info(f"Cache załadowany: {len(self.cache)} wpisów")  # Informacja o liczbie wpisów
    
    def _load_cache(self) -> Dict:  # Metoda wczytująca cache z pliku
        """Wczytuje cache z pliku JSON."""
        if os.path.exists(self.cache_file):  # Sprawdzenie czy plik istnieje
            try:  # Próba wczytania
                with open(self.cache_file, 'r', encoding='utf-8') as f:  # Otwarcie pliku
                    return json.load(f)  # Załadowanie JSON
            except Exception as e:  # Obsługa błędów
                logger.warning(f"Nie można wczytać cache: {e}")  # Ostrzeżenie
                return {}  # Zwróć pusty słownik
        return {}  # Zwróć pusty słownik jeśli plik nie istnieje
    
    def save_cache(self):  # Metoda zapisująca cache do pliku
        """Zapisuje cache do pliku JSON."""
        try:  # Próba zapisu
            with open(self.cache_file, 'w', encoding='utf-8') as f:  # Otwarcie pliku do zapisu
                json.dump(self.cache, f, ensure_ascii=False, indent=2)  # Zapisanie JSON
            logger.debug(f"Cache zapisany: {len(self.cache)} wpisów")  # Informacja debugowania
        except Exception as e:  # Obsługa błędów
            logger.error(f"Błąd zapisu cache: {e}")  # Błąd
    
    def get(self, key: str) -> Optional[Dict]:  # Metoda pobierająca wartość z cache
        """Pobiera wartość z cache."""
        return self.cache.get(key)  # Zwróć wartość dla klucza lub None
    
    def set(self, key: str, value: Dict):  # Metoda zapisująca wartość do cache
        """Zapisuje wartość do cache."""
        self.cache[key] = value  # Przypisanie wartości do klucza
    
    def has(self, key: str) -> bool:  # Metoda sprawdzająca czy klucz istnieje
        """Sprawdza czy klucz istnieje w cache."""
        return key in self.cache  # Zwróć True jeśli istnieje

# === KLASA SELENIUM SCRAPER Z ULEPSZENIAMI ===
class EnhancedSeleniumScraper:  # Ulepszona klasa scrapera
    """Ulepszona klasa do web scrapingu z Selenium."""
    
    def __init__(self, headless: bool = False):  # Konstruktor z opcją trybu headless
        """Inicjalizuje przeglądarkę Chrome z zaawansowanymi opcjami."""
        logger.info("🔧 Konfiguracja przeglądarki Chrome...")  # Informacja o starcie
        
        chrome_options = Options()  # Utworzenie obiektu opcji Chrome
        
        # Ustawienia podstawowe
        chrome_options.add_argument("--start-maximized")  # Maksymalizacja okna
        chrome_options.add_argument("--disable-notifications")  # Wyłączenie powiadomień
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")  # Ukrycie automatyzacji
        chrome_options.add_argument("--disable-dev-shm-usage")  # Optymalizacja pamięci
        chrome_options.add_argument("--no-sandbox")  # Wyłączenie sandboxu (zwiększa wydajność)
        
        # Losowy User-Agent dla każdej sesji
        user_agent = random.choice(CONFIG['USER_AGENTS'])  # Wylosowanie User-Agenta
        chrome_options.add_argument(f"user-agent={user_agent}")  # Ustawienie User-Agenta
        logger.debug(f"User-Agent: {user_agent[:50]}...")  # Log z fragmentem UA
        
        # Tryb headless (opcjonalny)
        if headless:  # Jeśli włączony tryb bez okna
            chrome_options.add_argument("--headless=new")  # Nowy tryb headless
            logger.info("Tryb headless włączony")  # Informacja
        
        # Wyłączenie logowania WebDriver
        chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])  # Wyciszenie logów
        
        # Ustawienia anty-detekcyjne
        chrome_options.add_experimental_option("useAutomationExtension", False)  # Wyłączenie rozszerzenia automation
        chrome_options.add_experimental_option("prefs", {  # Preferencje przeglądarki
            "credentials_enable_service": False,  # Wyłączenie zapisywania haseł
            "profile.password_manager_enabled": False  # Wyłączenie menedżera haseł
        })
        
        # Inicjalizacja sterownika
        try:  # Próba utworzenia sterownika
            self.driver = webdriver.Chrome(  # Utworzenie instancji Chrome
                service=Service(ChromeDriverManager().install()),  # Automatyczna instalacja sterownika
                options=chrome_options  # Przekazanie opcji
            )
            self.wait = WebDriverWait(self.driver, 10)  # Domyślne oczekiwanie 10 sekund
            
            # Ustawienie dodatkowych właściwości JavaScript (anty-detekcja)
            self.driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {  # Wykonanie komendy CDP
                'source': '''
                    Object.defineProperty(navigator, 'webdriver', {
                        get: () => undefined
                    })
                '''  # Nadpisanie właściwości navigator.webdriver
            })
            
            logger.info("✅ Przeglądarka uruchomiona pomyślnie")  # Sukces
        except Exception as e:  # Obsługa błędów
            logger.critical(f"❌ Nie można uruchomić przeglądarki: {e}")  # Krytyczny błąd
            raise  # Ponowne wyrzucenie wyjątku
    
    def human_like_delay(self, min_sec: float = None, max_sec: float = None):  # Metoda symulująca ludzkie opóźnienie
        """Symuluje ludzkie opóźnienie między akcjami."""
        min_delay = min_sec or CONFIG['MIN_DELAY']  # Użyj parametru lub domyślnej wartości
        max_delay = max_sec or CONFIG['MAX_DELAY']  # Użyj parametru lub domyślnej wartości
        delay = random.uniform(min_delay, max_delay)  # Losowe opóźnienie
        logger.debug(f"⏳ Czekam {delay:.2f}s...")  # Log opóźnienia
        time.sleep(delay)  # Właściwe opóźnienie
    
    def accept_cookies(self):  # Metoda akceptująca cookies (RODO)
        """Automatycznie akceptuje cookies/RODO jak człowiek."""
        cookie_patterns = [  # Lista wzorców przycisków cookies
            "//button[contains(translate(text(), 'ACEPTYUJ', 'aceptyuj'), 'akceptuj')]",  # 'Akceptuj'
            "//button[contains(translate(text(), 'ZGADZAM', 'zgadzam'), 'zgadzam')]",  # 'Zgadzam się'
            "//button[contains(translate(text(), 'AGREE', 'agree'), 'agree')]",  # 'Agree'
            "//a[contains(translate(text(), 'ACCEPT', 'accept'), 'accept')]",  # Link 'Accept'
            "//button[contains(@class, 'cookie')]",  # Klasa zawierająca 'cookie'
            "//button[contains(@id, 'cookie')]",  # ID zawierające 'cookie'
            "//button[contains(@class, 'consent')]"  # Klasa zawierająca 'consent'
        ]
        
        for pattern in cookie_patterns:  # Iteracja po wzorcach
            try:  # Próba znalezienia i kliknięcia
                buttons = self.driver.find_elements(By.XPATH, pattern)  # Znalezienie elementów
                for btn in buttons:  # Iteracja po przyciskach
                    if btn.is_displayed() and btn.is_enabled():  # Jeśli widoczny i aktywny
                        # Symulacja ruchu myszy przed kliknięciem
                        self.driver.execute_script("arguments[0].scrollIntoView(true);", btn)  # Przewiń do elementu
                        time.sleep(random.uniform(0.3, 0.7))  # Krótkie opóźnienie
                        self.driver.execute_script("arguments[0].click();", btn)  # Kliknięcie przez JS
                        logger.debug("🍪 Zaakceptowano cookies")  # Log sukcesu
                        time.sleep(random.uniform(0.5, 1.0))  # Opóźnienie po kliknięciu
                        return True  # Sukces
            except Exception:  # Ignorowanie błędów
                continue  # Przejdź do następnego wzorca
        
        logger.debug("🍪 Nie znaleziono przycisków cookies")  # Brak przycisków
        return False  # Niepowodzenie
    
    def expand_all_content(self):  # Metoda rozwijająca całą treść
        """Klika wszystkie przyciski 'czytaj więcej' / 'pokaż więcej'."""
        expand_patterns = [  # Wzorce przycisków rozwijania
            "//a[contains(translate(text(), 'WIĘCEJ', 'więcej'), 'więcej')]",  # Linki 'więcej'
            "//span[contains(translate(text(), 'WIĘCEJ', 'więcej'), 'więcej')]",  # Spany 'więcej'
            "//button[contains(translate(text(), 'WIĘCEJ', 'więcej'), 'więcej')]",  # Przyciski 'więcej'
            "//a[contains(translate(text(), 'POKAŻ', 'pokaż'), 'pokaż')]",  # 'pokaż'
            "//button[contains(translate(text(), 'ROZWIŃ', 'rozwiń'), 'rozwiń')]"  # 'rozwiń'
        ]
        
        expanded_count = 0  # Licznik rozwinięć
        for pattern in expand_patterns:  # Iteracja po wzorcach
            try:  # Próba rozwinięcia
                buttons = self.driver.find_elements(By.XPATH, pattern)  # Znalezienie przycisków
                for btn in buttons[:5]:  # Maksymalnie 5 przycisków (zapobieganie nieskończonym pętlom)
                    if btn.is_displayed():  # Jeśli widoczny
                        try:  # Próba kliknięcia
                            self.driver.execute_script("arguments[0].click();", btn)  # Kliknięcie JS
                            expanded_count += 1  # Zwiększenie licznika
                            time.sleep(random.uniform(0.2, 0.5))  # Krótkie opóźnienie
                        except:  # Ignorowanie błędów pojedynczego kliknięcia
                            pass  # Kontynuuj
            except:  # Ignorowanie błędów
                pass  # Kontynuuj
        
        if expanded_count > 0:  # Jeśli coś rozwinięto
            logger.debug(f"📖 Rozwinięto {expanded_count} sekcji")  # Log
    
    def search_duckduckgo(self, nip: str) -> Optional[str]:  # Wyszukiwanie w DuckDuckGo
        """Wyszukuje firmę w DuckDuckGo i zwraca link do bizraport.pl."""
        try:  # Rozpoczęcie bezpiecznego bloku
            logger.debug(f"🔍 Wyszukuję NIP: {nip} w DuckDuckGo")  # Log
            
            self.driver.get("https://duckduckgo.com/")  # Wejście na stronę DDG
            self.human_like_delay(1, 2)  # Ludzkie opóźnienie
            
            # Znalezienie pola wyszukiwania
            search_box = self.wait.until(  # Czekanie na element
                EC.presence_of_element_located((By.NAME, "q"))  # Lokalizacja po nazwie
            )
            search_box.clear()  # Wyczyszczenie pola
            
            # Wpisywanie zapytania znak po znaku (imitacja człowieka)
            query = f'bizraport {nip}'  # Zapytanie
            for char in query:  # Iteracja po znakach
                search_box.send_keys(char)  # Wpisanie znaku
                time.sleep(random.uniform(0.05, 0.15))  # Krótkie opóźnienie
            
            self.human_like_delay(0.5, 1)  # Opóźnienie przed Enterem
            search_box.send_keys(Keys.RETURN)  # Naciśnięcie Enter
            
            self.human_like_delay(2, 3)  # Czekanie na wyniki
            
            # Szukanie linków do bizraport.pl
            links = self.driver.find_elements(By.CSS_SELECTOR, "a[href*='bizraport.pl']")  # Znalezienie linków
            
            for link in links[:10]:  # Sprawdzenie pierwszych 10 linków
                url = link.get_attribute("href")  # Pobranie URL
                if url and ("/krs/" in url or "/firma/" in url):  # Sprawdzenie czy to profil firmy
                    logger.info(f"✅ Znaleziono: {url}")  # Log sukcesu
                    return url  # Zwróć URL
            
            logger.warning(f"⚠️ Brak wyników dla NIP: {nip}")  # Ostrzeżenie
            return None  # Brak wyniku
            
        except TimeoutException:  # Obsługa timeoutu
            logger.error(f"⏱️ Timeout przy wyszukiwaniu NIP: {nip}")  # Błąd
            return None  # Brak wyniku
        except Exception as e:  # Obsługa innych błędów
            logger.error(f"❌ Błąd wyszukiwania NIP {nip}: {e}")  # Błąd
            return None  # Brak wyniku
    
    def scrape_bizraport(self, url: str) -> Optional[Dict]:  # Główna metoda scrapingu Bizraport
        """Pobiera szczegółowe dane z profilu Bizraport."""
        try:  # Rozpoczęcie bezpiecznego bloku
            logger.debug(f"🌐 Otwieram: {url}")  # Log
            
            self.driver.get(url)  # Wejście na stronę
            self.human_like_delay(2, 3)  # Ludzkie opóźnienie
            
            # Akceptacja cookies
            self.accept_cookies()  # Automatyczna akceptacja
            
            # Rozwinięcie treści
            self.expand_all_content()  # Rozwinięcie sekcji
            
            self.human_like_delay(1, 2)  # Dodatkowe opóźnienie
            
            # Pobranie HTML i parsowanie
            soup = BeautifulSoup(self.driver.page_source, 'html.parser')  # Parsowanie HTML
            full_text = soup.get_text(separator="\n")  # Tekst ze strony
            
            # Inicjalizacja słownika danych
            data = {}  # Pusty słownik
            
            # === FUNKCJE POMOCNICZE ===
            def find_regex(pattern: str, source: str, group: int = 1) -> str:  # Wyszukiwanie regex
                """Wyszukuje wzorzec w tekście używając regex."""
                try:  # Próba wyszukania
                    match = re.search(pattern, source, re.IGNORECASE | re.DOTALL)  # Szukanie
                    if match:  # Jeśli znaleziono
                        return re.sub(r'\s+', ' ', match.group(group).strip())  # Zwróć oczyszczony tekst
                except:  # Obsługa błędów
                    pass  # Ignoruj
                return ""  # Zwróć pusty string
            
            def clean_money(text: str) -> str:  # Czyszczenie wartości finansowych
                """Czyści i formatuje wartości finansowe."""
                if not text:  # Jeśli pusty
                    return ""  # Zwróć pusty
                return text.replace("mln", " mln").replace("tys.", " tys.").strip()  # Formatowanie
            
            # === EKSTRAKCJA DANYCH ===
            
            # Tytu