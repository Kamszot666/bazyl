#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Skrypt do automatycznego uzupełniania bazy danych firmowych
WERSJA 2.0 - z checkpointami i przetwarzaniem wielowątkowym

NOWE FUNKCJE:
- Wszystkie pliki w tym samym folderze
- Automatyczny checkpoint co 10 rekordów
- Przeszukiwanie 10 domen jednocześnie (wielowątkowość)
- Dodatkowa podstrona /kontakt.html
"""

import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
import time
from urllib.parse import urljoin
import openpyxl
from typing import Dict, List, Optional
import logging
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import os

# Konfiguracja logowania
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('scraping_log.txt', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

class CompanyDataScraper:
    """Klasa do scrapowania danych firmowych ze stron internetowych"""
    
    def __init__(self, timeout=10, delay=1.0):
        self.timeout = timeout
        self.delay = delay
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'pl,en-US;q=0.7,en;q=0.3',
        })
        self.stats = {
            'processed': 0,
            'emails_found': 0,
            'contacts_found': 0,
            'social_found': 0,
            'errors': 0
        }
        self.lock = threading.Lock()  # Dla bezpiecznego dostępu do statystyk
    
    def get_page_content(self, url: str) -> Optional[BeautifulSoup]:
        """Pobiera zawartość strony internetowej"""
        try:
            # Dodaj https:// jeśli brak protokołu
            if not url.startswith(('http://', 'https://')):
                url = 'https://' + url
            
            # Usuń podwójne www
            url = url.replace('http://www.www.', 'http://www.')
            url = url.replace('https://www.www.', 'https://www.')
            
            response = self.session.get(url, timeout=self.timeout, allow_redirects=True)
            response.raise_for_status()
            response.encoding = response.apparent_encoding
            return BeautifulSoup(response.text, 'html.parser')
        except requests.exceptions.Timeout:
            logging.debug(f"Timeout dla {url}")
            return None
        except requests.exceptions.RequestException as e:
            logging.debug(f"Błąd pobierania {url}: {str(e)[:100]}")
            return None
        except Exception as e:
            logging.debug(f"Nieoczekiwany błąd dla {url}: {str(e)[:100]}")
            return None
    
    def extract_emails(self, soup: BeautifulSoup) -> List[str]:
        """Wydobywa adresy email ze strony"""
        emails = set()
        
        # Wzorzec email z polskimi znakami
        email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        
        # Szukaj w tekście strony
        page_text = soup.get_text()
        found_emails = re.findall(email_pattern, page_text)
        emails.update(found_emails)
        
        # Szukaj w linkach mailto:
        for link in soup.find_all('a', href=True):
            href = link['href']
            if href.startswith('mailto:'):
                email = href.replace('mailto:', '').split('?')[0].strip()
                if re.match(email_pattern, email):
                    emails.add(email)
        
        # Szukaj w atrybutach data-*
        for elem in soup.find_all(attrs={'data-email': True}):
            email = elem['data-email']
            if re.match(email_pattern, email):
                emails.add(email)
        
        # Filtruj niechciane emaile
        excluded_domains = ['example.com', 'test.com', 'domain.com', 'yoursite.com', 
                           'yourdomain.com', 'sentry.io', 'wix.com']
        excluded_prefixes = ['noreply', 'no-reply', 'donotreply']
        
        emails = {
            e.lower() for e in emails 
            if not any(domain in e.lower() for domain in excluded_domains)
            and not any(e.lower().startswith(prefix) for prefix in excluded_prefixes)
            and len(e) < 100
        }
        
        # Priorytetyzuj emaile kontaktowe
        priority_keywords = ['kontakt', 'biuro', 'sekretariat', 'info', 'office', 
                            'contact', 'recepcja', 'administracja', 'bok']
        priority_emails = [e for e in emails if any(kw in e.lower() for kw in priority_keywords)]
        
        if priority_emails:
            return priority_emails[:3]
        
        return list(emails)[:3] if emails else []
    
    def extract_contact_person(self, soup: BeautifulSoup) -> Optional[str]:
        """Wydobywa imię i nazwisko osoby kontaktowej"""
        contact_keywords = [
            'kontakt', 'osoba do kontaktu', 'osoba kontaktowa', 'kierownik',
            'dyrektor', 'prezes', 'contact person', 'manager', 'director',
            'właściciel', 'zarząd', 'sekretariat', 'biuro', 'recepcja'
        ]
        
        # Wzorzec na polskie imię i nazwisko
        name_pattern = r'\b[A-ZĄĆĘŁŃÓŚŹŻ][a-ząćęłńóśźż]{2,}\s+[A-ZĄĆĘŁŃÓŚŹŻ][a-ząćęłńóśźż]{2,}\b'
        
        # 1. Szukaj w sekcjach kontaktowych
        for section in soup.find_all(['div', 'section', 'article', 'p', 'span']):
            section_text = section.get_text()
            section_text_lower = section_text.lower()
            
            if any(keyword in section_text_lower for keyword in contact_keywords):
                names = re.findall(name_pattern, section_text)
                if names:
                    excluded_names = ['Lorem Ipsum', 'Jan Kowalski', 'Anna Nowak']
                    valid_names = [n for n in names if n not in excluded_names]
                    if valid_names:
                        return valid_names[0]
        
        # 2. Szukaj w nagłówkach
        for header in soup.find_all(['h1', 'h2', 'h3', 'h4']):
            header_text = header.get_text().lower()
            if any(keyword in header_text for keyword in contact_keywords):
                next_elem = header.find_next(['p', 'div', 'span'])
                if next_elem:
                    names = re.findall(name_pattern, next_elem.get_text())
                    if names:
                        return names[0]
        
        # 3. Szukaj wzorca "Osoba: Imię Nazwisko"
        contact_pattern = r'(?:kontakt|osoba|kierownik|dyrektor|prezes)[\s:]+([A-ZĄĆĘŁŃÓŚŹŻ][a-ząćęłńóśźż]+\s+[A-ZĄĆĘŁŃÓŚŹŻ][a-ząćęłńóśźż]+)'
        matches = re.findall(contact_pattern, soup.get_text(), re.IGNORECASE)
        if matches:
            return matches[0]
        
        return None
    
    def extract_social_media(self, soup: BeautifulSoup) -> str:
        """Wydobywa linki do social media i zwraca jako string"""
        social_platforms = {
            'facebook.com': 'Facebook',
            'instagram.com': 'Instagram',
            'linkedin.com': 'LinkedIn',
            'twitter.com': 'Twitter',
            'x.com': 'X',
            'youtube.com': 'YouTube',
            'tiktok.com': 'TikTok'
        }
        
        found_platforms = {}
        
        for link in soup.find_all('a', href=True):
            href = link['href'].lower()
            for platform_url, platform_name in social_platforms.items():
                if platform_url in href and platform_name not in found_platforms:
                    full_url = link['href']
                    if not full_url.startswith('http'):
                        continue
                    found_platforms[platform_name] = full_url
        
        if found_platforms:
            return ', '.join([f"{name}: {url}" for name, url in found_platforms.items()])
        
        return None
    
    def scrape_company_data(self, website_url: str) -> Dict:
        """
        Główna metoda do scrapowania danych firmy
        """
        if pd.isna(website_url) or not website_url or website_url.strip() == '':
            return {'emails': [], 'contact_person': None, 'social_media': None}
        
        all_data = {
            'emails': [],
            'contact_person': None,
            'social_media': None
        }
        
        # Lista URL do sprawdzenia - DODANO /kontakt.html
        urls_to_check = [
            website_url,
            urljoin(website_url, '/kontakt'),
            urljoin(website_url, '/kontakt.html'),  # NOWE!
            urljoin(website_url, '/contact'),
            urljoin(website_url, '/contact.html'),  # NOWE!
            urljoin(website_url, '/o-nas'),
            urljoin(website_url, '/about'),
            urljoin(website_url, '/about-us'),
            urljoin(website_url, '/o-firmie'),
        ]
        
        for url in urls_to_check:
            soup = self.get_page_content(url)
            if not soup:
                continue
            
            # Zbierz emaile
            if not all_data['emails']:
                emails = self.extract_emails(soup)
                if emails:
                    all_data['emails'] = emails
            
            # Zbierz osobę kontaktową
            if not all_data['contact_person']:
                contact = self.extract_contact_person(soup)
                if contact:
                    all_data['contact_person'] = contact
            
            # Zbierz social media
            if not all_data['social_media']:
                social = self.extract_social_media(soup)
                if social:
                    all_data['social_media'] = social
            
            # Jeśli znaleźliśmy wszystko, przerwij
            if all_data['emails'] and all_data['contact_person'] and all_data['social_media']:
                break
            
            time.sleep(self.delay)
        
        return all_data


def save_checkpoint(df: pd.DataFrame, checkpoint_file: str, last_processed_idx: int):
    """Zapisuje checkpoint"""
    try:
        df.to_excel(checkpoint_file, index=False, engine='openpyxl')
        with open('checkpoint.txt', 'w') as f:
            f.write(str(last_processed_idx))
        logging.info(f"💾 Checkpoint zapisany: wiersz {last_processed_idx}")
    except Exception as e:
        logging.error(f"❌ Błąd zapisu checkpointa: {e}")


def load_checkpoint():
    """Wczytuje ostatni checkpoint"""
    try:
        if os.path.exists('checkpoint.txt'):
            with open('checkpoint.txt', 'r') as f:
                last_idx = int(f.read().strip())
            logging.info(f"📂 Znaleziono checkpoint: wznowienie od wiersza {last_idx + 1}")
            return last_idx + 1
        return 0
    except:
        return 0


def process_single_company(row_data: tuple, scraper: CompanyDataScraper, 
                          col_website: str, col_email: str, col_social: str, 
                          col_contact: str, update_existing: bool):
    """
    Przetwarza pojedynczą firmę (funkcja pomocnicza dla wielowątkowości)
    """
    idx, row = row_data
    company_name = row['Nazwa']
    website = row[col_website]
    
    # Sprawdź czy firma ma stronę WWW
    if pd.isna(website) or not website or website.strip() == '':
        return idx, None, "BRAK_WWW"
    
    # Sprawdź czy trzeba uzupełnić dane
    needs_update = False
    if update_existing:
        needs_update = True
    else:
        if pd.isna(row[col_email]) or row[col_email] == '':
            needs_update = True
        elif pd.isna(row[col_contact]) or row[col_contact] == '':
            needs_update = True
        elif pd.isna(row[col_social]) or row[col_social] == '':
            needs_update = True
    
    if not needs_update:
        return idx, None, "KOMPLETNE"
    
    # Scrapuj dane
    try:
        data = scraper.scrape_company_data(website)
        return idx, data, "SUCCESS"
    except Exception as e:
        logging.error(f"  ✗ Błąd [{idx}] {company_name}: {str(e)}")
        return idx, None, f"ERROR: {str(e)}"


def process_database(input_file: str, output_file: str = None, max_rows: int = None, 
                     start_row: int = None, update_existing: bool = False,
                     max_workers: int = 10, checkpoint_interval: int = 10):
    """
    Główna funkcja przetwarzająca bazę danych
    
    Args:
        input_file: Ścieżka do pliku Excel wejściowego (tylko nazwa jeśli w tym samym folderze)
        output_file: Ścieżka do pliku Excel wyjściowego
        max_rows: Maksymalna liczba wierszy do przetworzenia (None = wszystkie)
        start_row: Od którego wiersza zacząć (None = sprawdź checkpoint)
        update_existing: Czy aktualizować istniejące dane
        max_workers: Liczba równoległych wątków (domyślnie 10)
        checkpoint_interval: Co ile rekordów zapisywać checkpoint (domyślnie 10)
    """
    
    logging.info("=" * 80)
    logging.info("START PRZETWARZANIA BAZY DANYCH - WERSJA 2.0")
    logging.info("=" * 80)
    
    # Wczytaj dane
    logging.info(f"Wczytuję plik: {input_file}")
    df = pd.read_excel(input_file)
    logging.info(f"Wczytano {len(df)} firm")
    
    # Sprawdź checkpoint jeśli start_row nie podany
    if start_row is None:
        start_row = load_checkpoint()
    
    # Określ zakres wierszy do przetworzenia
    end_row = min(start_row + max_rows, len(df)) if max_rows else len(df)
    logging.info(f"Przetwarzanie wierszy {start_row} do {end_row}")
    logging.info(f"Liczba równoległych wątków: {max_workers}")
    logging.info(f"Checkpoint co {checkpoint_interval} rekordów")
    
    # Kolumny do uzupełnienia
    COL_WEBSITE = 'Strona WWW'
    COL_EMAIL = 'Adres e-mail'
    COL_SOCIAL = 'Profil w mediach społecznościowych'
    COL_CONTACT = 'Osoba kontaktowa'
    
    # Inicjalizuj scraper
    scraper = CompanyDataScraper()
    
    # Statystyki
    processed_count = 0
    updated_count = 0
    skipped_count = 0
    error_count = 0
    
    # Plik checkpointa
    checkpoint_file = output_file if output_file else input_file.replace('.xlsx', '_checkpoint.xlsx')
    
    # Przetwarzaj w porcjach po checkpoint_interval
    for batch_start in range(start_row, end_row, checkpoint_interval):
        batch_end = min(batch_start + checkpoint_interval, end_row)
        batch_rows = list(df.iloc[batch_start:batch_end].iterrows())
        
        logging.info(f"\n{'='*80}")
        logging.info(f"PORCJA: wiersze {batch_start} - {batch_end}")
        logging.info(f"{'='*80}")
        
        # Przetwarzaj równolegle max_workers firm na raz
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Przygotuj zadania
            futures = {
                executor.submit(
                    process_single_company,
                    row_data,
                    scraper,
                    COL_WEBSITE,
                    COL_EMAIL,
                    COL_SOCIAL,
                    COL_CONTACT,
                    update_existing
                ): row_data for row_data in batch_rows
            }
            
            # Zbieraj wyniki
            for future in as_completed(futures):
                idx, data, status = future.result()
                row = df.iloc[idx]
                company_name = row['Nazwa']
                website = row[COL_WEBSITE]
                
                if status == "BRAK_WWW":
                    logging.debug(f"[{idx}] {company_name} - BRAK STRONY WWW")
                    skipped_count += 1
                    
                elif status == "KOMPLETNE":
                    logging.debug(f"[{idx}] {company_name} - dane kompletne")
                    skipped_count += 1
                    
                elif status == "SUCCESS" and data:
                    logging.info(f"[{idx}] ✓ {company_name}")
                    row_updated = False
                    
                    # Uzupełnij email
                    if data['emails']:
                        if update_existing or pd.isna(row[COL_EMAIL]) or row[COL_EMAIL] == '':
                            df.at[idx, COL_EMAIL] = ', '.join(data['emails'])
                            logging.info(f"      📧 Email: {data['emails'][0]}")
                            row_updated = True
                    
                    # Uzupełnij osobę kontaktową
                    if data['contact_person']:
                        if update_existing or pd.isna(row[COL_CONTACT]) or row[COL_CONTACT] == '':
                            df.at[idx, COL_CONTACT] = data['contact_person']
                            logging.info(f"      👤 Kontakt: {data['contact_person']}")
                            row_updated = True
                    
                    # Uzupełnij social media
                    if data['social_media']:
                        if update_existing or pd.isna(row[COL_SOCIAL]) or row[COL_SOCIAL] == '':
                            df.at[idx, COL_SOCIAL] = data['social_media']
                            logging.info(f"      📱 Social: {data['social_media'][:60]}...")
                            row_updated = True
                    
                    if row_updated:
                        updated_count += 1
                    
                    processed_count += 1
                    
                else:
                    logging.warning(f"[{idx}] {company_name} - {status}")
                    error_count += 1
        
        # Zapisz checkpoint po każdej porcji
        save_checkpoint(df, checkpoint_file, batch_end - 1)
        
        # Progress info
        progress = ((batch_end - start_row) / (end_row - start_row)) * 100
        logging.info(f"\n📊 Progress: {batch_end - start_row}/{end_row - start_row} ({progress:.1f}%)")
        logging.info(f"   Przetworzone: {processed_count} | Zaktualizowane: {updated_count}")
        logging.info(f"   Pominięte: {skipped_count} | Błędy: {error_count}")
    
    # Zapisz wynik końcowy
    if output_file is None:
        output_file = input_file.replace('.xlsx', f'_uzupelniona_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
    
    logging.info("=" * 80)
    logging.info(f"Zapisuję wyniki do: {output_file}")
    df.to_excel(output_file, index=False, engine='openpyxl')
    
    # Usuń plik checkpointa
    try:
        if os.path.exists('checkpoint.txt'):
            os.remove('checkpoint.txt')
        if os.path.exists(checkpoint_file) and checkpoint_file != output_file:
            os.remove(checkpoint_file)
    except:
        pass
    
    # Podsumowanie
    logging.info("=" * 80)
    logging.info("PODSUMOWANIE")
    logging.info("=" * 80)
    logging.info(f"Przetworzone firmy: {processed_count}")
    logging.info(f"Zaktualizowane firmy: {updated_count}")
    logging.info(f"Pominięte firmy: {skipped_count}")
    logging.info(f"Błędy: {error_count}")
    logging.info("=" * 80)
    
    return output_file


if __name__ == "__main__":
    # KONFIGURACJA
    INPUT_FILE = 'baza_firm_final_uzupelniona.xlsx'  # Tylko nazwa - plik w tym samym folderze!
    OUTPUT_FILE = 'baza_firm_uzupelniona_web.xlsx'   # Plik wynikowy
    
    # Parametry przetwarzania
    MAX_ROWS = 50          # Ile firm przetworzyć (None = wszystkie)
    START_ROW = None       # None = sprawdzi checkpoint automatycznie
    UPDATE_EXISTING = False # False = tylko puste pola
    MAX_WORKERS = 10       # Ile domen jednocześnie (NOWE!)
    CHECKPOINT_INTERVAL = 10  # Co ile rekordów checkpoint (NOWE!)
    
    print("""
    ╔═══════════════════════════════════════════════════════════════╗
    ║   SKRYPT DO UZUPEŁNIANIA BAZY DANYCH - WERSJA 2.0           ║
    ║   ✓ Przetwarzanie równoległe (10 domen jednocześnie)        ║
    ║   ✓ Automatyczne checkpointy co 10 rekordów                 ║
    ║   ✓ Dodatkowe podstrony: /kontakt.html, /contact.html       ║
    ╚═══════════════════════════════════════════════════════════════╝
    """)
    
    print(f"\n📋 KONFIGURACJA:")
    print(f"   Plik wejściowy: {INPUT_FILE}")
    print(f"   Plik wyjściowy: {OUTPUT_FILE}")
    print(f"   Liczba firm do przetworzenia: {MAX_ROWS if MAX_ROWS else 'WSZYSTKIE'}")
    print(f"   Równoległych wątków: {MAX_WORKERS}")
    print(f"   Checkpoint co: {CHECKPOINT_INTERVAL} rekordów")
    print(f"   Aktualizuj istniejące: {'TAK' if UPDATE_EXISTING else 'NIE'}")
    print(f"\n🚀 Rozpoczynam przetwarzanie...\n")
    
    # Uruchom przetwarzanie
    output_path = process_database(
        input_file=INPUT_FILE,
        output_file=OUTPUT_FILE,
        max_rows=MAX_ROWS,
        start_row=START_ROW,
        update_existing=UPDATE_EXISTING,
        max_workers=MAX_WORKERS,
        checkpoint_interval=CHECKPOINT_INTERVAL
    )
    
    print(f"\n✅ GOTOWE! Plik zapisany jako: {output_path}")
    print(f"📊 Szczegółowy log dostępny w pliku: scraping_log.txt")
