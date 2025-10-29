import requests
from bs4 import BeautifulSoup
import re
import random
import time
import urllib3
from openpyxl import Workbook
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Lista różnych User-Agentów
user_agents = [
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.6 Safari/605.1.15',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:121.0) Gecko/20100101 Firefox/121.0',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:121.0) Gecko/20100101 Firefox/121.0',
    'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Firefox/121.0'
    # Możesz dodać więcej User-Agentów
]

def extract_company_info(url):
    """Pobiera i wyciąga informacje o firmie z podanego URL (bardziej elastyczna wersja)."""
    company_data = {
        "Nazwa firmy": None,
        "Ulica": None,
        "Kod pocztowy": None,
        "Miasto": None,
        "Numery telefonu": [],
        "Strona internetowa": None,
        "Branże": []
    }
    try:
        headers = {'User-Agent': random.choice(user_agents)}
        response = requests.get(url, headers=headers, verify=False)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')

        # Próba wyciągnięcia nazwy firmy (szukamy różnych popularnych tagów i klas)
        name_selectors = ['h1[itemprop="name"]', 'h2', 'span.company-name', 'div.company-title']
        for selector in name_selectors:
            name_tag = soup.select_one(selector)
            if name_tag and name_tag.text.strip():
                company_data["Nazwa firmy"] = name_tag.text.strip()
                break

        # Próba wyciągnięcia adresu
        street_tag = soup.find('span', itemprop='streetAddress')
        postal_code_tag = soup.find('span', itemprop='postalCode')
        locality_tag = soup.find('span', itemprop='addressLocality')
        if street_tag:
            company_data["Ulica"] = street_tag.text.strip()
        if postal_code_tag:
            company_data["Kod pocztowy"] = postal_code_tag.text.strip()
        if locality_tag:
            company_data["Miasto"] = locality_tag.text.strip()

        # Próba wyciągnięcia numerów telefonu (szukamy spanów z itemprop="telephone" lub linków tel:)
        phone_selectors = ['span[itemprop="telephone"]', 'a[href^="tel:"]', 'span.phone', 'p.phone']
        for selector in phone_selectors:
            phone_tag = soup.select(selector)
            for phone in phone_tag:
                phone_number = phone.text.strip().replace('tel:', '').replace(' ', '').replace('-', '')
                if len(phone_number) > 8 and phone_number not in company_data["Numery telefonu"]:
                    company_data["Numery telefonu"].append(phone_number)

        # Próba wyciągnięcia strony internetowej (szukamy itemprop="url" lub klas 'www', 'website')
        website_selectors = ['a[itemprop="url"]', 'a.www', 'a.website']
        for selector in website_selectors:
            website_tag = soup.select_one(selector)
            if website_tag and 'href' in website_tag.attrs and website_tag['href'].startswith('http') and website_tag['href'] != url:
                company_data["Strona internetowa"] = website_tag['href']
                break

        # Próba wyciągnięcia branż (szukamy listy z id 'brLst' lub klas zawierających 'category', 'branch')
        industry_selectors = ['ul#brLst li a span', 'div.category a', 'div.branch a', 'span.industry']
        for selector in industry_selectors:
            industry_tags = soup.select(selector)
            for industry_tag in industry_tags:
                industry = industry_tag.text.strip()
                if industry and industry not in company_data["Branże"]:
                    company_data["Branże"].append(industry)
        if not company_data["Branże"]:
            breadcrumb_tags = soup.select('div.breadcrumb a')
            if breadcrumb_tags and len(breadcrumb_tags) > 1:
                company_data["Branże"].append(breadcrumb_tags[-2].text.strip())
            else:
                meta_description = soup.find('meta', {'name': 'description'})
                if meta_description and 'branża' in meta_description['content'].lower():
                    company_data["Branże"].append(meta_description['content'].split('branża')[1].split(',')[0].strip())


        time.sleep(random.uniform(5, 9))

    except requests.exceptions.RequestException as e:
        print(f"Błąd podczas pobierania strony {url}: {e}")
    except Exception as e:
        print(f"Wystąpił inny błąd: {e}")

    return company_data

# Adres URL pierwszej strony z firmami z województwa łódzkiego
base_url_lodzkie = "https://www.baza-firm.com.pl/wojewodztwo/%C5%82%C3%B3dzkie/strona-{}/"
start_page = 501    
num_pages = 502 - 501  # Obliczamy liczbę stron do pobrania
all_firm_links = []
all_companies_data = []

try:
    headers = {'User-Agent': random.choice(user_agents)}

    # Pobierz linki z określonej liczby stron
    for page_num in range(start_page, start_page + num_pages + 1):
        page_url = base_url_lodzkie.format(page_num)
        print(f"Pobieranie linków ze strony: {page_url}")
        response = requests.get(page_url, headers=headers, verify=False)
        response.raise_for_status()
        soup = BeautifulSoup(response.content, 'html.parser')
        all_links = soup.find_all('a', href=True)
        firm_pattern = re.compile(r'/[a-z0-9\-]+/([a-z0-9\-]+)/[a-z0-9\-]+/pl/\d+\.html')
        for link in all_links:
            href = link['href']
            if firm_pattern.search(href):
                if href.startswith("http"):
                    full_url = href
                else:
                    full_url = "https://www.baza-firm.com.pl" + href
                if full_url not in all_firm_links:
                    all_firm_links.append(full_url)
        print(f"Znaleziono {len(all_firm_links)} linków do firm łącznie.")
        time.sleep(random.uniform(1, 3)) # Dodatkowe opóźnienie między stronami

    print(f"\nŁącznie znaleziono {len(all_firm_links)} linków do firm na stronach od {start_page} do {start_page + num_pages}.")

    # Przetwarzamy każdy link do firmy i wyciągamy informacje
    for index, link in enumerate(all_firm_links, start=1):
        print(f"\nPobieranie informacji z [{index}/{len(all_firm_links)}]: {link}")
        company_info = extract_company_info(link)
        all_companies_data.append(company_info)
        print(company_info) # Wyświetlamy pobrane informacje

    # Zapisywanie danych do pliku XLSX
    workbook = Workbook()
    sheet = workbook.active

    # Nagłówki kolumn
    headers = ["Numer pozycji", "Nazwa firmy", "Adres", "Numery telefonów", "", "Adres strony internetowej", "Branże"]
    sheet.append(headers)

    # Wypełnianie danymi
    for index, data in enumerate(all_companies_data, start=1):
        address = f"{data.get('Ulica', '')} {data.get('Kod pocztowy', '')} {data.get('Miasto', '')}".strip()
        row_data = [
            index,
            data.get('Nazwa firmy', ''),
            address,
            '; '.join(data.get('Numery telefonu', [])),
            '',
            data.get('Strona internetowa', ''),
            '; '.join(data.get('Branże', ''))
        ]
        sheet.append(row_data)

    workbook.save(f'dane_firm_strony_{start_page}_do_{start_page + num_pages}.xlsx')
    print(f"\nDane zostały zapisane do pliku dane_firm_strony_{start_page}_do_{start_page + num_pages}.xlsx")

except requests.exceptions.RequestException as e:
    print(f"Wystąpił błąd podczas pobierania strony: {e}")
except Exception as e:
    print(f"Wystąpił inny błąd: {e}")