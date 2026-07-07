Narzędzie do jednorazowego budowania baz danych organizacji/firm przez web scraping.
Dane pobierane z Internetu trafiają do pliku XLSX zgodnego z formatką firmy.
Skrypt używa tylko właściciel projektu.
Użytkownik jest osobą niewidomą i pracuje wyłącznie z czytnikami ekranu.

## Priorytety (ważne, w tej kolejności)
1. Poprawność danych
2. Możliwość wznowienia po awarii
3. Łatwość utrzymania
4. Czytelność kodu
5. Wydajność

## Format odpowiedzi asystenta (YOU MUST — użytkownik czyta czytnikiem ekranu)
- Odczyt liniowy: jedna myśl na akapit, krótkie listy zamiast długich bloków.
- Bez ASCII-artu, ramek, banerów i tabel. Zamiast tabeli lista „Nazwa: wartość".
- Prawdziwe, opisowe nagłówki Markdown. Nie pomijaj poziomów nagłówków.
- Nie polegaj na pogrubieniu ani kursywie do przekazania sensu — pisz to słowami.
- Bez emoji i ozdobników. Najpierw wynik lub wniosek, potem uzasadnienie.
- Kod w blokach z nazwą języka. Przed blokiem linia „Początek kodu, język, opis", po bloku linia „Koniec kodu".
- Komentarze pełnymi zdaniami w osobnej linii nad kodem, nie na końcu linii.
- Zmiany w kodzie opisuj słownie, na przykład „w funkcji X zmieniono Y na Z", nie tylko diff.
- Po bloku kodu podaj krótkie, liniowe wyjaśnienie: co zapisać, gdzie, jak uruchomić, jakie biblioteki zainstalować, pełnym poleceniem pip.
- Nazwy zmiennych opisowe. Unikaj nazw jednoliterowych, trudnych do rozróżnienia w mowie.
- Linki jako tekst opisowy plus pełny adres w osobnej linii.
- Zanim stwierdzisz coś o kodzie lub źródle danych, zweryfikuj to. Jeśli czegoś nie wiesz na pewno, powiedz to wprost zamiast zgadywać.

## Środowisko uruchomieniowe
- System: Windows 11 Pro. Ścieżki plików i polecenia w stylu Windows (cmd lub PowerShell).
- Czytniki ekranu użytkownika: NVDA na Windows, VoiceOver na iPhone, TalkBack na Androidzie.
- Kodowanie musi poprawnie obsługiwać polskie znaki na Windows.

## Stos technologiczny
- Python 3.14
- requests, beautifulsoup4, lxml, openpyxl, tqdm (wszystkie biblioteki darmowe dozwolone)
- Kodowanie wyjścia: `utf-8-sig` (UTF-8 z BOM — polskie znaki w Excel)

## Zasady pisania kodu (YOU MUST)
- Preferuj prostotę nad elegancją. Funkcja zamiast klasy, gdy funkcja wystarczy.
- ZAKAZANE: singletony, CacheManager, REVIEW_REQUIRED, klasy bazowe dla retry.
- Konfiguracja w JEDNYM miejscu: blok `CONFIG` na początku `scraper.py`.
  Zmiana zadania = zmiana tylko CONFIG (linki, nazwy plików, filtry).
- Brak danych w komórce → wpisz `brak` (pola ogólne) lub `nie ustalono` (dane osobowe/kontaktowe),
  zgodnie z formatką. Nigdy nie wpisuj angielskiego słowa `none`.
- Daty: format `DD.MM.RRRR`. Telefony: format `XXX XXX XXX`, dla stacjonarnych z numerem kierunkowym.

## Metodyka pozyskiwania danych (YOU MUST)
- Przed napisaniem scrapera sprawdź, czy źródło ma oficjalne API lub gotowy plik CSV/XML/JSON.
  API i pliki urzędowe są stabilniejsze niż scraping HTML.
- Scraping stosuj tam, gdzie API/plik nie pokrywa danych, np. dane osoby kontaktowej.
- Wyciąganie danych kontaktowych ze znanych już stron organizacji: selenium jako podstawowe narzędzie
  (renderuje stronę tak jak zrobiłaby to przeglądarka, więc działa niezależnie od tego, czy strona jest
  statyczna czy ładowana przez JavaScript). requests + BeautifulSoup/lxml zostają dla źródeł w pełni
  statycznych i przewidywalnych, np. API oraz strony z Etapu 1 (gov.pl). Scrapy pomiń — to projekt
  jednorazowy, nie wymaga tej skali.
- Automatyzacja przeglądarki (selenium, Playwright) do pobierania stron jest w pełni dozwolona i pożądana.
  To NIE jest to samo co omijanie zabezpieczeń — patrz punkt niżej o CAPTCHA.
- Wyszukiwanie adresu WWW organizacji: wyłącznie przez narzędzie wyszukiwania asystenta (nie przez
  automatyczne zapytania do stron wyszukiwarek typu DuckDuckGo/Google) — te strony aktywnie blokują
  automatyczne zapytania zagadkami CAPTCHA, a ich omijanie jest zakazane niezależnie od kontekstu,
  patrz „Czego NIE robić” niżej. To ograniczenie dotyczy zabezpieczeń stron trzecich, nie jest tylko
  lokalną regułą tego projektu, więc nie da się go znieść zmianą tego pliku.
- Odpowiedzialny scraping: ustaw nagłówek User-Agent, rozsądne opóźnienia między żądaniami, ogranicz liczbę zapytań.
- Deduplikacja rekordów po numerze KRS.
- Danych niepewnych lub sprzecznych nie zgaduj. Wpisz `nie ustalono` i odnotuj zdarzenie w scraper_log.txt.

## Struktura kolumn wyjściowych (XLSX)
Pomijamy kolumnę Lp. (użytkownik uzupełnia ręcznie). Kolejność:
Kategoria, Branża/Typ, Nazwa, Adres korespondencyjny, Województwo, Numer telefonu,
Adres e-mail, Strona WWW, Profil social media, Osoba kontaktowa, Tel. os. kontaktowej,
E-mail os. kontaktowej, Data pozyskania, Krótka charakterystyka,
plus kolumny dodatkowe: URL źródła, KRS, REGON, NIP.
Pełne definicje kolumn: patrz Procedura_tworzenia_bazy_danych_w_Excelu.

## Reguły danych (YOU MUST)
- Tylko organizacje z siedzibą w POLSCE (filtruj po województwie i kodzie pocztowym XX-XXX).
- Tylko aktywne dziś lub w ciągu ostatnich 12 miesięcy.
- Telefony/e-maile firmy: maks. 3, priorytet sekretariat > biuro > ogólny.
- Strona WWW: preferuj podstronę /kontakt.

## Polityka błędów i wznowień (YOU MUST)
- Postęp zapisuj do `progress.json` po KAŻDYM rekordzie. Wznawiaj od miejsca zatrzymania.
- Błąd sieciowy → 1 ponowna próba po 5 s → jeśli nadal błąd, pomiń rekord.
- Pominięte rekordy → `output/odrzucone.xlsx` z kolumnami: URL, powód, czas.
- Loguj zdarzenia do `scraper_log.txt` w formacie: `zdarzenie | SUCCESS/ERROR | czas`.

## Architektura plików
```
scraper.py        # główny skrypt (CONFIG na górze)
config.py         # opcjonalnie: wydzielony CONFIG
requirements.txt
progress.json     # auto: stan postępu
scraper_log.txt   # auto: logi
output/           # auto: baza_*.xlsx, odrzucone.xlsx
STAN_PROJEKTU.md  # auto: checkpoint pamięci, patrz docs/skleroza.md
```
Szczegóły źródeł danych dla aktualnego zadania: @docs/zrodla_polonia.md

## Mechanizm checkpointu (hasło: skleroza)
Gdy użytkownik napisze słowo skleroza, zastosuj procedurę z: @docs/skleroza.md

## Workflow weryfikacji (YOU MUST)
Po napisaniu lub zmianie kodu ZAWSZE:
1. Uruchom `python -c "import ast; ast.parse(open('scraper.py').read())"` (sprawdź składnię).
2. Testuj parsowanie na 2–3 rekordach PRZED uruchomieniem pełnego przebiegu.
3. Weryfikuj selektory CSS przez realne pobranie strony, nie zgaduj.
4. Sprawdź, czy KRS/REGON/NIP poprawnie wyodrębnione (regexy na realnych danych).

## Czego NIE robić
- Nie pisz kodu na podstawie zgadywanych selektorów CSS — najpierw pobierz stronę i obejrzyj HTML.
- Nie nadpisuj `progress.json`, `scraper_log.txt` ani pliku wyjściowego przy wznowieniu — dopisuj.
- Nie dodawaj abstrakcji „na przyszłość" — to projekt jednorazowy.
- Przy operacjach nieodwracalnych (nadpisanie/usunięcie plików, wymuszony push) najpierw ostrzeż i poczekaj na potwierdzenie.
- Nie twórz narzędzi do omijania zabezpieczeń, CAPTCHA ani logowania.
