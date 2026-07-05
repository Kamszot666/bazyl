# Źródła danych kontaktowych polskich firm i organizacji pozarządowych

Dokument roboczy do projektu budowy bazy organizacji. Uporządkowany od źródeł najważniejszych dla Twojego bieżącego celu (organizacje z siedzibą w Polsce wspierające Polonię) do źródeł ogólnych i uwag prawnych. Wszystkie adresy podano w osobnych liniach, do skopiowania.

# Najważniejsza wskazówka dla Twojego projektu

Szukasz organizacji, które mają siedzibę w Polsce i wspierają Polonię oraz Polaków za granicą. Kluczowy wniosek: takich organizacji nie znajdziesz najlepiej w jednym rejestrze branżowym, tylko na listach beneficjentów konkursów polonijnych. W polskim systemie dotacje na pomoc Polonii mogą dostać wyłącznie podmioty zarejestrowane w Polsce, wymienione w ustawie o działalności pożytku publicznego i o wolontariacie, mające doświadczenie w pracy na rzecz Polonii. To dokładnie Twoja grupa docelowa. Dlatego listy wyników tych konkursów to Twoje najcelniejsze i w pełni legalne źródło.

Uwaga na pułapkę: istnieje baza GUS o nazwie „Baza organizacji i instytucji polskich i polonijnych za granicą”. Ona zawiera organizacje działające za granicą, czyli odwrotność tego, czego szukasz. Może się przydać pomocniczo, ale nie jest to lista polskich podmiotów. Opis tej pułapki znajdziesz w dalszej części.

# Grupa pierwsza: konkursy polonijne, czyli najlepsze źródło podmiotów z siedzibą w Polsce

## Kancelaria Prezesa Rady Ministrów, dotacje polonijne

Strona zbiorcza konkursów „Polonia i Polacy za Granicą” z wynikami i listami organizacji, które otrzymały dotacje. Podmioty na tych listach to organizacje zarejestrowane w Polsce.
https://www.gov.pl/web/polonia/dotacje-polonijne

## Senat, Konkurs Polonijny

Historycznie i obecnie ważne źródło. W rozstrzygnięciu z 2025 roku dotacje trafiły do 108 organizacji pozarządowych zarejestrowanych w Polsce, realizujących 334 zadania. Strona zawiera szczegółowe wyniki z nazwami organizacji.
https://www.senat.gov.pl/polonia/konkurs-polonijny-2025/

## Trzej najwięksi operatorzy, od których warto zacząć

Z kontroli Najwyższej Izby Kontroli wynika, że w latach 2020 do 2022 trzy organizacje otrzymały około 90 procent środków przekazywanych w trybie pozakonkursowym. To największe polskie podmioty wspierające Polonię i naturalne punkty startu, bo same prowadzą listy organizacji polonijnych, z którymi współpracują.

Stowarzyszenie Wspólnota Polska.
https://wspolnotapolska.org.pl

Fundacja Pomoc Polakom na Wschodzie.
Wyszukaj po nazwie w KRS oraz w wyszukiwarce internetowej, aby trafić na aktualną stronę.

Fundacja Wolność i Demokracja.
https://wid.org.pl

## Podsumowanie źródeł rządowych o Polonii

Zestawienie instytucji finansujących działania na rzecz Polonii, przydatne do zrozumienia, kto rozdziela środki i gdzie szukać kolejnych list.
https://powroty.gov.pl/instytucje-wspierajace-polonie/

# Grupa druga: portale i rejestry organizacji pozarządowych

## Portal ngo.pl, spis organizacji

Najbardziej rozpoznawalny portal organizacji pozarządowych w Polsce, prowadzony przez Stowarzyszenie Klon/Jawor. Spis pozwala filtrować organizacje między innymi po obszarze działalności międzynarodowej. Dobre miejsce, aby dołożyć organizacje spoza list dotacyjnych.
https://spis.ngo.pl

## Rejestr punkt io

Wyszukiwarka danych z Krajowego Rejestru Sądowego z dostępem programistycznym przez API. Umożliwia wyszukiwanie po nazwie, numerze KRS, NIP i REGON oraz pobieranie danych kontaktowych organizacji. Część danych kontaktowych wymaga płatnego planu.
https://rejestr.io
https://rejestr.io/api

## Rejestr NGO

Wykaz informacji sprawozdawczych o organizacjach pozarządowych, przydatny do weryfikacji kondycji i danych organizacji.
https://rejestr.ngo

# Grupa trzecia: oficjalne rejestry urzędowe i ich API

## Otwarte API Krajowego Rejestru Sądowego

Bezpłatne, oficjalne API Ministerstwa Sprawiedliwości. Zwraca dane odpowiadające odpisom z KRS w formacie JSON. Idealne do pobierania danych rejestrowych fundacji i stowarzyszeń po numerze KRS. Zwróć uwagę, że dane w otwartej usłudze są zanonimizowane w zakresie danych wrażliwych osób reprezentujących podmiot.
https://prs.ms.gov.pl/krs/openApi

## API REGON, usługa BIR udostępniana przez Główny Urząd Statystyczny

Bezpłatne API do rejestru REGON, aktualna wersja to BIR1.1. Wymaga wniosku o klucz produkcyjny wysłanego na adres podany na stronie usługi; dostępny jest też klucz testowy. Ważne ograniczenie: REGON zwraca dane rejestrowe i klasyfikacyjne, ale zwykle nie zawiera marketingowych danych kontaktowych typu e-mail i telefon, więc świetnie nadaje się do weryfikacji i wzbogacania rekordów po numerze NIP, REGON lub KRS, a słabiej do budowania listy od zera.
https://api.stat.gov.pl/Home/RegonApi

Gotowa biblioteka Pythona do obsługi tego API, ułatwia integrację.
https://pypi.org/project/gusregon/

## Uwaga prawna o dostępie do KRS przez usługi sieciowe

Trwają zmiany przepisów dotyczących masowego pobierania danych z KRS. Pojawił się projekt przepisu przewidującego odpowiedzialność karną za nieuprawnione pozyskiwanie informacji z rejestru za pośrednictwem usług sieciowych. W praktyce oznacza to: korzystaj z oficjalnego otwartego API zgodnie z jego regulaminem, a nie z obchodzenia zabezpieczeń wyszukiwarki. Przed większym projektem sprawdź aktualny stan prawny.

# Grupa czwarta: pułapka, o której trzeba wiedzieć

## Baza GUS organizacji polskich i polonijnych za granicą

Nazwa jest myląca w kontekście Twojego projektu. Ta baza dokumentuje organizacje i instytucje działające poza granicami kraju, w około 115 krajach. To nie są podmioty z siedzibą w Polsce. Może posłużyć pomocniczo, na przykład do skojarzenia polskiego operatora z jego zagranicznymi partnerami, ale nie jest listą polskich organizacji.
https://polonia.stat.gov.pl

# Uwagi prawne dotyczące pozyskiwania danych kontaktowych

Poniższe punkty to praktyczne zasady, nie porada prawna. Przy poważnym, komercyjnym lub dużym projekcie skonsultuj się z prawnikiem od ochrony danych.

Pozyskiwanie adresów e-mail i numerów telefonu podlega przepisom RODO, jeśli dane pozwalają zidentyfikować osobę fizyczną. Dotyczy to na przykład adresu w formacie imię kropka nazwisko małpa domena. Fakt, że dane są publiczne, nie znosi obowiązków wynikających z RODO.

Sam scraping nie jest w Polsce zakazany, ale jego legalność zależy od źródła, skali, rodzaju danych i sposobu wykorzystania. Trzeba brać pod uwagę regulamin serwisu, prawo autorskie, ustawę o ochronie baz danych oraz RODO.

Dobre praktyki, które warto stosować i dokumentować: sprawdzenie regulaminu i pliku robots.txt, respektowanie zastrzeżeń serwisu, ograniczenie częstotliwości żądań, pobieranie tylko danych niezbędnych, zapisywanie źródła i daty pozyskania każdego rekordu.

Dla organizacji pozarządowych dane kontaktowe podane oficjalnie jako kontakt instytucji, na przykład ogólny adres biuro małpa domena organizacji, są mniej wrażliwe niż dane osób prywatnych. Preferuj takie kontakty ogólne zamiast danych konkretnych osób.

# Proponowana kolejność działania w projekcie

1. Zbuduj rdzeń listy z wyników konkursów polonijnych Kancelarii Prezesa Rady Ministrów oraz Senatu. To da Ci organizacje z siedzibą w Polsce, dokładnie z Twojej grupy docelowej.
2. Dołóż kontakty ze stron trzech największych operatorów oraz z portalu ngo.pl.
3. Wzbogacaj i weryfikuj rekordy po numerze KRS lub NIP przez otwarte API KRS i API REGON, uzupełniając formę prawną, adres i status.
4. Dane kontaktowe, których brakuje w rejestrach, pozyskuj ze stron internetowych organizacji, odpowiedzialnym scraperem, z zapisem źródła i daty.
5. Scal wszystko w jeden arkusz Excel ze stałym zestawem kolumn i kolumną źródła.
