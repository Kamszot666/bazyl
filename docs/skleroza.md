## Mechanizm checkpointu projektu w Claude Code (hasło: „skleroza")

Ten plik jest wersją mechanizmu checkpointu dostosowaną do pracy w Claude Code.
Wersja używana w rozmowach na Claude.ai w przeglądarce opisana jest osobno,
w pliku Neo_memory.txt w Wiedzy tamtego projektu.

Gdy w wiadomości pojawi się słowo skleroza (samodzielnie lub w zdaniu typu
„zrób sklerozę", „czas na sklerozę"), potraktuj to jako polecenie wygenerowania
lub zaktualizowania pliku stanu projektu o nazwie STAN_PROJEKTU.md,
w języku polskim, ściśle według poniższego formatu.

### Gdzie zapisać plik

Zapisz plik jako STAN_PROJEKTU.md w głównym katalogu repozytorium, obok
pliku CLAUDE.md. Zawsze nadpisuj ten sam plik. Nie twórz kolejnych wersji
ani kopii z datą w nazwie.

Plik ma powstać lokalnie natychmiast po zapisaniu, więc mechanizm działa
również wtedy, gdy pracujesz wyłącznie lokalnie, bez połączenia z repozytorium
zdalnym. Jeśli masz w danej chwili dostęp do repozytorium zdalnego, wykonaj
dodatkowo git add, commit z komunikatem opisującym checkpoint, oraz push do
repozytorium Bazyl. Jeśli nie masz w danej chwili dostępu do zdalnego
repozytorium, zapisz plik tylko lokalnie i poinformuj użytkownika wprost,
że commit i push trzeba wykonać później ręcznie.

### Zasada nadrzędna: najważniejsze na górze

Kontekst rozmowy lub sesji bywa ograniczony, więc kolejność sekcji jest
odwrócona względem intuicji: stan, otwarte pytania i cofnięte decyzje idą
NA GÓRĘ, a dosłowny kod NA DÓŁ. Gdyby kontekst miał zostać obcięty, ma
przetrwać to, co najtrudniej odtworzyć z samego kodu — czyli ustalenia
i historia decyzji.

### Wymagany format pliku STAN_PROJEKTU.md (dokładnie te sekcje, w tej kolejności)

1. NAGŁÓWEK — nazwa projektu, data wygenerowania, jednozdaniowy cel
   projektu, oraz jawna informacja: „Ten plik zastępuje poprzedni
   STAN_PROJEKTU.md — to jedyne aktualne źródło stanu."

2. STAN NA TERAZ — punktowo: co jest zrobione i działa, co jest
   w trakcie, co jest następnym krokiem. Maksymalnie zwięźle, bez kodu.

3. OTWARTE PYTANIA / DECYZJE DO PODJĘCIA — każda pozycja: czego
   dotyczy, jakie są warianty, dlaczego nierozstrzygnięte.

4. DECYZJE COFNIĘTE / ODRZUCONE (sekcja krytyczna) — każda pozycja:
   co było pierwotnie, na co zmienione, dlaczego. Przy każdej dopisz
   jawne ostrzeżenie, jeśli stara i nowa wersja wyglądają podobnie
   i mogą się pomylić.

5. USTALENIA ZAIMPLEMENTOWANE — lista decyzji już zapisanych w kodzie,
   każda z formatem: decyzja, jednozdaniowe uzasadnienie, w którym pliku.

6. METADANE ROZMOWY — punkt wejścia (kontynuacja czy od zera),
   chronologiczny zarys przebiegu, charakterystyka stylu pracy istotna
   dla kolejnej sesji.

7. DOSŁOWNY KOD (zawsze na samym dole) — dla każdego pliku napisanego
   lub zmienionego: pełna, finalna treść w bloku kodu, z nazwą pliku
   jako nagłówkiem. Jeśli plik jest niedokończony, oznacz wprost:
   „CZĘŚCIOWY — brakuje X". Czytaj treść plików bezpośrednio z dysku,
   nie odtwarzaj z pamięci rozmowy.

### Zasady jakości (obowiązują przy każdym checkpoincie)

- Nie parafrazuj kodu i nie skracaj go „bo oczywiste". Dosłownie.
- Nie pomijaj plików drobnych (config, stałe, wyjątki, modele).
- Jeśli decyzja została zmieniona więcej niż raz, w sekcji 4 pokaż pełną
  sekwencję zmian, nie tylko stan końcowy.
- Jeśli czegoś nie wiesz na pewno, oznacz to jako niepewne, zamiast zgadywać.
- Po zapisaniu pliku podaj jedno zdanie podsumowania: czy plik zapisano
  tylko lokalnie, czy też wypchnięto do repozytorium Bazyl.
