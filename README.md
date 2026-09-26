# HM Quality Studio 4.0

Makro Pythona do HyperMesh 2023+ (sprawdzane na 2024): jakość siatki, delta REF ↔ INF,
elementy krytyczne z powierzchni, raporty, zrzuty widoków i prezentacje PPTX – jedno okno.

Plik makra: `HM_Quality_Studio.py` (jeden plik, czysty ASCII, PL / EN).

## Uruchomienie

HyperMesh: **File > Run > Python Script…** → `HM_Quality_Studio.py`
(albo w oknie Pythona: `exec(open(r"C:\sciezka\HM_Quality_Studio.py").read())`).
Ponowne uruchomienie przy otwartym oknie zamyka stare okno i otwiera nowe.
Ustawienia zapisują się same w `~/.hm_quality_studio.json`.

## Co nowego w 4.0 (względem 3.0)

### 1. Delta wielu metryk + wyświetlanie
* **Wiele metryk w jednym przebiegu** (Aspect Ratio, Jacobian, Skewness, przesunięcie węzłów):
  REF i INF są wczytywane raz, wszystkie metryki czytane hurtowo; wynik każdej metryki zostaje w pamięci.
* **Widok metryki** („Pokaż na siatce”): elementy trafiają do komponentów pasm tej metryki, komponenty
  pozostałych metryk są puste i **wygaszone**, komponenty **spoza narzędzia są wygaszane** (opcja, domyślnie
  włączona). Na ekranie i na zrzutach zostają tylko pasma bieżącej metryki i elementy „bez zmian”.
* Elementy **„bez zmian” są bezbarwne (białe) i przezroczyste**. Przezroczystość jest ustawiana kilkoma
  sposobami (różne wersje HM: `*setvalue … transparency`, `*transparencyvalue` + `*transparencymark`,
  API Pythona) z odczytem kontrolnym; działający sposób jest zapamiętywany. **Poziom i kolor zmieniasz po
  analizie od razu** (suwak / „Zastosuj na siatce”) – w 3.0 wartość była używana tylko w chwili analizy,
  stąd wrażenie, że „zmiana procentu nie działa”. Gdy żaden sposób nie działa w danej wersji HM, elementy
  zostają białe (Diagnostyka pokazuje, który sposób działa).
* **Zrzuty delty: metryki × widoki** – dla każdej metryki widok na siatce (inne wygaszone) × każdy zaznaczony
  widok → pliki PNG + legenda SVG + slajdy PPTX; „Sprawdź deltę (po ID)” i skala ręczna per metryka.

### 2. Powierzchnie – elementy krytyczne (nowa karta)
* Wskazanie powierzchni: ID (`12 13 20-25`), interaktywnie w HM albo z zaznaczenia HM; alternatywnie
  komponent (np. siatka 2D powierzchni), zestaw lub ID elementów.
* Makro bierze węzły na powierzchniach i elementy 3D (i/lub 2D) mające na nich ≥ N węzłów (3 = cała ściana).
* **Kryteria tolerancji [od, do]** osobne dla tej funkcji (niezależne od progów raportu i delty).
* Wynik: N i % elementów w tolerancji w **REF i INF + delta (punkty procentowe)**, udział elementów
  spełniających wszystkie kryteria, najgorsze elementy, zestawy poza tolerancją; TXT / CSV / XLSX / HTML,
  slajd PPTX.

### 3. Legenda w stylu ANSYS
* Paleta z 24 punktów gradientu (niebieski → cyjan → zielony → żółty → czerwony, jak domyślny paletyzator
  ANSYS), gładkie przejścia także dla wielu pasm; pasek pionowy z ciemnymi przegrodami, podziałkami,
  liczbą elementów w paśmie i znacznikami MIN / MAX (okno, SVG, PPTX).
* Komponenty dostają **dokładny kolor RGB** (`color_rgb`, HM 2021+) albo najbliższy z palety 64 kolorów –
  legenda zawsze pokazuje kolor faktycznie użyty na siatce.

### 4. Automat REF vs INF (nowa karta, F8)
* Dwa pliki + folder wyników (+ opcjonalnie widoki z karty „Widoki”). Automat wczytuje REF, potem INF
  (każdy raz) i na tym samym modelu liczy deltę, raport jakości i analizę powierzchni; następnie robi zrzuty
  delty (metryka × widok), raport porównawczy, raport powierzchni i prezentację (z podglądem slajdów przed
  zapisem).
* Struktura folderu przebiegu:
  ```
  <folder>/<nazwa>[_data_godzina]/
    01_delta/<metryka>/<widok>.png + legenda_<metryka>.svg
    02_raport_jakosci/   raporty _REF, _INF i porównawczy (TXT / HTML / XLSX / CSV)
    03_powierzchnie/     elementy krytyczne z powierzchni (TXT / XLSX / HTML)
    04_prezentacja/      <nazwa>.pptx
    podsumowanie.txt, index.html (galeria zrzutów, tabele, lista plików)
  ```

### 5. Ergonomia
* Pasek boczny z kartami (Start, Automat, Metryki, Delta, Powierzchnie, Raport, Widoki, Prezentacja,
  Podgląd) i narzędziami (Przywróć siatkę, Podgląd na żywo, Diagnostyka, Pomoc); karta Start jako pulpit
  z kartą automatu i stanem każdego kroku.
* Każda karta ma nagłówek z jednym zdaniem „co tu robisz”; rzadkie opcje w zwijanych sekcjach
  „Opcje zaawansowane”; jednolite kolory przycisków (ciemnozielony = analiza, jasnozielony = generuj /
  eksport, niebieski = nawigacja / widoki).
* Skróty: F8 automat, F5 analiza metryk, F6 podgląd na żywo, F7 seria PPTX, 1 zapamiętaj widok, K zrzuty,
  F1 pomoc.

## Testy poza HyperMeshem

`tests/fake_hm.py` to atrapa API HyperMesha (moduły `hm`, `hw`, mini-dyspozytor Tcl), `tests/make_models.py`
generuje syntetyczne modele REF / INF, a `tests/run_smoke.py` uruchamia całość (silniki, automat, okno
z zrzutami kart) bez HyperMesha:

```
pip install PyQt5 Pillow numpy python-pptx XlsxWriter
QT_QPA_PLATFORM=offscreen python3 tests/run_smoke.py /tmp/hmqs_smoke
```

Zrzuty kart i legend trafiają do `/tmp/hmqs_smoke/gui`, wyniki automatu do `/tmp/hmqs_smoke/wyniki`.
Atrapa nie zastępuje testu w prawdziwym HyperMeshu: nazwy poleceń Tcl (`color_rgb`, przezroczystość,
zaznaczenie „by geoms”) różnią się między wersjami – makro próbuje kolejnych wariantów, a **Diagnostyka**
(pasek boczny) pokazuje, które działają w danej instalacji.
