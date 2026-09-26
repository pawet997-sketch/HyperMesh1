# -*- coding: ascii -*-
# =====================================================================
#  HM QUALITY STUDIO  -  jakosc siatki MES w HyperMesh (Python)
# =====================================================================
#  Jedno narzedzie w miejsce trzech makr Tcl (raport jakosci siatki,
#  delta jakosci REF <-> INF, zrzuty ekranu / widoki) rozszerzone o:
#   - WIELE METRYK NARAZ: Aspect Ratio, Jacobian Ratio, Jacobian Zero,
#     Skewness czytane w jednym przebiegu po elementach,
#   - PODZIAL WG PROGU: elementy w normie -> jedna grupa bazowa
#     (zielona), poza norma -> pasma kolorow (np. AR 4-5 niebieski ...
#     12-15 czerwony), wartosci poza skala -> osobna grupa,
#   - LEGENDA reczna (liczba kolorow, granice, kolory) i automatyczna
#     (przedzialy z wartosci skrajnych i zadanej liczby pasm),
#   - PREZENTACJA POWERPOINT z PODGLADEM SLAJDU (jak w makrze ANSYS
#     SHOTS): karta "Podglad slajdu" tylko do odczytu, okno "Podglad na
#     zywo" (F6) odswiezane po kazdej zmianie opcji, INTERAKTYWNY kadr
#     (przeciaganie obrazu na slajdzie, szybkie dopasowania, zapisane
#     kadry), podglad WSZYSTKICH slajdow przed zapisem (Utworz / Anuluj)
#     i galeria wyeksportowanych slajdow po zapisie; seria metryka x
#     kamera, slajd zbiorczy, widoki, delta, raport - nowa prezentacja
#     albo dopisanie slajdow do istniejacej (np. firmowego szablonu),
#   - RAPORT POROWNAWCZY z ZESTAWIENIEM WEZLOW (model A vs B: wspolne
#     ID, tylko A / tylko B, wezly przesuniete, maks. przesuniecie),
#   - DIAGNOSTYKA: sprawdzenie API, nazw danych HM dla kazdej metryki,
#     zgodnosci odczytu hurtowego z pojedynczym i wzorcow geometrii.
#
#  NOWE W 3.0 (wzgledem 2.1):
#   * przyciski radiowe "Metryka" i "Typ elementow" (delta) oraz
#     "Metryka" i "tryb legendy" (edytor) sa w OSOBNYCH grupach - wybor
#     2D/3D nie odznacza juz metryki (blad Qt: radio w jednym rodzicu),
#   * "Sprawdz delte" (po ID elementu) dziala z danych OSTATNIEJ analizy
#     i nie zalezy od aktualnego stanu okna ani od API atrybutow,
#   * odczyt hurtowy PARUJE wartosci z ID (dataname=id) zamiast zakladac
#     kolejnosc znacznika; nazwy danych HM sa dobierane z listy
#     kandydatow (aspect / aspectratio, skew / skewness ...),
#   * tryb delty wybierany jednym przelacznikiem (delta / jeden model /
#     podzial wg progu) zamiast dwoch nakladajacych sie opcji,
#   * pasek postepu, globalne "Przerwij", skroty F5 / F6 / F7, karta
#     "Start" z przegladem stanu pracy, podpowiedzi przy kontrolkach.
#
#  Srodowisko: Altair HyperMesh 2023+ (API Pythona "hm"); sprawdzone
#  w HyperMesh 2024.0 (Python 3.8, PyQt5 / Qt 5.12, python-pptx,
#  XlsxWriter, Pillow, numpy - wszystko w instalacji Altaira).
#  Wydajnosc: dane elementow i wezlow sa czytane HURTOWO (jedno
#  polecenie Tcl hm_getvalue ... mark= na ceche) - w GUI kilkanascie razy
#  szybciej niz odczyt atrybutow encja po encji przez API Pythona, ktore
#  zostaje jako sciezka zapasowa (tryb wsadowy, brak mostu Tcl).
# =====================================================================

VERSION = "3.0"
APP_TITLE = "HM Quality Studio"

# ---------------------------- JAK URUCHOMIC --------------------------
#  HyperMesh: File > Run > Python Script...  ->  HM_Quality_Studio.py
#  albo w oknie Pythona HyperMesha:
#      exec(open(r"C:\sciezka\HM_Quality_Studio.py").read())
#  Ponowne uruchomienie przy otwartym oknie zamyka stare okno (ustawienia
#  zostaja zapisane) i otwiera nowe - bez mnozenia okien.
#  Bez okna (konsola, testy):  zmienna srodowiska HMQS_NO_GUI=1 - plik
#  tylko definiuje funkcje i klasy (MQ.analyze(), REPORT.analyze() ...).
#  Uwaga hmbatch: tryb wsadowy nie odpowiada na pytanie HyperMesha
#  "model nie zapisany, kontynuowac?" (konczy sesje bledem) - przed
#  wczytaniem plikow REF / INF zapisz zmieniony model. W GUI pytanie
#  jest obslugiwane automatycznie (okno narzedzia pyta wczesniej).

# ---------------------------- MAPA PLIKU -----------------------------
# Kazda sekcja zaczyna sie naglowkiem "# ==== NAZWA ====" - szukaj po
# nazwie. Kolejnosc jak w pliku:
#   IMPORTY I BIBLIOTEKI OPCJONALNE
#   I18N (PL / EN, ASCII + \u)
#   KOMUNIKATY, POSTEP, PRZERWANIE (BUS)
#   NARZEDZIA: LICZBY, KOLORY, PLIKI
#   KLASYFIKACJA ELEMENTOW (CONFIG)
#   DOSTEP DO HYPERMESHA (API hm)
#   GEOMETRIA ELEMENTU: JACOBIAN ZERO, TET COLLAPSE
#   SKALE, PASMA I PALETA HYPERMESHA
#   STAN SIATKI I PRZYWRACANIE
#   WIELE METRYK -> GRUPY KOLOROW
#   DELTA JAKOSCI REF <-> INF
#   RAPORT JAKOSCI SIATKI
#   ZAPIS RAPORTU: TXT / CSV / XLSX / HTML
#   RAPORT POROWNAWCZY REF vs INF
#   OBRAZY: MINIATURY, TLO, KONWERSJE
#   WIDOKI I ZRZUTY EKRANU
#   UKLAD SLAJDU (PRYMITYWY WSPOLNE DLA PODGLADU I PPTX)
#   PREZENTACJA POWERPOINT (.pptx)
#   ELEMENTY PREZENTACJI (SLAJDY Z MODULOW)
#   USTAWIENIA UZYTKOWNIKA
#   GUI: WSPOLNE KONTROLKI I RYSOWANIE
#   GUI: LEGENDA, EDYTOR LEGENDY, PALETA
#   GUI: PODGLAD SLAJDU (PLOTNO, KADR, PODGLAD SERII, NA ZYWO, PO EKSPORCIE)
#   GUI: KARTA "START"
#   GUI: KARTA "METRYKI I GRUPY"
#   GUI: KARTA "DELTA REF / INF"
#   GUI: KARTA "RAPORT JAKOSCI"
#   GUI: KARTA "WIDOKI I ZRZUTY"
#   GUI: KARTA "PREZENTACJA PPTX"
#   GUI: KARTA "PODGLAD SLAJDU"
#   GUI: DIAGNOSTYKA
#   GUI: OKNO GLOWNE
#   POMOC
#   START
# Okno (klasa StudioWindow) - karty: Start, 1 Metryki i grupy, 2 Delta
# REF/INF, 3 Raport jakosci, 4 Widoki i zrzuty, 5 Prezentacja PPTX,
# Podglad slajdu. Skroty: F5 analiza, F6 podglad na zywo, F7 seria PPTX,
# 1 zapamietaj widok, K zrzuty, F1 pomoc.

# ---------------------------- KONFIGURACJA ---------------------------
LANG             = "PL"      # jezyk startowy ("PL" / "EN"); przelacznik
                             # w oknie jest zapamietywany w ustawieniach
SETTINGS_FILE    = ".hm_quality_studio.json"   # w katalogu uzytkownika
WORK_FOLDER      = "hm_quality_studio"         # folder roboczy w %TEMP%
HDR_COLOR        = "#1c5a96"  # pasek naglowka, akcenty
GO_COLOR         = "#5BA314"  # przyciski "generuj / utworz"
RUN_COLOR        = "#2e7d32"  # przyciski "analizuj"
OK_COLOR         = "#1a7a1a"  # pasek stanu: sukces
INFO_COLOR       = "#1c5a96"  # pasek stanu: praca w toku
WARN_COLOR       = "#b35a00"  # pasek stanu: ostrzezenie
ERR_COLOR        = "#c00000"  # pasek stanu: blad
THUMB_W          = 320        # maks. szerokosc miniatury widoku [px]
THUMB_H          = 200        # maks. wysokosc miniatury widoku [px]
PPT_SLIDE_W      = 12192000   # nowa prezentacja 16:9 (EMU = 1/914400 cala)
PPT_SLIDE_H      = 6858000
TOPN             = 10         # ile najgorszych elementow w raporcie
DIFFMAX          = 5000       # max wierszy w arkuszu "Roznice elementow"
PROGRESS_EVERY   = 2000       # co ile elementow odswiezac pasek stanu
CAPTURE_DELAY_MS = 150        # pauza po zmianie widoku, przed zrzutem

# Paleta 64 kolorow HyperMesha 2024 ("hm_winfo entitycolors"). Komponent
# w HM 2024 ma kolor = INDEKS palety (1..64), wiec kazdy kolor pasma jest
# dobierany z tej palety, a legenda / PPTX pokazuja DOKLADNIE ten kolor.
# W sesji GUI paleta jest czytana na zywo (uwzglednia zmiany uzytkownika);
# ta tabela to wartosci domyslne (np. dla trybu wsadowego).
HM_PALETTE = [
    (0, 0, 0), (255, 255, 255), (255, 0, 0), (0, 255, 0),            # 1-4
    (0, 0, 255), (255, 255, 0), (0, 255, 255), (255, 0, 255),        # 5-8
    (204, 204, 204), (185, 185, 185), (163, 163, 163), (140, 140, 140),
    (118, 118, 118), (96, 96, 96), (73, 73, 73), (51, 51, 51),       # 13-16
    (254, 56, 23), (235, 49, 23), (200, 40, 23), (130, 5, 23),       # 17-20
    (49, 111, 255), (44, 102, 236), (35, 85, 199), (21, 49, 126),    # 21-24
    (255, 131, 88), (238, 117, 81), (201, 99, 65), (130, 56, 23),    # 25-28
    (95, 181, 255), (89, 167, 236), (74, 140, 199), (44, 85, 126),   # 29-32
    (254, 53, 138), (235, 49, 127), (199, 40, 105), (129, 5, 65),    # 33-36
    (163, 124, 255), (150, 115, 236), (126, 94, 199), (80, 56, 126), # 37-40
    (255, 189, 185), (240, 175, 170), (203, 146, 142), (131, 31, 88),
    (252, 62, 255), (233, 56, 236), (198, 49, 199), (129, 27, 126),  # 45-48
    (255, 179, 23), (240, 165, 23), (203, 139, 23), (131, 83, 23),   # 49-52
    (97, 254, 110), (90, 236, 100), (78, 200, 82), (53, 125, 44),    # 53-56
    (255, 235, 124), (244, 217, 114), (207, 183, 96), (133, 116, 57),
    (201, 254, 23), (187, 236, 23), (158, 200, 23), (100, 125, 23),  # 61-64
]
# Rampy indeksow palety od "zimnych" do "goracych" - kolejne pasma
# dostaja ROZNE kolory (zamiast kilku pasm w tym samym odcieniu).
RAMP_FULL = (24, 5, 21, 29, 7, 53, 4, 61, 6, 49, 25, 17, 3)   # bez progu
RAMP_BAD  = (5, 21, 29, 7, 6, 49, 25, 17, 3)                  # poza norma
PAL_OK    = 55     # "w normie"           (zielony)
PAL_OVER  = 8      # "poza skala"         (magenta)
PAL_NA    = 11     # "n/d - brak wartosci" (szary)
PAL_COMB  = (55, 6, 49, 3, 45)   # widok zbiorczy: 0, 1, 2, 3, 4+ metryk

# Modul "Wiele metryk": ustawienia domyslne (zmieniane w oknie i zapisywane).
#   thr   - prog "w normie"; dir above = wyzsze gorsze, below = nizsze gorsze
#   mode  - "manual" (granice edges) albo "auto" (nb pasm z wartosci skrajnych)
MQ_DEFAULTS = {
    "ar":   {"thr": 4.0,  "mode": "manual", "nb": 5, "edges": [4, 5, 7, 10, 12, 15]},
    "jac":  {"thr": 0.6,  "mode": "manual", "nb": 5, "edges": [0, 0.2, 0.3, 0.4, 0.5, 0.6]},
    "jz":   {"thr": 0.0,  "mode": "auto",   "nb": 4, "edges": []},
    "skew": {"thr": 60.0, "mode": "manual", "nb": 5, "edges": [60, 65, 70, 75, 80, 90]},
}

# Modul "Delta": punkty gradientu {R G B} od minimum do maksimum skali
# (jak legenda "quality index" HM). Kolor pasma = najblizszy z palety HM.
GRADIENT_STOPS = [
    (0, 0, 205), (0, 80, 255), (0, 170, 255), (0, 240, 240), (0, 230, 120),
    (60, 215, 0), (210, 235, 0), (255, 210, 0), (255, 120, 0), (235, 0, 0),
]
DELTA_GRAY_IDX      = 11   # "bez zmian"
DELTA_UNMATCHED_IDX = 8    # "bez odpowiednika" / "poprawione"

# Raport jakosci: metryki, progi i biny histogramu -> REPORT_METRICS
# w sekcji "RAPORT JAKOSCI SIATKI".


# ================ IMPORTY I BIBLIOTEKI OPCJONALNE ====================
# HyperMesh 2024 ma Pythona 3.8 z bogatym zestawem bibliotek. Kazda
# biblioteka opcjonalna jest importowana ostroznie - brak ktorejs wylacza
# tylko zalezna funkcje (z czytelnym komunikatem), a nie cale narzedzie.
import os
import io
import re
import sys
import json
import math
import time
import bisect
import contextlib
import base64
import shutil
import zipfile
import builtins
import tempfile
import datetime
import traceback

try:
    import numpy as np                  # szybkie przebarwianie tla zrzutow
except Exception:                       # pragma: no cover
    np = None
try:
    from PIL import Image               # obrazy: miniatury, formaty, tlo
except Exception:                       # pragma: no cover
    Image = None
try:
    import pptx                         # zapis prezentacji PowerPoint
    from pptx.util import Emu, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
except Exception:                       # pragma: no cover
    pptx = None
try:
    import xlsxwriter                   # skoroszyty XLSX raportu
except Exception:                       # pragma: no cover
    xlsxwriter = None

QT = None                               # PyQt5 ladowane leniwie (qt())


def qt():
    """Moduly PyQt5 (QtCore, QtGui, QtWidgets) albo None, gdy brak Qt."""
    global QT
    if QT is None:
        try:
            from PyQt5 import QtCore, QtGui, QtWidgets
            QT = (QtCore, QtGui, QtWidgets)
        except Exception:
            QT = False
    return QT or None


# ====================== I18N (PL / EN, ASCII + \u) ====================
# Jeden przelacznik jezyka dla CALEGO narzedzia (okno, legendy, raporty,
# slajdy). Teksty w kodzie pisane sa parami T(pl, en) - bez katalogow
# kluczy, wiec tekst widac dokladnie tam, gdzie jest uzywany.
_LANG = {"lang": "en" if LANG.upper().startswith("EN") else "pl"}


def T(pl, en, *args):
    """Tekst w jezyku interfejsu; opcjonalne argumenty jak dla operatora %."""
    s = en if _LANG["lang"] == "en" else pl
    if args:
        try:
            s = s % args
        except Exception:
            pass
    return s


def lang():
    return _LANG["lang"]


def set_lang(code):
    if code in ("pl", "en"):
        _LANG["lang"] = code


# ================ KOMUNIKATY, POSTEP, PRZERWANIE (BUS) ================
# Silniki (analizy, raporty, eksport) nie znaja okna. Rozmawiaja z nim przez
# BUS: status (pasek stanu), progress (postep + obsluga przycisku Przerwij),
# confirm / info (okna pytan). Bez okna (tryb wsadowy, testy) komunikaty
# ida na konsole, a pytania dostaja odpowiedz domyslna.
class Cancelled(Exception):
    """Przerwanie dlugiej operacji przyciskiem "Przerwij"."""


class Bus(object):
    LEVELS = {"ok": OK_COLOR, "info": INFO_COLOR, "warn": WARN_COLOR, "err": ERR_COLOR}

    def __init__(self):
        self.status_cb = None      # f(tekst, poziom)
        self.log_cb = None         # f(tekst)  - dziennik karty
        self.confirm_cb = None     # f(tytul, tekst) -> bool
        self.info_cb = None        # f(tytul, tekst, poziom)
        self.cancel = False
        self.quiet = False
        self._last_pump = 0.0

    def status(self, msg, level="ok"):
        if self.status_cb is not None:
            try:
                self.status_cb(msg, level)
                return
            except Exception:
                pass
        if not self.quiet:
            try:
                print("[%s] %s" % (APP_TITLE, msg))
            except Exception:
                pass

    def log(self, msg):
        if self.log_cb is not None:
            try:
                self.log_cb(msg)
                return
            except Exception:
                pass
        if not self.quiet:
            try:
                print("[%s] %s" % (APP_TITLE, msg))
            except Exception:
                pass

    def pump(self, force=False):
        """Pozwala oknu odswiezyc sie w trakcie dlugiej petli (max 10x/s)."""
        now = time.time()
        if not force and now - self._last_pump < 0.1:
            return
        self._last_pump = now
        q = qt()
        if q is not None:
            app = q[2].QApplication.instance()
            if app is not None:
                app.processEvents()

    def progress(self, msg):
        """Postep dlugiej operacji; rzuca Cancelled po klikniecu Przerwij."""
        self.status(msg, "info")
        self.pump()
        if self.cancel:
            self.cancel = False
            raise Cancelled(T("Przerwano na \u017cyczenie u\u017cytkownika.", "Stopped by the user."))

    def confirm(self, msg, title=None, default=True):
        if self.confirm_cb is not None:
            try:
                return bool(self.confirm_cb(title or APP_TITLE, msg))
            except Exception:
                pass
        return default

    def info(self, msg, title=None, level="info"):
        if self.info_cb is not None:
            try:
                self.info_cb(title or APP_TITLE, msg, level)
                return
            except Exception:
                pass
        self.status(msg, level)


BUS = Bus()


# ===================== NARZEDZIA: LICZBY, KOLORY, PLIKI ================
def to_float(v, default=None):
    """Liczba z tekstu / wartosci HM; None albo default, gdy sie nie da."""
    if v is None:
        return default
    if isinstance(v, (int, float)):
        f = float(v)
    else:
        try:
            f = float(str(v).strip().replace(",", "."))
        except Exception:
            return default
    if f != f or f in (float("inf"), float("-inf")):
        return default
    return f


def fmt_num(v, dec=3):
    """Liczba do legendy: bez zbednych zer ("4.00" -> "4", "0.50" -> "0.5")."""
    f = to_float(v)
    if f is None:
        return "" if v is None else str(v)
    if f != 0 and (abs(f) >= 1e6 or abs(f) < 1e-4):
        return "%.3g" % f
    s = "%.*f" % (max(0, int(dec)), f)
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    if s in ("-0", ""):
        s = "0"
    return s


def pnum(v):
    """Liczba w raporcie: 4 cyfry znaczace (jak %.4g w Tcl)."""
    f = to_float(v)
    return "-" if f is None else "%.4g" % f


def pe(v):
    f = to_float(v)
    return "-" if f is None else "%g" % f


def sd1(x):
    """Liczba do formatu %+.1f bez "-0.0"."""
    x = round(float(x), 1)
    return 0.0 if x == 0 else x


def pct(n, tot):
    return 100.0 * n / tot if tot else 0.0


def pct_text(n, tot):
    if not tot:
        return str(n)
    return "%d  (%.2f%%)" % (n, pct(n, tot))


def dec_of(edges):
    """Liczba miejsc po przecinku wynikajaca z najmniejszego kroku granic."""
    step = 0.0
    for a, b in zip(edges, edges[1:]):
        d = abs(b - a)
        if d > 0 and (step == 0 or d < step):
            step = d
    return dec_for(step, 3, 0)


def dec_for(step, default=3, lo=1):
    if not step or step <= 0:
        return default
    d = int(math.ceil(-math.log10(step))) + 1
    return max(lo, min(6, d))


def nice_step(x):
    """Krok zaokraglony w gore do 1 / 2 / 2.5 / 5 x 10^n."""
    if not x or x <= 0:
        return 1.0
    p = 10.0 ** math.floor(math.log10(x))
    m = x / p
    for c in (1.0, 2.0, 2.5, 5.0, 10.0):
        if m <= c * (1.0 + 1e-9):
            return c * p
    return 10.0 * p


def nice_ceil(x):
    """Gorna granica skali zaokraglona do 1 / 1.2 / 1.5 / 2 / 2.5 / 3 / 4 / 5 / 6 / 8."""
    if x <= 0:
        return x
    p = 10.0 ** math.floor(math.log10(x))
    m = x / p
    for c in (1.0, 1.2, 1.5, 2.0, 2.5, 3.0, 4.0, 5.0, 6.0, 8.0, 10.0):
        if m <= c + 1e-9:
            return c * p
    return 10.0 * p


def clean10(x):
    """Usuwa szum zmiennoprzecinkowy (0.30000000000000004 -> 0.3)."""
    return float("%.10g" % x)


def sanit_num(v, dec):
    """Liczba do nazwy komponentu: kropka -> p, minus -> m."""
    return ("%.*f" % (dec, v)).replace(".", "p").replace("-", "m").replace("+", "")


def rgb_hex(rgb):
    r, g, b = [int(round(max(0, min(255, c)))) for c in rgb[:3]]
    return "#%02x%02x%02x" % (r, g, b)


def hex6(rgb):
    return rgb_hex(rgb)[1:].upper()


def hex_rgb(h, default=(128, 128, 128)):
    h = str(h or "").strip().lstrip("#")
    if len(h) != 6:
        return default
    try:
        return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except ValueError:
        return default


def grad_at(stops, f):
    """Kolor z listy punktow gradientu dla ulamka f w [0, 1]."""
    n = len(stops)
    if n == 0:
        return (128, 128, 128)
    if n == 1 or f <= 0.0:
        return tuple(stops[0])
    if f >= 1.0:
        return tuple(stops[-1])
    pos = f * (n - 1)
    i = min(int(pos), n - 2)
    t = pos - i
    a, b = stops[i], stops[i + 1]
    return tuple(int(round(a[k] + (b[k] - a[k]) * t)) for k in range(3))


def is_dark(rgb):
    return (0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]) < 128.0


def now_text(fmt="%Y-%m-%d %H:%M"):
    return datetime.datetime.now().strftime(fmt)


def clean_file_name(s, maxlen=80):
    """Nazwa pliku / folderu: znaki niedozwolone -> "_"."""
    s = re.sub(r'[\\/:*?"<>|]', "_", str(s or "").strip())
    s = re.sub(r"\s+", "_", s)
    return (s or "plik")[:maxlen]


def next_index(folder, prefix=""):
    """Nastepny wolny numer pliku <prefix>_NNN (tylko pliki z tym prefiksem -
    np. "raport_2024.html" w folderze nie zmienia numeracji zrzutow)."""
    mx = 0
    try:
        names = os.listdir(folder)
    except OSError:
        names = []
    pat = re.compile(r"^%s_(\d+)$" % re.escape(prefix)) if prefix else re.compile(r"(\d+)$")
    for f in names:
        m = pat.search(os.path.splitext(f)[0])
        if m:
            mx = max(mx, int(m.group(1)))
    return mx + 1


def writable_dir(folder):
    if not folder or not os.path.isdir(folder):
        return False
    t = os.path.join(folder, ".hmqs_wtest_%d" % os.getpid())
    try:
        with open(t, "w") as fh:
            fh.write("x")
        os.remove(t)
        return True
    except Exception:
        return False


def work_dir(sub=""):
    """Folder roboczy narzedzia (klatki, miniatury) w katalogu tymczasowym."""
    base = os.path.join(tempfile.gettempdir(), WORK_FOLDER)
    d = os.path.join(base, sub) if sub else base
    try:
        os.makedirs(d, exist_ok=True)
    except OSError:
        pass
    return d


def cleanup_work(days=3.0):
    """Usuwa stare pliki robocze (klatki, zrzuty do slajdow) z poprzednich
    sesji - zapamietane widoki nie przechodza miedzy sesjami, a pelne
    klatki zajmuja po kilka MB. Pliki mlodsze niz `days` zostaja (moga
    nalezec do innej, rownolegle otwartej sesji HyperMesha)."""
    d = work_dir("frames")
    limit = time.time() - days * 86400.0
    n = 0
    try:
        for f in os.listdir(d):
            p = os.path.join(d, f)
            try:
                if os.path.isfile(p) and os.path.getmtime(p) < limit:
                    os.remove(p)
                    n += 1
            except OSError:
                pass
    except OSError:
        pass
    return n


def tmp_file(ext, prefix="tmp"):
    return os.path.join(work_dir("frames"), "%s_%d_%d.%s" % (
        prefix, os.getpid(), int(time.time() * 1000000) % 10 ** 12, ext.lstrip(".")))


def open_path(path):
    """Otwiera plik / folder programem domyslnym Windows."""
    try:
        os.startfile(path)            # noqa - tylko Windows
        return True
    except Exception:
        return False


def write_text(path, text):
    with io.open(path, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(text)


def backup_file(path, tag="bak"):
    """Kopia <nazwa>.<tag><ext> obok pliku; zwraca sciezke kopii albo ""."""
    root, ext = os.path.splitext(path)
    bak = "%s.%s%s" % (root, tag, ext)
    try:
        shutil.copy2(path, bak)
        return bak
    except Exception:
        return ""


def assign_attrs(obj, d, keys, choices=None):
    """Ustawienia z pliku JSON -> atrybuty obiektu. Przyjmowane sa tylko
    wartosci tego samego rodzaju co domyslne (bool / liczba / tekst / lista)
    i z dozwolonej listy - uszkodzony albo recznie zmieniony plik ustawien
    nie wywroci okna, tylko zostanie czesciowo pominiety."""
    choices = choices or {}
    for k in keys:
        if k not in d:
            continue
        v, cur = d[k], getattr(obj, k)
        if k in choices and v not in choices[k]:
            continue
        if isinstance(cur, bool):
            ok = isinstance(v, bool)
        elif isinstance(cur, int):
            ok = isinstance(v, (int, float)) and not isinstance(v, bool)
            v = int(v) if ok else v
        elif isinstance(cur, float):
            ok = isinstance(v, (int, float)) and not isinstance(v, bool)
            v = float(v) if ok else v
        elif isinstance(cur, str):
            ok = isinstance(v, str)
        elif isinstance(cur, list):
            ok = isinstance(v, list)
        elif isinstance(cur, dict):
            if isinstance(v, dict):
                for kk, vv in v.items():
                    if kk in cur and isinstance(vv, type(cur[kk])):
                        cur[kk] = vv
            continue
        else:
            ok = True
        if ok:
            setattr(obj, k, v)


def file_locked(path):
    """Czy plik jest zablokowany (np. otwarty w PowerPoincie / Excelu)."""
    if not os.path.isfile(path):
        return False
    try:
        with open(path, "ab"):
            pass
        return False
    except OSError:
        return True


# ================== KLASYFIKACJA ELEMENTOW (CONFIG) ===================
# Config HyperMesha: 1xx = 2D (powloki), 2xx = 3D (bryly), 1..99 = 1D.
TYPE_NAMES = {
    103: "tria3", 104: "quad4", 106: "tria6", 108: "quad8",
    204: "tetra4", 206: "penta6", 208: "hexa8", 210: "tetra10",
    212: "penta15", 214: "hexa20", 205: "pyramid5", 213: "pyramid13",
    215: "pyramid5", 216: "pyramid13",
    1: "mass", 2: "rigid", 3: "rbe3", 5: "rigidlink", 11: "bar2", 12: "bar3",
    21: "spring", 31: "gap", 51: "joint", 55: "weld", 60: "plot", 61: "rod", 63: "plot",
}
SHAPES = {103: "tria", 106: "tria", 104: "quad", 108: "quad",
          204: "tet", 210: "tet", 208: "hex", 214: "hex",
          206: "penta", 212: "penta", 205: "pyr", 213: "pyr", 215: "pyr", 216: "pyr"}
CORNERS = {"tria": 3, "quad": 4, "tet": 4, "hex": 8, "penta": 6, "pyr": 5}


def elem_dim(cfg):
    if cfg is None:
        return "other"
    if 200 <= cfg <= 299:
        return "3d"
    if 100 <= cfg <= 199:
        return "2d"
    if 1 <= cfg <= 99:
        return "1d"
    return "other"


def elem_shape(cfg):
    s = SHAPES.get(cfg)
    if s:
        return s
    d = elem_dim(cfg)
    return {"2d": "shell", "3d": "solid", "1d": "line"}.get(d, "other")


def elem_type_name(cfg):
    if cfg is None:
        return "-"
    return TYPE_NAMES.get(cfg, "config%s" % cfg)


# ==================== DOSTEP DO HYPERMESHA (API hm) ====================
# Cala komunikacja z HyperMeshem przechodzi przez obiekt HM (klasa HmApi).
# Reguly sprawdzone na HyperMesh 2024:
#  - hm.Collection trzeba PRZYPISAC do zmiennej przed iteracja; kolekcja
#    tymczasowa w "for x in hm.Collection(...)" potrafi zniknac w trakcie
#    petli i daje pusta liste (robi to za nas metoda items()),
#  - hm.Collection(model, typ, []) z PUSTA lista ID konczy sie bledem -
#    zawsze sprawdzamy liste przed utworzeniem kolekcji,
#  - usuniecie komponentu kasuje tez jego elementy - dlatego po kazdym
#    movemark liczymy elementy w komponentach, zanim cokolwiek usuniemy,
#  - tagi, zestawy (sets) i kolor tla sa tylko w "debug API" (HmModelDebug),
#  - atrybuty elementu (aspect, jacobian, skew, config, nodes ...) czyta sie
#    wprost (e.aspect); zla nazwa daje None, a nie blad ani crash sesji
#    (w odroznieniu od hm_getelemcheckvalues z nieznana nazwa testu),
#  - kazdy odczyt atrybutu encji przez API Pythona kosztuje ~30 us, a lista
#    wezlow elementu ~300 us - dane WIELU elementow czytamy wiec hurtowo:
#    read_elements() / read_nodes() (Tcl w GUI, API Pythona zapasowo),
#  - czego API Pythona nie ma (pelny widok z zoomem "hm_winfo viewmatrix",
#    sciezka otwartego pliku, paleta kolorow, odtwarzanie komend ukladu
#    z command.tcl, kolor tla sceny) idzie przez most Tcl hw.evalTcl -
#    dostepny w sesji GUI; w trybie wsadowym sa odpowiedniki zastepcze.
def attr(obj, name, default=None):
    """Atrybut encji HM albo default (nigdy wyjatek)."""
    try:
        v = getattr(obj, name)
    except Exception:
        return default
    return default if v is None else v


def num_attr(obj, name):
    """Atrybut liczbowy encji HM (float) albo None."""
    try:
        v = getattr(obj, name)
    except Exception:
        return None
    return to_float(v)


# Procedury Tcl do HURTOWEGO odczytu (tylko sesja GUI). Jedno wywolanie
# "hm_getvalue elems mark=1 dataname=..." zwraca wartosci wszystkich
# elementow w kolejnosci hm_getmark. Pomiar HM 2024, 20 333 elementy:
# ~0.02-0.05 s na ceche, wobec ~0.65 s na atrybut i ~6 s na liste wezlow
# przez API Pythona (kazdy odczyt atrybutu encji to ~30 us). Gdy wywolanie
# hurtowe zawiedzie (dana nie dotyczy czesci elementow), petla z catch.
TCL_BULK = r"""
namespace eval ::hmqs { variable ids {} }
proc ::hmqs::ping {} { return hmqs-ok }
proc ::hmqs::mark {type sel} {
    if {$sel eq "add"} {
        if {[llength $::hmqs::ids]} { eval *appendmark $type 1 $::hmqs::ids }
        return [hm_marklength $type 1]
    }
    *clearmark $type 1
    if {$sel eq "ids"} {
        if {[llength $::hmqs::ids]} { eval *createmark $type 1 $::hmqs::ids }
    } elseif {$sel ne "none"} {
        *createmark $type 1 $sel
    }
    return [hm_getmark $type 1]
}
proc ::hmqs::getmark {type} { return [hm_getmark $type 1] }
proc ::hmqs::vals {type dn} {
    set ids [hm_getmark $type 1]
    if {![catch {hm_getvalue $type mark=1 dataname=$dn} out] && [llength $out] == [llength $ids]} {
        return $out
    }
    set out {}
    foreach e $ids {
        if {[catch {hm_getvalue $type id=$e dataname=$dn} v]} { lappend out NA } else { lappend out $v }
    }
    return $out
}
proc ::hmqs::nodes {} {
    set ids [hm_getmark elems 1]
    if {![catch {hm_getvalue elems mark=1 dataname=nodes} out] && [llength $out] == [llength $ids]} {
        return $out
    }
    set out {}
    foreach e $ids { lappend out [hm_nodelist $e] }
    return $out
}
"""

_TCL_TOKEN = re.compile(r"\{([^{}]*)\}|([^\s{}]+)")


def tcl_items(text):
    """Lista Tcl (plaska albo z podlistami w {}) -> lista napisow."""
    return [a if b == "" else b for a, b in _TCL_TOKEN.findall(text or "")]


def tcl_num(tok):
    if tok in ("", "NA"):
        return None
    return to_float(tok)


class ElemData(object):
    """Wynik HmApi.read_elements: listy rownolegle dla WYBRANYCH elementow."""
    def __init__(self):
        self.all_ids = []        # wszystkie elementy zakresu (takze pominiete)
        self.ids = []            # elementy wybrane (filtr keep)
        self.cfg = []            # ich config
        self.vals = {}           # nazwa danej -> lista (float / int / None)
        self.nodes = None        # lista krotek ID wezlow (gdy nodes=True)

    @property
    def skip(self):
        return len(self.all_ids) - len(self.ids)


INT_DATA = ("config", "collector.id", "id")

# Kandydaci nazw danych HM dla kazdej metryki - PIERWSZA dzialajaca wygrywa
# (sprawdzane na elementach modelu; rozne wersje HyperMesha nazywaja te same
# dane inaczej - lista jak w sprawdzonych makrach Tcl raport_jakosci).
DATANAME_CANDIDATES = {
    "aspect": ("aspect", "aspectratio", "aspect_ratio"),
    "jacobian": ("jacobian",),
    "skew": ("skew", "skewness"),
    "warpage": ("warpage", "warp", "warpage_angle"),
    "taper": ("taper",),
    "minangle": ("minangle", "min_angle", "minanglequad", "mininterioranglequad"),
    "maxangle": ("maxangle", "max_angle", "maxanglequad", "maxinterioranglequad"),
    "shortestside": ("shortestside", "minlength", "minlen", "min_len", "minelemlength", "length"),
    "length": ("length", "minlength", "minlen"),
    "tetcollapse": ("tetcollapse", "tet_collapse", "tetcollapseratio"),
    "volume": ("volume",),
    "area": ("area",),
}


class HmApi(object):
    def __init__(self):
        self._hm = None           # modul hm albo False
        self._ent = None
        self._model = None
        self._dbg = None
        self._tcl = None          # hw.evalTcl albo False
        self._bulk = None         # procedury TCL_BULK zaladowane (True/False)
        self.force_python = False # testy: wymus odczyt przez API Pythona
        self._palette = None
        self._names = {}          # nazwa preferowana -> dzialajaca nazwa danej HM
        self._order = {}          # typ encji -> kolejnosc ID w odpowiedziach mark=1
        self.error = ""

    # ---------------------------------------------------------- srodowisko
    def ok(self):
        """Czy API HyperMesha jest dostepne (skrypt uruchomiony w HM)."""
        if self._hm is None:
            try:
                import hm as _hm
                import hm.entities as _ent
                self._hm, self._ent = _hm, _ent
            except Exception as e:
                self._hm = False
                self.error = "%s" % e
        return bool(self._hm)

    @property
    def hm(self):
        self.ok()
        return self._hm

    @property
    def ent(self):
        self.ok()
        return self._ent

    def model(self):
        if self._model is None:
            self._model = self.hm.Model()
        return self._model

    def debug(self):
        """Model "debug API" (tagi, zestawy, tlo) albo None."""
        if self._dbg is None:
            try:
                from hm.mdi.apis import HmModelDebug
                self._dbg = HmModelDebug()
            except Exception:
                self._dbg = False
        return self._dbg or None

    def refresh(self):
        """Po wczytaniu innego pliku - nowe uchwyty modelu."""
        self._model = None
        self._dbg = None
        self._names = {}
        self._order = {}

    @staticmethod
    def status_of(r):
        """(kod, komunikat) z wyniku funkcji API (0 = sukces)."""
        try:
            if isinstance(r, (list, tuple)):
                r = r[0]
            return int(r.status), "%s" % r.message
        except Exception:
            return 0, ""

    def tcl(self, cmd, default=None):
        """Polecenie Tcl HyperMesha (tylko sesja GUI); default gdy brak/blad."""
        if self._tcl is None:
            try:
                import hw as _hw
                self._tcl = getattr(_hw, "evalTcl", None) or False
            except Exception:
                self._tcl = False
        if not self._tcl:
            return default
        try:
            return self._tcl(cmd)
        except Exception:
            return default

    def has_tcl(self):
        self.tcl("expr 1")
        return bool(self._tcl)

    def redraw(self):
        try:
            self.model().hm_redraw()
        except Exception:
            pass
        BUS.pump(True)

    @contextlib.contextmanager
    def quiet(self):
        """Seria operacji na komponentach bez odswiezania Model Browsera
        i grafiki po kazdej z nich (GUI). Zawsze odblokowane w finally;
        brak polecenia w danej wersji HM = zwykle dzialanie."""
        blocked = []
        for cmd in ("hm_blockbrowserupdate", "hm_blockredraw"):
            if self.tcl("%s 1" % cmd) is not None:
                blocked.append(cmd)
        try:
            yield
        finally:
            for cmd in reversed(blocked):
                self.tcl("%s 0" % cmd)

    # ---------------------------------------------------------- kolekcje
    def items(self, cls, ids=None):
        """Lista encji (wszystkie albo o podanych ID) - bezpieczna iteracja."""
        m = self.model()
        if ids is None:
            col = self.hm.Collection(m, cls)
        else:
            ids = [int(i) for i in ids]
            if not ids:
                return []
            col = self.hm.Collection(m, cls, ids)
        return list(col)

    def collection(self, cls, ids):
        ids = [int(i) for i in ids]
        if not ids:
            return None
        return self.hm.Collection(self.model(), cls, ids)

    def count(self, cls):
        try:
            return len(self.hm.Collection(self.model(), cls))
        except Exception:
            return -1

    # ---------------------------------------------------------- elementy
    def elements(self, displayed=False):
        """Elementy modelu: wszystkie albo tylko wyswietlone."""
        if displayed:
            col = self.hm.CollectionByDisplayed(self.model(), self.ent.Element)
            return list(col)
        return self.items(self.ent.Element)

    def element(self, eid):
        return self.ent.Element(self.model(), int(eid))

    def element_ids(self):
        return [e.id for e in self.items(self.ent.Element)]

    def elem_comp_name(self, e, cache=None):
        """Nazwa komponentu elementu (cache: id -> nazwa)."""
        c = attr(e, "collector")
        if c is None:
            return ""
        cid = attr(c, "id")
        if cache is not None and cid in cache:
            return cache[cid]
        nm = attr(c, "name", "")
        if cache is not None:
            cache[cid] = nm
        return nm

    @staticmethod
    def node_xyz(n):
        try:
            return (float(n.x), float(n.y), float(n.z))
        except Exception:
            return None

    def elem_node_ids(self, e):
        ns = attr(e, "nodes")
        if not ns:
            return []
        out = []
        for n in ns:
            try:
                out.append(n.id)
            except Exception:
                pass
        return out

    def elem_centroid(self, eid):
        try:
            e = self.element(eid)
            cx, cy, cz = num_attr(e, "centerx"), num_attr(e, "centery"), num_attr(e, "centerz")
            if None not in (cx, cy, cz):
                return (cx, cy, cz)
            pts = [self.node_xyz(n) for n in (attr(e, "nodes") or [])]
            pts = [p for p in pts if p]
            if pts:
                k = float(len(pts))
                return tuple(sum(p[i] for p in pts) / k for i in range(3))
        except Exception:
            pass
        return None

    # ---------------------------------------------------------- odczyt hurtowy
    def bulk_ok(self):
        """Czy dziala szybki odczyt przez Tcl (sesja GUI z hw.evalTcl)."""
        if self.force_python:
            return False
        if self._bulk is None:
            self._bulk = False
            if self.has_tcl():
                self.tcl(TCL_BULK)
                self._bulk = self.tcl("::hmqs::ping") == "hmqs-ok"
        return self._bulk

    MARK_CHUNK = 100000       # tyle ID na jedno *createmark / *appendmark

    def _tcl_mark(self, etype, sel, ids=None):
        """Znacznik 1 typu etype: wszystkie / wyswietlone / lista ID (w porcjach,
        zeby jedno polecenie Tcl nie mialo milionow argumentow). Zwraca ID."""
        self._order.pop(etype, None)
        if ids is None:
            out = self.tcl("::hmqs::mark %s %s" % (etype, sel))
        else:
            ids = ["%d" % int(i) for i in ids]
            step = max(1, int(self.MARK_CHUNK))
            out = self.tcl("::hmqs::mark %s none" % etype)
            for a in range(0, len(ids), step):
                self.tcl("set ::hmqs::ids {%s}" % " ".join(ids[a:a + step]))
                if self.tcl("::hmqs::mark %s add" % etype) is None:
                    raise RuntimeError("Tcl: *appendmark %s" % etype)
            if ids:
                out = self.tcl("::hmqs::getmark %s" % etype)
            self.tcl("set ::hmqs::ids {}")
        if out is None:
            raise RuntimeError("Tcl: ::hmqs::mark")
        return [int(x) for x in out.split()]

    def _clear_marks(self):
        """Czysci znaczniki Tcl uzywane przez odczyt hurtowy. WAZNE (sprawdzone
        w HM 2024): operacje API Pythona na kolekcjach (movemark, deletemark)
        DOLICZAJA encje pozostawione w znaczniku przez Tcl - bez czyszczenia
        przeniosly / usunely by encje spoza podanej listy."""
        self._order = {}
        if self._bulk:
            self.tcl("foreach t {elems nodes comps} { foreach k {1 2} { catch {*clearmark $t $k} } }")

    def _tcl_order(self, etype, ids):
        """Kolejnosc encji w odpowiedziach "hm_getvalue ... mark=1": czytana
        z danej "id" zamiast ZAKLADANA - kazda wartosc jest potem parowana
        z ID po ID, wiec inna kolejnosc wewnetrzna dwoch modeli (REF / INF)
        nie moze pomieszac wartosci elementow."""
        order = self._order.get(etype)
        if order is None:
            toks = tcl_items(self.tcl("::hmqs::vals %s id" % etype))
            order = []
            for t in toks:
                f = tcl_num(t)
                order.append(None if f is None else int(f))
            if len(order) != len(ids) or None in order or set(order) != set(ids):
                order = list(ids)              # brak danej "id" -> kolejnosc znacznika
            self._order[etype] = order
        return order

    def _tcl_vals(self, etype, name, ids):
        """Wartosci danej `name` dla encji ZNACZNIKA 1 - lista w kolejnosci
        listy `ids` (ta sama, ktora zostala oznaczona)."""
        ids = list(ids)
        toks = tcl_items(self.tcl("::hmqs::vals %s %s" % (etype, name)))
        if len(toks) != len(ids):
            raise RuntimeError("Tcl: %s (%d != %d)" % (name, len(toks), len(ids)))
        if name in INT_DATA:
            vals = [None if v is None else int(v) for v in (tcl_num(t) for t in toks)]
        else:
            vals = [tcl_num(t) for t in toks]
        order = self._tcl_order(etype, ids)
        if order == ids:
            return vals
        by_id = dict(zip(order, vals))
        return [by_id.get(i) for i in ids]

    def _tcl_nodes(self, ids):
        """Krotki ID wezlow elementow znacznika 1 - w kolejnosci `ids`."""
        ids = list(ids)
        toks = tcl_items(self.tcl("::hmqs::nodes"))
        if len(toks) != len(ids):
            raise RuntimeError("Tcl: nodes (%d != %d)" % (len(toks), len(ids)))
        vals = [tuple(int(float(x)) for x in t.split()) for t in toks]
        order = self._tcl_order("elems", ids)
        if order == ids:
            return vals
        by_id = dict(zip(order, vals))
        return [by_id.get(i, ()) for i in ids]

    def _probe_value(self, eid, dn):
        """Wartosc danej `dn` jednego elementu (Tcl albo API Pythona) albo None."""
        if self.bulk_ok():
            s = self.tcl("if {[catch {hm_getvalue elems id=%d dataname=%s} v]} {return NA} else {return $v}" % (int(eid), dn))
            return tcl_num((s or "NA").strip())
        try:
            return num_attr(self.element(eid), dn)
        except Exception:
            return None

    def resolve_name(self, name, sample_ids):
        """Nazwa danej HM, ktora NAPRAWDE dziala dla nazwy preferowanej `name`
        (kandydaci z DATANAME_CANDIDATES sprawdzani na probce elementow).
        Wynik jest zapamietywany do zmiany modelu."""
        if name in self._names:
            return self._names[name]
        cands = DATANAME_CANDIDATES.get(name, (name,))
        found = None
        for c in cands:
            for eid in list(sample_ids)[:6]:
                if self._probe_value(eid, c) is not None:
                    found = c
                    break
            if found:
                break
        self._names[name] = found or cands[0]
        if found and found != name:
            BUS.log("dataname: %s -> %s" % (name, found))
        elif not found and sample_ids:
            BUS.log("dataname: %s - %s" % (name, T("brak warto\u015bci na pr\u00f3bce element\u00f3w", "no value on the element sample")))
        return self._names[name]

    def read_elements(self, displayed=False, ids=None, keep=None, names=(), nodes=False):
        """Hurtowy odczyt elementow -> ElemData.
        Zakres: ids (lista ID) albo wszystkie / wyswietlone. keep(config) -
        filtr elementow, dla ktorych czytane sa dane. names: lista nazw danych
        HM (aspect, jacobian, collector.id ...) albo funkcja config -> nazwy
        (czytane tylko tam, gdzie maja sens). nodes=True: krotki ID wezlow.
        Szybka sciezka Tcl w GUI, zapasowa - API Pythona (hmbatch, testy)."""
        if self.bulk_ok():
            try:
                return self._read_elements_tcl(displayed, ids, keep, names, nodes)
            except Exception as e:
                BUS.log("Tcl bulk read -> Python API: %s" % e)
        return self._read_elements_py(displayed, ids, keep, names, nodes)

    @staticmethod
    def _names_for(names):
        return names if callable(names) else (lambda cfg, _n=tuple(names): _n)

    def _read_elements_tcl(self, displayed, ids, keep, names, nodes):
        try:
            return self._read_elements_tcl_body(displayed, ids, keep, names, nodes)
        finally:
            self._clear_marks()

    def _read_elements_tcl_body(self, displayed, ids, keep, names, nodes):
        D = ElemData()
        D.all_ids = self._tcl_mark("elems", "displayed" if displayed else "all", ids)
        cfg = [c or 0 for c in self._tcl_vals("elems", "config", D.all_ids)] if D.all_ids else []
        sel = [i for i, c in enumerate(cfg) if keep is None or keep(c)]
        D.ids = [D.all_ids[i] for i in sel]
        D.cfg = [cfg[i] for i in sel]
        if not D.ids:
            D.nodes = [] if nodes else None
            return D
        marked = D.ids == D.all_ids
        need = {}
        nf = self._names_for(names)
        for i, c in enumerate(D.cfg):
            for nm in nf(c):
                need.setdefault(nm, []).append(i)
        for k, (nm, idx) in enumerate(sorted(need.items())):
            BUS.progress(T("Odczyt danych: %s (%d / %d)", "Reading data: %s (%d / %d)", nm, k + 1, len(need)))
            sub = [D.ids[i] for i in idx]
            if len(idx) != len(D.ids) or not marked:
                self._tcl_mark("elems", "ids", sub)
                marked = len(idx) == len(D.ids)
            dn = nm if nm in INT_DATA or "." in nm else self.resolve_name(nm, sub)
            vs = self._tcl_vals("elems", dn, sub)
            if len(idx) == len(D.ids):
                D.vals[nm] = vs
            else:
                full = [None] * len(D.ids)
                for i, v in zip(idx, vs):
                    full[i] = v
                D.vals[nm] = full
        if nodes:
            if not marked:
                self._tcl_mark("elems", "ids", D.ids)
                marked = True
            D.nodes = self._tcl_nodes(D.ids)
        return D

    def _read_elements_py(self, displayed, ids, keep, names, nodes):
        D = ElemData()
        if ids is not None:
            objs = self.items(self.ent.Element, ids)
        else:
            objs = self.elements(displayed)
        nf = self._names_for(names)
        allnames = set()
        rows = []
        N = len(objs)
        for i, e in enumerate(objs):
            try:
                eid = e.id
            except Exception:
                continue
            D.all_ids.append(eid)
            cfg = attr(e, "config")
            try:
                cfg = int(cfg)
            except (TypeError, ValueError):
                cfg = 0
            if keep is not None and not keep(cfg):
                continue
            row = {}
            for nm in nf(cfg):
                allnames.add(nm)
                if nm == "collector.id":
                    row[nm] = attr(attr(e, "collector"), "id")
                else:
                    v = num_attr(e, nm)
                    if v is None:
                        for cand in DATANAME_CANDIDATES.get(nm, ()):
                            if cand != nm:
                                v = num_attr(e, cand)
                                if v is not None:
                                    break
                    row[nm] = v
            D.ids.append(eid)
            D.cfg.append(cfg)
            rows.append(row)
            if nodes:
                if D.nodes is None:
                    D.nodes = []
                D.nodes.append(tuple(self.elem_node_ids(e)))
            if i % PROGRESS_EVERY == 0:
                BUS.progress(T("Odczyt element\u00f3w: %d / %d", "Reading elements: %d / %d", i, N))
        for nm in allnames:
            D.vals[nm] = [r.get(nm) for r in rows]
        if nodes and D.nodes is None:
            D.nodes = []
        return D

    def read_nodes(self, nids):
        """Wspolrzedne wezlow: slownik id -> (x, y, z)."""
        nids = sorted(set(int(n) for n in nids))
        if not nids:
            return {}
        if self.bulk_ok():
            try:
                out = {}
                step = max(1, int(self.MARK_CHUNK) * 5)
                for a in range(0, len(nids), step):
                    part = nids[a:a + step]
                    got = self._tcl_mark("nodes", "ids", part)
                    xs, ys, zs = [self._tcl_vals("nodes", c, got) for c in ("x", "y", "z")]
                    for n, x, y, z in zip(got, xs, ys, zs):
                        if None not in (x, y, z):
                            out[n] = (x, y, z)
                    if len(nids) > step:
                        BUS.progress(T("Wsp\u00f3\u0142rz\u0119dne w\u0119z\u0142\u00f3w: %d / %d", "Node coordinates: %d / %d", min(a + step, len(nids)), len(nids)))
                return out
            except Exception as e:
                BUS.log("Tcl node read -> Python API: %s" % e)
            finally:
                self._clear_marks()
        out = {}
        for i, n in enumerate(self.items(self.ent.Node, nids)):
            p = self.node_xyz(n)
            if p is not None:
                out[n.id] = p
            if i % (PROGRESS_EVERY * 5) == 0:
                BUS.progress(T("Wsp\u00f3\u0142rz\u0119dne w\u0119z\u0142\u00f3w: %d / %d", "Node coordinates: %d / %d", i, len(nids)))
        return out

    def comp_names(self):
        """Slownik id komponentu -> nazwa."""
        out = {}
        for c in self.components():
            try:
                out[c.id] = c.name
            except Exception:
                pass
        return out

    # ---------------------------------------------------------- komponenty
    def components(self):
        return self.items(self.ent.Component)

    def comp_index(self):
        """Slownik nazwa -> id wszystkich komponentow."""
        out = {}
        for c in self.components():
            try:
                out[c.name] = c.id
            except Exception:
                pass
        return out

    def comp_id(self, name, index=None):
        idx = index if index is not None else self.comp_index()
        return idx.get(name, 0)

    def ensure_comp(self, name, color_idx=None, index=None):
        """ID komponentu o nazwie (tworzy, gdy brak); opcjonalnie kolor."""
        cid = self.comp_id(name, index)
        if not cid:
            c = self.ent.Component(self.model())
            c.name = name
            cid = c.id
            if index is not None:
                index[name] = cid
        if color_idx:
            self.set_comp_color(cid, color_idx)
        return cid

    def set_comp_color(self, cid, idx):
        try:
            c = self.ent.Component(self.model(), int(cid))
            c.color = int(idx)
            return True
        except Exception:
            return False

    def comp_elem_ids(self, cid):
        m = self.model()
        src = self.hm.Collection(m, self.ent.Component, [int(cid)])
        col = self.hm.Collection(m, self.hm.FilterByCollection(self.ent.Element, self.ent.Component), src)
        return [e.id for e in list(col)]

    def comp_elem_count(self, cid):
        """Liczba elementow komponentu(-ow). len() kolekcji liczy w C++ -
        bez tworzenia obiektu Pythona dla kazdego elementu (c.elements)."""
        cids = [int(c) for c in (cid if isinstance(cid, (list, tuple, set)) else [cid])]
        if not cids:
            return 0
        try:
            m = self.model()
            src = self.hm.Collection(m, self.ent.Component, cids)
            col = self.hm.Collection(m, self.hm.FilterByCollection(self.ent.Element, self.ent.Component), src)
            return len(col)
        except Exception:
            try:
                return sum(len(self.ent.Component(self.model(), c).elements or []) for c in cids)
            except Exception:
                return -1

    def delete_comps(self, ids):
        self._clear_marks()
        col = self.collection(self.ent.Component, ids)
        if col is None:
            return True
        return self.status_of(self.model().deletemark(col))[0] == 0

    def show_comps(self, ids, on=True):
        col = self.collection(self.ent.Component, ids)
        if col is None:
            return
        try:
            self.model().displaycollectorsbymark(col, "on" if on else "off", 1, 0)
        except Exception:
            pass

    def show_all_comps(self, on=True):
        try:
            col = self.hm.Collection(self.model(), self.ent.Component)
            self.model().displaycollectorsbymark(col, "on" if on else "off", 1, 0)
        except Exception:
            pass

    def move_elements(self, ids, comp_name):
        """Przenosi elementy do komponentu (po nazwie); True gdy bez bledu."""
        self._clear_marks()
        col = self.collection(self.ent.Element, ids)
        if col is None:
            return True
        try:
            return self.status_of(self.model().movemark(col, comp_name))[0] == 0
        except Exception:
            return False

    def set_transparency(self, comp_ids, level):
        col = self.collection(self.ent.Component, comp_ids)
        if col is None:
            return False
        try:
            m = self.model()
            m.transparencyvalue(int(level))
            return self.status_of(m.transparencymark(col))[0] == 0
        except Exception:
            return False

    # ---------------------------------------------------------- zestawy, tagi
    def create_set(self, name, ids):
        self._clear_marks()
        dm = self.debug()
        col = self.collection(self.ent.Element, ids)
        if dm is None or col is None:
            return 0
        try:
            if self.status_of(dm.entitysetcreate(name, col))[0] != 0:
                return 0
        except Exception:
            return 0
        for s in self.items(self.ent.Set):
            if attr(s, "name") == name:
                return s.id
        return 0

    def delete_sets(self, names):
        names = set(names)
        ids = [s.id for s in self.items(self.ent.Set) if attr(s, "name") in names]
        col = self.collection(self.ent.Set, ids)
        if col is not None:
            try:
                self.model().deletemark(col)
            except Exception:
                pass

    def create_tag(self, eid, label, body, color=3):
        dm = self.debug()
        if dm is None:
            return False
        try:
            return self.status_of(dm.tagcreate(self.element(eid), label, body, int(color)))[0] == 0
        except Exception:
            return False

    def delete_tags(self, labels):
        labels = set(labels)
        ids = [t.id for t in self.items(self.ent.Tag) if attr(t, "label") in labels]
        col = self.collection(self.ent.Tag, ids)
        if col is not None:
            try:
                self.model().deletemark(col)
            except Exception:
                pass

    # ---------------------------------------------------------- pliki
    def model_file(self):
        """Sciezka otwartego pliku .hm ("" gdy nieznana / nowy model)."""
        s = self.tcl("hm_info currentfile", "") or ""
        s = s.strip()
        if s and os.path.isfile(s):
            return os.path.normpath(s)
        return ""

    def model_dir(self):
        f = self.model_file()
        return os.path.dirname(f) if f else ""

    def model_name(self):
        f = self.model_file()
        return os.path.splitext(os.path.basename(f))[0] if f else T("model", "model")

    def signature(self):
        """Podpis modelu w sesji: plik + liczba elementow (wykrycie podmiany)."""
        return "%s|%d" % (self.model_file(), self.count(self.ent.Element))

    def read_file(self, path):
        """Wczytuje plik .hm (zastepuje model w sesji). Zwraca (ok, komunikat).
        Kod bledu z readfile nie zawsze znaczy porazke (ostrzezenia), a brak
        bledu nie gwarantuje, ze w sesji jest nowy plik - w GUI porownujemy
        sciezke otwartego modelu z zadana (inaczej np. delta liczylaby sie
        na STARYM modelu, ktory zostal w sesji po nieudanym wczytaniu)."""
        p = os.path.normpath(path).replace("\\", "/")
        try:
            try:
                self.model().hm_answernext("yes")
            except Exception:
                pass
            code, msg = self.status_of(self.model().readfile(p, 0))
        except Exception as e:
            code, msg = -1, "%s" % e
        self.refresh()
        cur = self.model_file()
        if cur:
            try:
                same = os.path.samefile(cur, path)          # takze nazwy 8.3 / dyski sieciowe
            except OSError:
                same = os.path.normcase(os.path.normpath(cur)) == os.path.normcase(os.path.normpath(path))
            return (True, msg) if same else (False, msg or T("w sesji jest inny plik: %s", "another file is in the session: %s", cur))
        if code != 0 and self.count(self.ent.Element) <= 0:
            return False, msg
        return True, msg

    # ---------------------------------------------------------- dolaczanie pliku
    # Dolaczenie .hm (nakladka REF) scala CALY model: elementy, wezly,
    # komponenty, materialy, obciazenia, uklady wspolrzednych... Zeby
    # "Przywroc siatke" zostawilo model dokladnie takim, jaki byl, przed
    # scaleniem liczymy encje KAZDEGO typu (len kolekcji - szybko), a po
    # scaleniu zapamietujemy ID nowych encji typow, ktorych przybylo.
    # Usuwanie: elementy, potem pozostale typy, na koncu wezly.
    def entity_classes(self):
        out = []
        for nm in dir(self.ent):
            c = getattr(self.ent, nm, None)
            if isinstance(c, type) and not nm.startswith("_"):
                out.append((nm, c))
        return out

    def entity_counts(self, classes):
        out = {}
        for nm, c in classes:
            try:
                col = self.hm.Collection(self.model(), c)
                out[nm] = len(col)
            except Exception:
                pass
        return out

    def entity_ids(self, cls_name):
        """ID encji typu (lista int) albo None, gdy typ nie ma ID liczbowych."""
        if cls_name in ("Element", "Node") and self.bulk_ok():
            try:
                return self._tcl_mark("elems" if cls_name == "Element" else "nodes", "all")
            except Exception:
                pass
            finally:
                self._clear_marks()
        try:
            col = self.hm.Collection(self.model(), getattr(self.ent, cls_name))
            return [int(x.id) for x in col]
        except Exception:
            return None

    def merge_file(self, path):
        """Dolacza plik .hm do sesji. Zwraca {typ: [ID nowych encji]} albo {}
        (gdy sie nie udalo albo przybyly encje, ktorych nie umiemy sledzic -
        wtedy scalenie jest od razu cofane)."""
        p = os.path.normpath(path).replace("\\", "/")
        classes = self.entity_classes()
        c0 = self.entity_counts(classes)
        ids0 = {}
        for nm, n in c0.items():
            if n:
                got = self.entity_ids(nm)
                if got is not None:
                    ids0[nm] = set(got)
        m = self.model()
        merged = False
        for fn in (lambda: m.mergefile2(p), lambda: m.mergefile(p, 1, 1)):
            try:
                fn()
            except Exception:
                continue
            self.refresh()
            if self.count(self.ent.Element) > c0.get("Element", 0):
                merged = True
                break
        if not merged:
            return {}
        c1 = self.entity_counts(classes)
        new, lost = {}, []
        for nm, n in c1.items():
            if n == c0.get(nm, 0):
                continue
            if n < c0.get(nm, 0):
                lost.append(nm)
                continue
            before = ids0.get(nm)
            if before is None and c0.get(nm, 0) == 0:
                before = set()
            now = self.entity_ids(nm)
            if before is None or now is None:
                lost.append(nm)
                continue
            new[nm] = sorted(set(now) - before)
        if lost:
            # nie umiemy tego bezpiecznie cofnac w przyszlosci - cofamy od razu
            BUS.log("merge: untracked entity types %s - overlay undone" % ", ".join(lost))
            self.delete_entities(new)
            return {}
        return new

    def delete_entities(self, ents):
        """Usuwa encje {typ: [ID]} w kolejnosci zaleznosci. Zwraca liczbe
        ID, ktore mimo to zostaly w modelu."""
        order = ["Element"] + sorted(k for k in ents if k not in ("Element", "Node")) + ["Node"]
        m = self.model()
        self._clear_marks()
        for nm in order:
            ids = ents.get(nm) or []
            col = self.collection(getattr(self.ent, nm), ids) if ids else None
            if col is None or len(col) == 0:
                continue
            try:
                m.deletemark(col)
            except Exception:
                pass
        left = 0
        for nm, ids in ents.items():
            col = self.collection(getattr(self.ent, nm), ids) if ids else None
            if col is not None:
                left += len(col)
        return left

    # ---------------------------------------------------------- widok, zrzut
    def get_view(self):
        """Widok jako lista liczb: 20 (macierz 4x4 + zakres okna, jak *viewset)
        w sesji GUI; w trybie wsadowym 16 (sama macierz)."""
        s = self.tcl("hm_winfo viewmatrix")
        if s:
            nums = [to_float(x) for x in s.split()]
            nums = [x for x in nums if x is not None]
            if len(nums) >= 16:
                return nums
        try:
            r = self.model().hm_getcurrentview()
            vm = r[1].viewMatrix
            return [float(v) for row in vm for v in row]
        except Exception:
            return []

    def set_view(self, nums):
        nums = [float(x) for x in nums]
        if len(nums) == 16:
            cur = self.get_view()
            nums = nums + (cur[16:20] if len(cur) >= 20 else [-1.0, -1.0, 1.0, 1.0])
        if len(nums) < 20:
            return False
        try:
            return self.status_of(self.model().viewset(*nums[:20]))[0] == 0
        except Exception:
            return False

    def capture(self, path):
        """Zrzut okna graficznego do pliku (png/jpg/bmp/tif). True gdy plik jest."""
        ext = os.path.splitext(path)[1].lower().lstrip(".") or "png"
        ext = {"jpeg": "jpg", "tiff": "tif"}.get(ext, ext)
        try:
            os.remove(path)
        except OSError:
            pass
        try:
            import hw as _hw
            tool = _hw.CaptureImageTool()
            tool.type = ext
            tool.file = path
            tool.capture()
        except Exception:
            pass
        if not (os.path.isfile(path) and os.path.getsize(path) > 0) and self.has_tcl():
            # zapasowo: natywny zrzut JPG panelu i konwersja formatu. Tylko z
            # oknem graficznym: w hmbatch jpegfilenamed konczy proces bledem
            # libjpeg ("Empty JPEG image").
            jp = path if ext == "jpg" else os.path.splitext(path)[0] + "_hm.jpg"
            try:
                self.model().jpegfilenamed(jp)
            except Exception:
                pass
            if jp != path and os.path.isfile(jp):
                convert_image(jp, path)
                try:
                    os.remove(jp)
                except OSError:
                    pass
        return os.path.isfile(path) and os.path.getsize(path) > 0

    def set_scene_background(self, hexcol):
        """Kolor tla SCENY (ekranu). Zrzuty HM maja zawsze tlo biale - kolor
        tla na zrzucie jest nakladany osobno (przebarwienie pikseli)."""
        if self.tcl("hwf::setbackgroundcolor %s" % hexcol) is not None:
            return True
        dm = self.debug()
        if dm is not None:
            try:
                r, g, b = hex_rgb(hexcol)
                return self.status_of(dm.setbackgroundcolor(r, g, b))[0] == 0
            except Exception:
                pass
        return False

    def refresh_palette(self):
        """Paleta czytana od nowa przy nastepnej analizie (uzytkownik mogl
        zmienic kolory palety w HM w trakcie sesji)."""
        self._palette = None

    def leftover_comps(self, prefixes):
        """Komponenty o nazwach narzedzia (prefiks_) z elementami, ktorych
        biezacy stan nie zna - np. zapisany pokolorowany model."""
        own = set(MESH.own)
        out = []
        for nm, cid in self.comp_index().items():
            if nm in own or nm == MESH.FALLBACK:
                continue
            if any(nm.startswith(p + "_") for p in prefixes if p) and self.comp_elem_count(cid) > 0:
                out.append(nm)
        return sorted(out)

    def palette(self):
        """Paleta 64 kolorow {indeks: (r, g, b)} - z sesji albo domyslna."""
        if self._palette is None:
            pal = {}
            s = self.tcl("hm_winfo entitycolors")
            if s:
                nums = [int(x) for x in re.findall(r"-?\d+", s)]
                for i in range(len(nums) // 3):
                    pal[i + 1] = tuple(nums[3 * i:3 * i + 3])
            if len(pal) < 64:
                pal = dict((i + 1, tuple(c)) for i, c in enumerate(HM_PALETTE))
            self._palette = pal
        return self._palette

    def diagnostics(self, sample=6):
        """Samokontrola: API, most Tcl, nazwy danych metryk (z wartosciami na
        probce elementow), zgodnosc odczytu hurtowego z pojedynczym, zrzut
        okna, biblioteki. Lista krotek (nazwa, ok, szczegoly)."""
        out = []
        ok = self.ok()
        out.append((T("API HyperMesha (modu\u0142 hm)", "HyperMesh API (hm module)"), ok, self.error or "OK"))
        if not ok:
            return out
        out.append((T("Most Tcl (hw.evalTcl)", "Tcl bridge (hw.evalTcl)"), self.has_tcl(),
                    T("dost\u0119pny", "available") if self.has_tcl() else T("brak \u2013 odczyt przez API Pythona (wolniej)", "missing \u2013 Python API reads (slower)")))
        out.append((T("Odczyt hurtowy (hm_getvalue mark=1)", "Bulk read (hm_getvalue mark=1)"), self.bulk_ok(),
                    T("aktywny", "active") if self.bulk_ok() else T("nieaktywny", "inactive")))
        out.append((T("Plik modelu", "Model file"), bool(self.model_file()), self.model_file() or T("(nieznany / niezapisany)", "(unknown / unsaved)")))
        n = self.count(self.ent.Element)
        out.append((T("Elementy w modelu", "Elements in the model"), n > 0, "%d" % n))
        if n <= 0:
            return out
        D = self.read_elements(names=())
        reps = {}
        for eid, cfg in zip(D.ids, D.cfg):
            reps.setdefault(cfg, eid)
        rep_ids = [reps[c] for c in sorted(reps)]
        out.append((T("Typy element\u00f3w", "Element types"), True,
                    ", ".join("%s (id %d)" % (elem_type_name(c), reps[c]) for c in sorted(reps))))
        for pref in ("aspect", "jacobian", "skew", "warpage", "taper", "minangle", "maxangle", "shortestside", "volume", "area"):
            self._names.pop(pref, None)
            dn = self.resolve_name(pref, rep_ids)
            vals = []
            for eid in rep_ids[:sample]:
                v = self._probe_value(eid, dn)
                vals.append("%s" % ("-" if v is None else fmt_num(v, 4)))
            good = any(v != "-" for v in vals)
            out.append(("dataname %s" % pref, good, "%s%s: %s" % (dn, "" if dn == pref else " (%s)" % pref, ", ".join(vals))))
        if self.bulk_ok() and D.ids:
            import random
            ids = random.sample(D.ids, min(sample * 4, len(D.ids)))
            dn = self.resolve_name("aspect", ids)
            bulk = self.read_elements(ids=ids, names=(dn,))
            mism, checked = 0, 0
            for eid, v in zip(bulk.ids, bulk.vals.get(dn, [])):
                s = self._probe_value(eid, dn)
                checked += 1
                if (v is None) != (s is None) or (v is not None and abs(v - s) > 1e-9 * max(1.0, abs(s))):
                    mism += 1
            out.append((T("Zgodno\u015b\u0107 odczytu hurtowego z pojedynczym (%s, %d el.)", "Bulk vs single read consistency (%s, %d el.)", dn, checked),
                        mism == 0, T("niezgodnych: %d", "mismatches: %d", mism)))
            ids2 = self._tcl_mark("elems", "ids", ids)
            order = self._tcl_order("elems", ids2)
            self._clear_marks()
            out.append((T("Kolejno\u015b\u0107 odpowiedzi mark=1 vs znacznik", "mark=1 reply order vs mark"), True,
                        T("identyczna", "identical") if order == ids2 else T("R\u00d3\u017bNA \u2013 warto\u015bci parowane po ID (obs\u0142u\u017cone)", "DIFFERENT \u2013 values paired by ID (handled)")))
        g = geometry_selftest()
        out.append((T("Wzorce geometrii (Jacobian Zero, tet collapse)", "Geometry references (Jacobian Zero, tet collapse)"), not g, g or "OK"))
        p = tmp_file("png", "diag")
        cap = self.capture(p)
        out.append((T("Zrzut okna graficznego", "Graphics window capture"), cap, ("%d x %d px" % image_size(p)) if cap else T("nieudany", "failed")))
        try:
            os.remove(p)
        except OSError:
            pass
        for nm, mod in (("python-pptx", pptx), ("XlsxWriter", xlsxwriter), ("Pillow", Image), ("numpy", np)):
            out.append((nm, mod is not None, T("jest", "present") if mod is not None else T("brak", "missing")))
        out.append(("PyQt5", qt() is not None, T("jest", "present") if qt() is not None else T("brak", "missing")))
        return out

    def eval_commands(self, text):
        """Odtwarza komendy Tcl ukladu (z command.tcl). Zwraca (ok, bledow)."""
        ok = bad = 0
        for ln in (text or "").splitlines():
            t = ln.strip()
            if not t or t.startswith("#"):
                continue
            if self.tcl(t) is None:
                bad += 1
            else:
                ok += 1
        return ok, bad


HM = HmApi()


# ============ GEOMETRIA ELEMENTU: JACOBIAN ZERO, TET COLLAPSE ==========
# JACOBIAN ZERO = najmniejszy wyznacznik macierzy Jacobiego det(J) w narozach
# elementu. HyperMesh 2024 nie udostepnia tej wartosci w API, wiec liczymy
# ja ze wspolrzednych wezlow narozy (tria, quad, tetra, hexa, penta,
# piramida; elementy II rzedu - z wezlow narozy). Wartosc <= 0 oznacza
# element zdegenerowany albo wywiniety. Znak dla bryl jest potem
# normalizowany do orientacji WIEKSZOSCI elementow danego typu (numeracja
# wezlow bywa w modelach odwrotna) - patrz WIELE METRYK.
# TET COLLAPSE (raport jakosci): dla kazdego wezla tetry odleglosc od
# przeciwleglej sciany / sqrt(pole sciany), unormowana tak, by tetra
# foremna miala 1.0; minimum po czterech wezlach (definicja HyperMesha).
def v_sub(a, b):
    return (a[0] - b[0], a[1] - b[1], a[2] - b[2])


def v_cross(a, b):
    return (a[1] * b[2] - a[2] * b[1], a[2] * b[0] - a[0] * b[2], a[0] * b[1] - a[1] * b[0])


def v_dot(a, b):
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def v_len(a):
    return math.sqrt(v_dot(a, a))


def triple(a, b, c):
    return v_dot(a, v_cross(b, c))


def det_tri(P):
    """Tria: pole rownolegloboku (dodatnie; zero = element zdegenerowany)."""
    return v_len(v_cross(v_sub(P[1], P[0]), v_sub(P[2], P[0])))


def det_quad(P):
    """Quad (xi, eta w [-1, 1]): det w 4 narozach wzgledem normalnej w srodku."""
    p0, p1, p2, p3 = P[:4]
    gx = tuple((p1[k] + p2[k] - p0[k] - p3[k]) / 4.0 for k in range(3))
    gy = tuple((p2[k] + p3[k] - p0[k] - p1[k]) / 4.0 for k in range(3))
    n = v_cross(gx, gy)
    nl = v_len(n)
    if nl <= 0.0:
        return 0.0
    n = (n[0] / nl, n[1] / nl, n[2] / nl)
    a01, a32 = v_sub(p1, p0), v_sub(p2, p3)
    b03, b12 = v_sub(p3, p0), v_sub(p2, p1)
    return min(v_dot(v_cross(a, b), n) / 4.0
               for a, b in ((a01, b03), (a01, b12), (a32, b12), (a32, b03)))


def det_tet(P):
    return triple(v_sub(P[1], P[0]), v_sub(P[2], P[0]), v_sub(P[3], P[0]))


_HEX_CORNERS = (
    ((0, 1), (0, 3), (0, 4)), ((0, 1), (1, 2), (1, 5)), ((3, 2), (1, 2), (2, 6)), ((3, 2), (0, 3), (3, 7)),
    ((4, 5), (4, 7), (0, 4)), ((4, 5), (5, 6), (1, 5)), ((7, 6), (5, 6), (2, 6)), ((7, 6), (4, 7), (3, 7)),
)


def det_hex(P):
    """Hexa (xi, eta, zeta w [-1, 1]): w narozu pochodne = krawedzie / 2."""
    best = None
    for row in _HEX_CORNERS:
        d = [v_sub(P[b], P[a]) for a, b in row]
        v = triple(d[0], d[1], d[2]) / 8.0
        if best is None or v < best:
            best = v
    return best


def det_penta(P):
    best = None
    for i in range(6):
        if i < 3:
            dr, ds, dz = v_sub(P[1], P[0]), v_sub(P[2], P[0]), v_sub(P[i + 3], P[i])
        else:
            dr, ds, dz = v_sub(P[4], P[3]), v_sub(P[5], P[3]), v_sub(P[i], P[i - 3])
        v = triple(dr, ds, dz) / 2.0
        if best is None or v < best:
            best = v
    return best


def det_pyr(P):
    apex = P[4]
    best = None
    for i in range(4):
        pi = P[i]
        v = triple(v_sub(P[(i + 1) % 4], pi), v_sub(P[(i + 3) % 4], pi), v_sub(apex, pi))
        if best is None or v < best:
            best = v
    return best


DET_FUNCS = {"tria": det_tri, "quad": det_quad, "tet": det_tet,
             "hex": det_hex, "penta": det_penta, "pyr": det_pyr}


def jacobian_zero(points, shape):
    """Najmniejszy det(J) w narozach elementu (None gdy typ nieobslugiwany)."""
    f = DET_FUNCS.get(shape)
    if f is None or points is None or len(points) < CORNERS[shape]:
        return None
    try:
        return f(points)
    except Exception:
        return None


def tet_collapse(points):
    """Tet collapse (1.0 = tetra foremna, 0 = splaszczona)."""
    if points is None or len(points) < 4:
        return None
    best = None
    for i in range(4):
        others = [points[j] for j in range(4) if j != i]
        n = v_cross(v_sub(others[1], others[0]), v_sub(others[2], others[0]))
        area2 = v_len(n)
        if area2 <= 0:
            return 0.0
        h = abs(v_dot(v_sub(points[i], others[0]), n)) / area2
        v = h / math.sqrt(area2 / 2.0) / 1.2408064788027997
        if best is None or v < best:
            best = v
    return best


def geometry_selftest():
    """Sprawdza funkcje geometrii na wzorcach o znanym wyniku. Zwraca ""
    (wszystko OK) albo tekst z pierwsza rozbieznoscia."""
    def close(a, b, eps=1e-9):
        return a is not None and abs(a - b) <= eps
    checks = []
    tri = [(0, 0, 0), (1, 0, 0), (0, 1, 0)]
    checks.append(("tria", close(det_tri(tri), 1.0)))
    quad = [(0, 0, 0), (1, 0, 0), (1, 1, 0), (0, 1, 0)]
    checks.append(("quad", close(det_quad(quad), 0.25)))
    checks.append(("quad odwrocony = ten sam quad (normalna wlasna)", close(det_quad([quad[0], quad[3], quad[2], quad[1]]), 0.25)))
    checks.append(("quad niewypukly < 0", det_quad([(0, 0, 0), (1, 0, 0), (0.2, 0.2, 0), (0, 1, 0)]) < 0))
    checks.append(("quad zdegenerowany", close(det_quad([(0, 0, 0), (1, 0, 0), (1, 0, 0), (0, 0, 0)]), 0.0)))
    tet = [(0, 0, 0), (1, 0, 0), (0, 1, 0), (0, 0, 1)]
    checks.append(("tet", close(det_tet(tet), 1.0)))
    hexa = [(0, 0, 0), (1, 0, 0), (1, 1, 0), (0, 1, 0), (0, 0, 1), (1, 0, 1), (1, 1, 1), (0, 1, 1)]
    checks.append(("hexa", close(det_hex(hexa), 0.125)))
    penta = [(0, 0, 0), (1, 0, 0), (0, 1, 0), (0, 0, 1), (1, 0, 1), (0, 1, 1)]
    checks.append(("penta", close(det_penta(penta), 0.5)))
    pyr = [(0, 0, 0), (1, 0, 0), (1, 1, 0), (0, 1, 0), (0.5, 0.5, 1)]
    checks.append(("piramida", close(det_pyr(pyr), 1.0)))
    a = 1.0
    reg = [(0, 0, 0), (a, 0, 0), (a / 2.0, a * math.sqrt(3) / 2.0, 0), (a / 2.0, a * math.sqrt(3) / 6.0, a * math.sqrt(2.0 / 3.0))]
    checks.append(("tet collapse foremny = 1", close(tet_collapse(reg), 1.0, 1e-6)))
    checks.append(("tet collapse plaski = 0", close(tet_collapse([(0, 0, 0), (1, 0, 0), (0, 1, 0), (1, 1, 0)]), 0.0, 1e-9)))
    checks.append(("jacobian_zero shape n/d", jacobian_zero(tet, "shell") is None))
    bad = [n for n, ok in checks if not ok]
    return ", ".join(bad)


# ================ SKALE, PASMA I PALETA HYPERMESHA ====================
# Skala metryki = lista granic pasm (rosnaco) + kolor kazdego pasma.
# Kierunek metryki: "above" = wyzsze gorsze (AR, Skewness), "below" =
# nizsze gorsze (Jacobian Ratio, Jacobian Zero).
# PODZIAL WG PROGU: wartosc po "dobrej" stronie progu -> grupa "w normie";
# pozostale -> pasma od progu do wartosci skrajnej; wartosci poza ostatnia
# granica -> osobna grupa "poza skala" (albo najgorsze pasmo).
# Kolory: komponent HM 2024 ma kolor = indeks palety 64 kolorow, wiec
# kazde pasmo dostaje kolor Z PALETY. Domyslnie kolejne pasma biora
# kolejne kolory z "rampy" (RAMP_BAD / RAMP_FULL) - sa zawsze rozne.
class PaletteMap(object):
    """Dopasowanie kolorow RGB do palety HyperMesha (indeksy 1..64)."""

    NEUTRAL = (1, 2, 9, 10, 11, 12, 13, 14, 15, 16)   # czern, biel, szarosci

    def __init__(self, pal=None):
        self.pal = pal or dict((i + 1, tuple(c)) for i, c in enumerate(HM_PALETTE))
        self.rev = {}
        for i, c in sorted(self.pal.items()):
            self.rev.setdefault(tuple(c), i)

    def rgb(self, idx):
        return tuple(self.pal.get(int(idx), (128, 128, 128)))

    @staticmethod
    def dist(a, b):
        return 3.0 * (a[0] - b[0]) ** 2 + 4.0 * (a[1] - b[1]) ** 2 + 2.0 * (a[2] - b[2]) ** 2

    def index_of(self, rgb):
        """Indeks palety dla koloru: dokladny, a gdy go brak - najblizszy
        (kolory szare/czarne/biale tylko dla kolorow szarych)."""
        rgb = tuple(int(c) for c in rgb[:3])
        if rgb in self.rev:
            return self.rev[rgb]
        grayish = max(rgb) - min(rgb) < 30
        best, bd = 3, None
        for i, c in self.pal.items():
            if not grayish and i in self.NEUTRAL:
                continue
            d = self.dist(rgb, c)
            if bd is None or d < bd:
                best, bd = i, d
        return best

    def snap(self, rgb):
        """Kolor palety najblizszy podanemu (to, co faktycznie bedzie na siatce)."""
        return self.rgb(self.index_of(rgb))


def palette_map():
    return PaletteMap(HM.palette() if HM.ok() else None)


def ramp_indices(k, ramp):
    """k indeksow palety rozlozonych rowno wzdluz rampy (rozne, gdy k <= len)."""
    n = len(ramp)
    if k <= 0:
        return []
    if k == 1:
        return [ramp[-1]]
    if k <= n:
        return [ramp[int(round(i * (n - 1) / float(k - 1)))] for i in range(k)]
    # wiecej pasm niz kolorow rampy: interpolacja i najblizszy kolor palety
    pm = PaletteMap()
    stops = [pm.rgb(i) for i in ramp]
    return [pm.index_of(grad_at(stops, i / float(k - 1))) for i in range(k)]


def default_band_colors(direction, k, thr_on):
    """Kolory k pasm w kolejnosci ROSNACYCH wartosci. Z progiem: od pasma
    najblizszego normie (niebieski) do najgorszego (czerwony); bez progu:
    pelna rampa. Dla metryk "nizsze gorsze" kolejnosc jest odwrocona."""
    ramp = RAMP_BAD if thr_on else RAMP_FULL
    idx = ramp_indices(k, ramp)
    pm = PaletteMap()
    out = []
    for i in range(k):
        rank = i if direction == "above" else k - 1 - i
        out.append(pm.rgb(idx[rank]))
    return out


def manual_edges(edges, thr_on, thr, direction):
    """Granice trybu recznego: liczby, rosnaco, granica po stronie normy = prog."""
    E = sorted(set(f for f in (to_float(e) for e in edges) if f is not None))
    if thr_on and thr is not None:
        eps = 1e-12 * max(1.0, abs(thr))
        if direction == "above":
            E = [thr] + [e for e in E if e > thr + eps]
        else:
            E = [e for e in E if e < thr - eps] + [thr]
    return E


def auto_edges(direction, thr_on, thr, nb, vmin, vmax, nice=True):
    """Granice trybu automatycznego: od progu do wartosci skrajnej (albo
    min..max, gdy prog wylaczony), podzielone na nb pasm."""
    n = max(1, min(40, int(nb or 5)))
    vmin = to_float(vmin, 0.0)
    vmax = to_float(vmax, 1.0)
    if thr_on and thr is not None:
        if direction == "above":
            lo, hi = thr, max(vmax, thr)
        else:
            lo, hi = min(vmin, thr), thr
        span = hi - lo
        if span <= 1e-12 * max(1.0, abs(thr)):
            span = abs(thr) * 0.5 if abs(thr) > 0 else 1.0
        step = span / float(n)
        if nice:
            step = nice_step(step)
        if direction == "above":
            E = [lo + i * step for i in range(n + 1)]
        else:
            E = [hi - i * step for i in range(n, -1, -1)]
    else:
        lo, hi = vmin, vmax
        if hi <= lo:
            hi = lo + (abs(lo) * 0.1 if abs(lo) > 0 else 1.0)
        step = (hi - lo) / float(n)
        if nice:
            step = nice_step(step)
            lo2 = math.floor(lo / step) * step
            guard = 0
            while lo2 + n * step < hi - 1e-12 * max(1.0, abs(hi)) and guard < 30:
                step = nice_step(step * 1.01)
                lo2 = math.floor(lo / step) * step
                guard += 1
            lo = lo2
        E = [lo + i * step for i in range(n + 1)]
    return [clean10(e) for e in E]


class Scale(object):
    """Efektywna skala metryki: granice, kolory, prog, kierunek, tryb."""

    OK, OVER, UNDER, NA = -1, -2, -3, -9     # klasy specjalne

    def __init__(self, edges, colors, thr_on, thr, direction, mode, strict=False):
        self.strict = bool(strict)       # wartosc == prog jest juz poza norma
        self.edges = list(edges)
        self.k = len(self.edges) - 1
        self.colors = list(colors)
        self.thr_on = bool(thr_on)
        self.thr = thr if thr is not None else 0.0
        self.dir = direction
        self.mode = mode
        self.dec = dec_of(self.edges)

    @classmethod
    def build(cls, cfg, direction, vmin, vmax, nice=True, strict=False):
        """Skala z ustawien metryki (slownik thr_on, thr, mode, nb, edges,
        colors) i zakresu danych [vmin, vmax]."""
        thr = to_float(cfg.get("thr"))
        t_on = bool(cfg.get("thr_on")) and thr is not None
        if vmin is None or vmax is None:
            vmin = vmax = thr if thr is not None else 0.0
        mode = cfg.get("mode", "auto")
        E = []
        if mode == "manual":
            E = manual_edges(cfg.get("edges") or [], t_on, thr, direction)
        if len(E) < 2:
            mode = "auto"
            E = auto_edges(direction, t_on, thr, cfg.get("nb", 5), vmin, vmax, nice)
        k = len(E) - 1
        cols = [tuple(c) for c in (cfg.get("colors") or [])]
        if mode != "manual" or len(cols) != k:
            cols = default_band_colors(direction, k, t_on)
        return cls(E, cols, t_on, thr, direction, mode, strict)

    def ok_op(self):
        """Znak warunku "w normie" (do legend i tabel)."""
        if self.dir == "above":
            return "<" if self.strict else "\u2264"
        return ">" if self.strict else "\u2265"

    def bad_op(self):
        """Znak warunku "poza norma"."""
        if self.dir == "above":
            return "\u2265" if self.strict else ">"
        return "\u2264" if self.strict else "<"

    def classify(self, v, over_sep=True):
        """-1 w normie, -2 poza skala (za najgorszym koncem), -3 ponizej skali
        (tylko bez progu), -9 brak wartosci, k >= 0 = pasmo (rosnaco)."""
        if v is None:
            return self.NA
        E, K = self.edges, self.k
        if self.thr_on:
            if self.dir == "above":
                if v < self.thr or (v == self.thr and not self.strict):
                    return self.OK
                if v > E[-1]:
                    return self.OVER if over_sep else K - 1
                i = bisect.bisect_left(E, v) - 1       # przedzialy (a, b]
            else:
                if v > self.thr or (v == self.thr and not self.strict):
                    return self.OK
                if v < E[0]:
                    return self.OVER if over_sep else 0
                i = bisect.bisect_right(E, v) - 1      # przedzialy [a, b)
        else:
            if v < E[0]:
                return self.UNDER if over_sep else 0
            if v > E[-1]:
                return self.OVER if over_sep else K - 1
            i = bisect.bisect_right(E, v) - 1
        return max(0, min(K - 1, i))

    def band_rank(self, i):
        """Pozycja pasma liczona od normy: 1 = najblizej normy, K = najgorsze."""
        return i + 1 if self.dir == "above" else self.k - i

    def worst_first(self):
        """Indeksy pasm od najgorszego do najblizszego normie."""
        return list(range(self.k - 1, -1, -1)) if self.dir == "above" else list(range(self.k))


# ======================= STAN SIATKI I PRZYWRACANIE ====================
# Moduly "Wiele metryk" i "Delta" koloruja siatke, PRZENOSZAC elementy do
# wlasnych komponentow (kolor komponentu = kolor grupy). Wspolny obiekt MESH
# pamieta, w jakim komponencie kazdy element byl PIERWOTNIE - niezaleznie od
# tego, ile razy i w jakiej kolejnosci uruchamiano analizy - wiec
# "Przywroc siatke" zawsze odklada elementy tam, skad przyszly.
# Bezpieczenstwo: usuniecie komponentu w HyperMeshu kasuje tez jego
# elementy, dlatego komponent narzedzia jest usuwany TYLKO wtedy, gdy
# policzenie elementow potwierdzi, ze jest pusty. Elementy, ktorych nie
# da sie przywrocic, trafiaja do komponentu zastepczego (nic nie ginie).
class MeshState(object):
    FALLBACK = "HMQS_przywrocone"

    def __init__(self):
        self.orig = {}           # id elementu -> id pierwotnego komponentu (0 = nieznany)
        self.orig_names = {}     # id komponentu -> nazwa (z chwili zapamietania)
        self.own = []            # nazwy komponentow utworzonych przez narzedzie
        self.sets = []           # nazwy zestawow (sets)
        self.tags = []           # etykiety tagow
        self.overlay = {}        # encje dolaczone z pliku REF {typ: [ID]} (usuwane)
        self.owner = ""          # kto ostatnio kolorowal: "mq" / "delta"

    def active(self):
        return bool(self.orig) or bool(self.own)

    def add_own(self, name):
        if name not in self.own:
            self.own.append(name)

    def remember(self, eids):
        """Zapamietuje pierwotne komponenty elementow, ktorych jeszcze nie
        ma w stanie. W GUI - jeden odczyt hurtowy (Tcl); bez niego: duze
        zbiory - przeglad komponentow, male - komponent kazdego elementu."""
        need = [int(e) for e in eids if int(e) not in self.orig]
        if not need:
            return 0
        own = set(self.own)
        comps = HM.components()
        for c in comps:
            try:
                self.orig_names[c.id] = c.name
            except Exception:
                pass
        needset = set(need)
        own_ids = set(cid for cid, nm in self.orig_names.items() if nm in own)
        if HM.bulk_ok():
            D = HM.read_elements(ids=need, names=("collector.id",))
            for eid, cid in zip(D.ids, D.vals.get("collector.id", [])):
                if cid and cid not in own_ids:
                    self.orig[eid] = cid
            for eid in need:
                self.orig.setdefault(eid, 0)
            return len(need)
        total = max(1, HM.count(HM.ent.Element))
        if len(need) > 20000 or len(need) > 0.3 * total:
            for i, c in enumerate(comps):
                if attr(c, "name") in own:
                    continue
                try:
                    els = c.elements or []
                except Exception:
                    els = []
                for e in els:
                    try:
                        eid = e.id
                    except Exception:
                        continue
                    if eid in needset:
                        self.orig[eid] = c.id
                if i % 20 == 0:
                    BUS.progress(T("Zapami\u0119tywanie komponent\u00f3w: %d / %d", "Remembering components: %d / %d", i + 1, len(comps)))
        else:
            for n, eid in enumerate(need):
                try:
                    col = HM.element(eid).collector
                    if col is not None and col.name not in own:
                        self.orig[eid] = col.id
                except Exception:
                    pass
                if n % PROGRESS_EVERY == 0:
                    BUS.progress(T("Zapami\u0119tywanie komponent\u00f3w: %d / %d", "Remembering components: %d / %d", n, len(need)))
        for eid in need:
            self.orig.setdefault(eid, 0)
        return len(need)

    def drop_empty_own(self, keep=()):
        """Usuwa PUSTE komponenty narzedzia spoza listy keep; zwraca pozostale."""
        keep = set(keep)
        idx = HM.comp_index()
        dele, left = [], []
        for nm in self.own:
            if nm in keep:
                left.append(nm)
                continue
            cid = idx.get(nm, 0)
            if not cid:
                continue
            if HM.comp_elem_count(cid) == 0:
                dele.append(cid)
            else:
                left.append(nm)
        if dele:
            HM.delete_comps(dele)
        self.own = left
        return left

    def restore(self, target=""):
        """Przywraca pierwotne komponenty; zwraca (przeniesione, usuniete, zastepcze)."""
        moved = deleted = fallback = 0
        with HM.quiet():
            if self.overlay:
                left = HM.delete_entities(self.overlay)
                if left:
                    BUS.status(T("Nak\u0142adka REF: %d encji nie uda\u0142o si\u0119 usun\u0105\u0107.", "REF overlay: %d entities could not be removed.", left), "warn")
                self.overlay = {}
            idx = HM.comp_index()
            byname = {}
            for cid, nm in self.orig_names.items():
                byname.setdefault(cid, nm)
            live = dict((cid, nm) for nm, cid in idx.items())
            groups = {}
            for eid, cid in self.orig.items():
                if target:
                    nm = target
                else:
                    nm = live.get(cid) or byname.get(cid) or self.FALLBACK
                groups.setdefault(nm, []).append(eid)
            for nm, ids in groups.items():
                if nm not in idx:
                    HM.ensure_comp(nm, index=idx)
                if HM.move_elements(ids, nm):
                    moved += len(ids)
            # komponenty narzedzia: usuwane tylko, gdy na pewno puste. Elementy
            # "obce" (nie z naszych analiz - np. komponent o tej samej nazwie
            # zapisany wczesniej w modelu) zostaja nietkniete razem z komponentem.
            mine_all = set(self.orig)
            idx = HM.comp_index()
            for nm in list(self.own):
                cid = idx.get(nm, 0)
                if not cid:
                    continue
                n = HM.comp_elem_count(cid)
                if n > 0:
                    rest = HM.comp_elem_ids(cid)
                    mine = [e for e in rest if e in mine_all]
                    if mine:
                        HM.ensure_comp(self.FALLBACK, index=idx)
                        HM.move_elements(mine, self.FALLBACK)
                        fallback += len(mine)
                    n = HM.comp_elem_count(cid)
                if n == 0 and HM.delete_comps([cid]):
                    deleted += 1
            if self.sets:
                HM.delete_sets(self.sets)
            if self.tags:
                HM.delete_tags(self.tags)
        HM.show_all_comps(True)
        HM.redraw()
        self.orig, self.orig_names, self.own = {}, {}, []
        self.sets, self.tags, self.owner = [], [], ""
        return moved, deleted, fallback


MESH = MeshState()


# ================== WIELE METRYK -> GRUPY KOLOROW =====================
# Jednoczesne sprawdzenie kilku metryk jakosci siatki:
#     Aspect Ratio, Jacobian Ratio, Jacobian Zero, Skewness.
# W JEDNYM przebiegu po elementach czytane sa wszystkie zaznaczone metryki
# (wartosci zostaja w pamieci). Potem - BEZ ponownego czytania:
#  - PODZIAL WG PROGU: elementy w normie -> jedna grupa bazowa (zielona),
#    poza norma -> pasma kolorow wg przedzialow (np. AR: 4-5 niebieski ...
#    12-15 czerwony), wartosci poza skala -> osobna grupa,
#  - widok per metryka (komponenty w kolorach pasm) albo ZBIORCZY (ile
#    metryk dany element przekracza: 0 = zielony, 1, 2, 3, 4),
#  - ZESTAWY (sets) dla WSZYSTKICH metryk naraz - element moze nalezec do
#    wielu zestawow (komponent moze byc jeden - to on daje kolor),
#  - LEGENDA reczna (liczba kolorow, granice, kolory) albo automatyczna
#    (przedzialy z wartosci skrajnych i liczby pasm) - zmiana legendy
#    przekolorowuje siatke od razu, na danych z pamieci.
# Jacobian Ratio = "jacobian" z HyperMesha (min/max det(J), 0..1, nizszy =
# gorszy). Jacobian Zero = najmniejszy det(J) w narozach (patrz GEOMETRIA).
class MetricDef(object):
    """strict=True: wartosc ROWNA progowi jest juz poza norma (Jacobian Zero:
    det(J) = 0 to element zdegenerowany, np. quad ze zbiegajacymi sie wezlami)."""
    def __init__(self, key, label, tag, direction, attr_name, strict=False):
        self.key, self.label, self.tag, self.dir, self.attr = key, label, tag, direction, attr_name
        self.strict = strict


MQ_METRICS = [
    MetricDef("ar", "Aspect Ratio", "AR", "above", "aspect"),
    MetricDef("jac", "Jacobian Ratio", "JAC", "below", "jacobian"),
    MetricDef("jz", "Jacobian Zero", "JZ", "below", "", strict=True),    # det(J) z wezlow
    MetricDef("skew", "Skewness", "SKEW", "above", "skew"),
]
MQ_ORDER = [m.key for m in MQ_METRICS]
MQ_DEF = dict((m.key, m) for m in MQ_METRICS)


class Group(object):
    """Grupa na siatce = komponent: klucz klasy, nazwa, kolor, etykieta."""

    def __init__(self, key, name, rgb, label):
        self.key, self.name, self.rgb, self.label = key, name, tuple(rgb), label


class MultiMetric(object):
    def __init__(self):
        pm = PaletteMap()
        self.cfg = {}
        for m in MQ_METRICS:
            d = MQ_DEFAULTS.get(m.key, {})
            edges = list(d.get("edges", []))
            k = max(1, len(edges) - 1)
            self.cfg[m.key] = {
                "use": True, "thr_on": True, "thr": d.get("thr", 0.0),
                "mode": d.get("mode", "auto"), "nb": d.get("nb", 5), "edges": edges,
                "colors": default_band_colors(m.dir, k, True),
            }
        self.scope = "all"            # all | displayed
        self.dim2 = True
        self.dim3 = True
        self.prefix = "MQ"
        self.nice = True              # "ladne" liczby w trybie auto
        self.over_sep = True          # poza skala = osobna grupa
        self.make_sets = True
        self.mark_worst = False
        self.hide_ok = False
        self.col_ok = pm.rgb(PAL_OK)
        self.col_over = pm.rgb(PAL_OVER)
        self.col_na = pm.rgb(PAL_NA)
        self.col_comb = [pm.rgb(i) for i in PAL_COMB]
        self.clear_results()

    # ---------------------------------------------------------- pomocnicze
    def clear_results(self):
        self.elems, self.cfgs = [], []
        self.vals, self.stats, self.scales, self.cls = {}, {}, {}, {}
        self.analyzed = []
        self.view = ""
        self.view_comps = []
        self.act_col = {}
        self.jz_note = ""
        self.n2d = self.n3d = self.nskip = 0
        self.sig = ""
        self.when = ""

    def clean_prefix(self):
        p = re.sub(r"[^A-Za-z0-9_]", "_", (self.prefix or "").strip())
        return p or "MQ"

    def selected(self):
        return [m for m in MQ_ORDER if self.cfg[m]["use"]]

    @staticmethod
    def label(v):
        return T("Zbiorczo", "Combined") if v == "all" else MQ_DEF[v].label

    @staticmethod
    def tag(v):
        return "ALL" if v == "all" else MQ_DEF[v].tag

    @staticmethod
    def source_text(m):
        """Skad pochodzi wartosc (tlumaczone przy wyswietlaniu, nie przy analizie)."""
        return T("det(J) z w\u0119z\u0142\u00f3w", "det(J) from nodes") if m == "jz" else MQ_DEF[m].attr

    @staticmethod
    def view_label(v):
        if v == "all":
            return T("Zbiorczo \u2013 liczba metryk poza norm\u0105", "Combined \u2013 number of metrics out of limits")
        return MQ_DEF[v].label

    def thr_metrics(self):
        return [m for m in self.analyzed if self.scales[m].thr_on]

    def combined_ok(self):
        return "all" in self.cls

    def count(self, v, key):
        return self.stats.get(v, {}).get("counts", {}).get(key, 0)

    def views(self):
        return list(self.analyzed) + (["all"] if self.combined_ok() else [])

    # ---------------------------------------------------------- analiza
    def analyze(self):
        """Czyta wszystkie zaznaczone metryki w jednym przebiegu po elementach,
        liczy skale i klasy, tworzy zestawy i naklada pierwszy widok."""
        sel = self.selected()
        if not sel:
            raise ValueError(T("Zaznacz przynajmniej jedn\u0105 metryk\u0119.", "Tick at least one metric."))
        if not (self.dim2 or self.dim3):
            raise ValueError(T("Zaznacz elementy 2D i/lub 3D.", "Tick 2D and/or 3D elements."))
        if not HM.ok():
            raise RuntimeError(T("Brak API HyperMesha (uruchom narz\u0119dzie w HyperMesh 2023+).",
                                 "No HyperMesh API (run the tool inside HyperMesh 2023+)."))
        HM.refresh_palette()
        if not MESH.active():
            old = HM.leftover_comps([self.clean_prefix(), DELTA.clean_prefix("dAR")])
            if old:
                BUS.log(T("Uwaga: model zawiera komponenty z wcze\u015bniejszego kolorowania (zapisany pokolorowany model?): %s. "
                          "S\u0105 traktowane jak zwyk\u0142e komponenty \u2013 po \u201ePrzywr\u00f3\u0107 siatk\u0119\u201d ich elementy trafi\u0105 do %s.",
                          "Note: the model contains components from an earlier coloring (a saved colored model?): %s. "
                          "They are treated as regular components \u2013 after \u201cRestore mesh\u201d their elements go to %s.",
                          ", ".join(old[:6]) + (" \u2026" if len(old) > 6 else ""), MESH.FALLBACK))
        BUS.progress(T("Odczyt element\u00f3w\u2026", "Reading elements\u2026"))
        dims = set(d for d, on in (("2d", self.dim2), ("3d", self.dim3)) if on)
        want_jz = "jz" in sel
        D = HM.read_elements(displayed=(self.scope == "displayed"), keep=lambda c: elem_dim(c) in dims,
                             names=[MQ_DEF[m].attr for m in sel if m != "jz"], nodes=want_jz)
        if not D.all_ids:
            raise ValueError(T("Brak element\u00f3w w wybranym zakresie.", "No elements in the selected scope."))
        E, C, skip = D.ids, D.cfg, D.skip
        if not E:
            raise ValueError(T("Brak element\u00f3w 2D/3D wybranego typu (pomini\u0119to %d).",
                               "No 2D/3D elements of the selected type (%d skipped).", skip))
        n2 = sum(1 for c in C if elem_dim(c) == "2d")
        n3 = len(C) - n2
        vals = dict((m, D.vals[MQ_DEF[m].attr]) for m in sel if m != "jz")
        jz_fam = []
        if want_jz:
            # Jacobian Zero z naroznikow: wspolrzedne wezlow czytane raz, hurtowo
            BUS.progress(T("Wsp\u00f3\u0142rz\u0119dne w\u0119z\u0142\u00f3w\u2026", "Node coordinates\u2026"))
            xyz = HM.read_nodes(n for ns in D.nodes for n in ns)
            jz = []
            for i, cfg in enumerate(C):
                shape = SHAPES.get(cfg)
                ns = D.nodes[i]
                pts = None
                if shape and len(ns) >= CORNERS[shape]:
                    pts = [xyz.get(n) for n in ns[:CORNERS[shape]]]
                    if None in pts:
                        pts = None
                jz.append(jacobian_zero(pts, shape) if pts else None)
                jz_fam.append(shape or "")
                if i % (PROGRESS_EVERY * 5) == 0:
                    BUS.progress(T("Jacobian Zero: %d / %d", "Jacobian Zero: %d / %d", i, len(C)))
            vals["jz"] = jz
        # Jacobian Zero: znak bryl wg orientacji wiekszosci elementow danego typu
        self.jz_note = ""
        if want_jz:
            pos, neg = {}, {}
            for v, f in zip(vals["jz"], jz_fam):
                if v is None or f in ("", "tria", "quad"):
                    continue
                if v < 0:
                    neg[f] = neg.get(f, 0) + 1
                elif v > 0:
                    pos[f] = pos.get(f, 0) + 1
            flip = [f for f in neg if neg[f] > pos.get(f, 0)]
            if flip:
                vals["jz"] = [(-v if (v is not None and f in flip) else v) for v, f in zip(vals["jz"], jz_fam)]
                self.jz_note = T("orientacja w\u0119z\u0142\u00f3w odwr\u00f3cona dla: %s", "node orientation flipped for: %s", ", ".join(flip))
        stats, av = {}, []
        for m in sel:
            vs = vals[m]
            good = [(v, eid) for v, eid in zip(vs, E) if v is not None]
            if not good:
                stats[m] = {"n": 0, "nNa": len(vs)}
                continue
            vmin, emin = min(good)
            vmax, emax = max(good)
            stats[m] = {"n": len(good), "nNa": len(vs) - len(good), "min": vmin, "max": vmax,
                        "minEl": emin, "maxEl": emax, "mean": sum(v for v, _ in good) / len(good)}
            av.append(m)
        if not av:
            raise ValueError(T("\u017badna metryka nie zwr\u00f3ci\u0142a warto\u015bci.", "No metric returned values."))
        self.elems, self.cfgs, self.vals, self.stats = E, C, vals, stats
        self.analyzed = av
        self.n2d, self.n3d, self.nskip = n2, n3, skip
        self.recompute()
        self.sig = HM.signature()
        self.when = now_text()
        if self.make_sets:
            BUS.progress(T("Zestawy element\u00f3w dla wszystkich metryk\u2026", "Element sets for all metrics\u2026"))
            self.build_sets()
        v0 = self.view
        if v0 not in self.views():
            v0 = self.analyzed[0]
        self.apply_view(v0)
        return True

    def recompute(self):
        """Skale, klasy i liczniki (po zmianie legendy - bez czytania z HM)."""
        self.cls = {}
        for m in self.analyzed:
            st = self.stats[m]
            sc = Scale.build(self.cfg[m], MQ_DEF[m].dir, st.get("min"), st.get("max"), self.nice, MQ_DEF[m].strict)
            self.scales[m] = sc
            cl = [sc.classify(v, self.over_sep) for v in self.vals[m]]
            self.cls[m] = cl
            cnt = {}
            for k in cl:
                cnt[k] = cnt.get(k, 0) + 1
            st["counts"] = cnt
            st["nOk"] = cnt.get(Scale.OK, 0)
            st["nBad"] = (len(cl) - st["nOk"] - cnt.get(Scale.NA, 0)) if sc.thr_on else 0
            worse_high = sc.dir == "above"
            st["worstEl"] = st.get("maxEl") if worse_high else st.get("minEl")
            st["worstV"] = st.get("max") if worse_high else st.get("min")
        tm = self.thr_metrics()
        if tm:
            lists = [self.cls[m] for m in tm]
            out, cnt = [], {}
            for i in range(len(self.elems)):
                f, anyv = 0, False
                for lst in lists:
                    k = lst[i]
                    if k == Scale.NA:
                        continue
                    anyv = True
                    if k != Scale.OK:
                        f += 1
                out.append(f if anyv else Scale.NA)
                key = out[-1]
                cnt[key] = cnt.get(key, 0) + 1
            n = len(self.elems)
            self.cls["all"] = out
            self.stats["all"] = {"counts": cnt, "nOk": cnt.get(0, 0), "nNa": cnt.get(Scale.NA, 0),
                                 "nBad": n - cnt.get(0, 0) - cnt.get(Scale.NA, 0),
                                 "n": n - cnt.get(Scale.NA, 0)}
        else:
            self.stats.pop("all", None)

    # ---------------------------------------------------------- grupy i legenda
    def comb_word(self, k):
        if lang() == "pl":
            if k == 1:
                return "1 metryka poza norm\u0105"
            if 2 <= k <= 4:
                return "%d metryki poza norm\u0105" % k
            return "%d metryk poza norm\u0105" % k
        return "1 metric out of limits" if k == 1 else "%d metrics out of limits" % k

    def group_defs(self, v):
        """Grupy (komponenty) widoku w kolejnosci legendy - najgorsze u gory."""
        P = self.clean_prefix()
        out = []
        if v == "all":
            M = len(self.thr_metrics())
            for k in range(M, 0, -1):
                col = self.col_comb[min(k, len(self.col_comb) - 1)]
                out.append(Group(k, "%s_ALL_%d_poza" % (P, k), col, self.comb_word(k)))
            out.append(Group(0, "%s_ALL_0_OK" % P, self.col_ok, T("wszystkie w normie", "all within limits")))
            out.append(Group(Scale.NA, "%s_ALL_nd" % P, self.col_na, T("n/d (brak warto\u015bci)", "n/a (no value)")))
            return out
        sc = self.scales[v]
        E, dec, tg = sc.edges, sc.dec, self.tag(v)
        lo0, hiN = fmt_num(E[0], dec), fmt_num(E[-1], dec)
        bands = []
        for i in sc.worst_first():
            a, b = E[i], E[i + 1]
            nm = "%s_%s_s%02d_%s_%s" % (P, tg, sc.band_rank(i), sanit_num(a, dec), sanit_num(b, dec))
            bands.append(Group(i, nm, sc.colors[i], "%s \u2013 %s" % (fmt_num(a, dec), fmt_num(b, dec))))
        if sc.thr_on:
            tt = fmt_num(sc.thr, max(dec, 2))
            if self.over_sep:
                out.append(Group(Scale.OVER, "%s_%s_poza_skala" % (P, tg), self.col_over,
                                 ("> %s" % hiN) if sc.dir == "above" else ("< %s" % lo0)))
            out += bands
            out.append(Group(Scale.OK, "%s_%s_OK" % (P, tg), self.col_ok,
                             T("%s %s (w normie)", "%s %s (within limits)", sc.ok_op(), tt)))
        else:
            hi = Group(Scale.OVER, "%s_%s_powyzej" % (P, tg), self.col_over, "> %s" % hiN)
            lo = Group(Scale.UNDER, "%s_%s_ponizej" % (P, tg), self.col_over, "< %s" % lo0)
            if self.over_sep:
                out.append(hi if sc.dir == "above" else lo)
            out += bands
            if self.over_sep:
                out.append(lo if sc.dir == "above" else hi)
        out.append(Group(Scale.NA, "%s_%s_nd" % (P, tg), self.col_na, T("n/d (brak warto\u015bci)", "n/a (no value)")))
        return out

    def legend_rows(self, v):
        """Wiersze legendy (okno, podglad slajdu, PPTX): kolor FAKTYCZNY z palety."""
        tot = len(self.elems)
        pm = palette_map()
        out = []
        for g in self.group_defs(v):
            n = self.count(v, g.key)
            if g.key in (Scale.NA, Scale.OVER, Scale.UNDER) and n == 0:
                continue
            rgb = self.act_col.get((v, g.key)) or pm.snap(g.rgb)
            out.append({"key": g.key, "rgb": rgb, "label": g.label, "count": n,
                        "pct": pct(n, tot), "comp": g.name})
        return out

    def legend_head(self, v):
        if v == "all":
            return (self.view_label("all"),
                    T("metryki z progiem: %s", "metrics with a threshold: %s", ", ".join(self.tag(m) for m in self.thr_metrics())))
        sc = self.scales[v]
        md = T("skala r\u0119czna", "manual scale") if sc.mode == "manual" else T("skala automatyczna", "automatic scale")
        if sc.thr_on:
            sub = T("poza norm\u0105: %s %s \u2022 %s \u2022 %d pasm", "out of limits: %s %s \u2022 %s \u2022 %d bands",
                    sc.bad_op(), fmt_num(sc.thr, max(sc.dec, 2)), md, sc.k)
        else:
            sub = T("bez progu \u2022 %s \u2022 %d pasm", "no threshold \u2022 %s \u2022 %d bands", md, sc.k)
        return (self.label(v), sub)

    def stat_pairs(self, v):
        """Statystyki widoku jako pary (etykieta, wartosc) - tabela na slajdzie."""
        tot = len(self.elems)
        if v == "all":
            s = self.stats["all"]
            return [(T("Elementy", "Elements"), "%d" % tot),
                    (T("W normie (wszystkie metryki)", "Within limits (all metrics)"), pct_text(s["nOk"], tot)),
                    (T("Poza norm\u0105 (\u2265 1 metryka)", "Out of limits (\u2265 1 metric)"), pct_text(s["nBad"], tot))]
        s, sc = self.stats[v], self.scales[v]
        dec = max(sc.dec, 2)
        out = [(T("Elementy z warto\u015bci\u0105", "Elements with a value"), "%d" % s["n"])]
        if sc.thr_on:
            out.append((T("W normie", "Within limits"), pct_text(s["nOk"], tot)))
            out.append((T("Poza norm\u0105", "Out of limits"), pct_text(s["nBad"], tot)))
        out.append(("Min", "%s  (el. %s)" % (fmt_num(s["min"], dec), s["minEl"])))
        out.append(("Max", "%s  (el. %s)" % (fmt_num(s["max"], dec), s["maxEl"])))
        return out

    def thr_text(self, v):
        if v == "all" or v not in self.scales:
            return ""
        sc = self.scales[v]
        if not sc.thr_on:
            return T("bez progu", "no threshold")
        return T("poza norm\u0105 %s %s", "out of limits %s %s", sc.bad_op(), fmt_num(sc.thr, 3))

    # ---------------------------------------------------------- widok na siatce
    def apply_view(self, v):
        """Naklada widok (metryka albo "all") na siatke: komponenty w kolorach grup."""
        if not self.analyzed:
            raise ValueError(T("Najpierw wykonaj analiz\u0119.", "Run the analysis first."))
        if v == "all" and not self.combined_ok():
            raise ValueError(T("Widok zbiorczy wymaga co najmniej jednej metryki z progiem.",
                               "The combined view needs at least one metric with a threshold."))
        if v != "all" and v not in self.analyzed:
            v = self.analyzed[0]
        if HM.signature() != self.sig:
            self.clear_results()
            raise ValueError(T("Model w sesji si\u0119 zmieni\u0142 (inny plik / liczba element\u00f3w) \u2013 wykonaj analiz\u0119 ponownie.",
                               "The model in the session has changed (another file / element count) \u2013 run the analysis again."))
        if MESH.owner == "delta":
            DELTA.done = False      # kolory delty zostana zastapione
        MESH.remember(self.elems)
        MESH.owner = "mq"
        BUS.progress(T("Kolorowanie: %s\u2026", "Coloring: %s\u2026", self.label(v)))
        buckets = {}
        for k, eid in zip(self.cls[v], self.elems):
            buckets.setdefault(k, []).append(eid)
        with HM.quiet():
            pm = palette_map()
            idx = HM.comp_index()
            new, fail, ok_name = [], 0, ""
            for g in self.group_defs(v):
                ids = buckets.get(g.key, [])
                if not ids and (g.key < 0 or v == "all"):
                    continue          # pasma zostaja (pelna skala w Model Browser), reszta tylko niepuste
                ci = pm.index_of(g.rgb)
                self.act_col[(v, g.key)] = pm.rgb(ci)
                existed = g.name in idx
                try:
                    cid = HM.ensure_comp(g.name, ci, index=idx)
                except Exception:
                    fail += 1
                    continue
                MESH.add_own(g.name)
                if ids and not HM.move_elements(ids, g.name):
                    fail += 1
                if existed:
                    HM.show_comps([cid], True)
                if (v != "all" and g.key == Scale.OK) or (v == "all" and g.key == 0):
                    ok_name = g.name
                new.append(g.name)
            # kontrola: wszystkie elementy analizy musza siedziec w komponentach widoku
            idx = HM.comp_index()
            placed = max(0, HM.comp_elem_count([idx[n] for n in new if n in idx]))
            lost = len(self.elems) - placed
            MESH.drop_empty_own(keep=new)
        self.view_comps = new
        if self.hide_ok and ok_name in idx:
            HM.show_comps([idx[ok_name]], False)
        if MESH.tags:
            HM.delete_tags(MESH.tags)
            MESH.tags = []
        if self.mark_worst and v != "all":
            st = self.stats[v]
            if st.get("worstEl"):
                lab = "%s_%s_WORST" % (self.clean_prefix(), self.tag(v))
                if HM.create_tag(st["worstEl"], lab, "%s=%s (el. %s)" % (self.tag(v), fmt_num(st["worstV"], 4), st["worstEl"])):
                    MESH.tags.append(lab)
        self.view = v
        HM.redraw()
        msg = T("Widok: %s \u2013 %d grup, %d element\u00f3w.", "View: %s \u2013 %d groups, %d elements.", self.label(v), len(new), len(self.elems))
        level = "ok"
        if fail:
            msg += T(" UWAGA: %d grup nie uda\u0142o si\u0119 utworzy\u0107 / wype\u0142ni\u0107.", " WARNING: %d groups could not be created / filled.", fail)
            level = "warn"
        if lost > 0:
            msg += T(" UWAGA: %d element\u00f3w nie trafi\u0142o do grup widoku.", " WARNING: %d elements did not reach the view groups.", lost)
            level = "warn"
        BUS.status(msg, level)
        return True

    def set_hide_ok(self, hide):
        self.hide_ok = bool(hide)
        if not self.view:
            return
        idx = HM.comp_index()
        ids = [idx[n] for n in self.view_comps if n.endswith("_OK") and n in idx]
        HM.show_comps(ids, not self.hide_ok)
        HM.redraw()

    # ---------------------------------------------------------- zestawy (sets)
    def build_sets(self):
        """Zestawy elementow dla WSZYSTKICH metryk (i widoku zbiorczego)."""
        if MESH.sets:
            HM.delete_sets(MESH.sets)
            MESH.sets = []
        P = self.clean_prefix()
        ok = bad = 0
        for v in self.views():
            buckets = {}
            for k, eid in zip(self.cls[v], self.elems):
                buckets.setdefault(k, []).append(eid)
            all_bad = []
            for g in self.group_defs(v):
                if g.key in (Scale.OK, Scale.NA) or (v == "all" and g.key == 0):
                    continue
                ids = buckets.get(g.key, [])
                if not ids:
                    continue
                sn = P + "S" + g.name[len(P):]
                if HM.create_set(sn, ids):
                    MESH.sets.append(sn)
                    ok += 1
                else:
                    bad += 1
                all_bad += ids
            if all_bad and v != "all":
                sn = "%sS_%s_POZA_NORMA" % (P, self.tag(v))
                if HM.create_set(sn, all_bad):
                    MESH.sets.append(sn)
                    ok += 1
                else:
                    bad += 1
        if bad and not ok:
            BUS.status(T("Zestaw\u00f3w (sets) nie uda\u0142o si\u0119 utworzy\u0107 \u2013 komponenty dzia\u0142aj\u0105 normalnie.",
                         "Sets could not be created \u2013 components work normally."), "warn")
        return ok

    # ---------------------------------------------------------- przywracanie
    def restore(self):
        moved, deleted, fb = MESH.restore()
        self.view = ""
        self.view_comps = []
        msg = T("Przywr\u00f3cono pierwotne komponenty (%d element\u00f3w). Dane analizy zostaj\u0105 \u2013 widok mo\u017cna na\u0142o\u017cy\u0107 ponownie bez czytania.",
                "Original components restored (%d elements). Analysis data kept \u2013 a view can be re-applied without re-reading.", moved)
        if fb:
            msg += T(" UWAGA: %d element\u00f3w w komponencie %s.", " WARNING: %d elements in component %s.", fb, MESH.FALLBACK)
        BUS.status(msg, "warn" if fb else "ok")

    # ---------------------------------------------------------- tabele
    def summary_rows(self):
        """Tabela wszystkich metryk: naglowek + wiersze (komorki tekst/dict)."""
        hdr = [T("Metryka", "Metric"), T("Pr\u00f3g", "Threshold"), T("Elementy", "Elements"),
               T("W normie", "Within"), T("Poza norm\u0105", "Out"), T("% poza", "% out"), "Min", "Max",
               T("Najgorszy el.", "Worst el.")]
        rows = [hdr]
        tot = len(self.elems)
        for m in self.analyzed:
            s, sc = self.stats[m], self.scales[m]
            dec = max(sc.dec, 2)
            if sc.thr_on:
                thr_t = "%s %s" % (sc.ok_op(), fmt_num(sc.thr, dec))
                nb = s["nBad"]
                cells = [thr_t, "%d" % s["n"], "%d" % s["nOk"],
                         {"t": "%d" % nb, "color": "C00000" if nb else "2E7D32", "bold": True},
                         "%.2f%%" % pct(nb, tot)]
            else:
                cells = ["\u2013", "%d" % s["n"], "\u2013", "\u2013", "\u2013"]
            rows.append([self.label(m)] + cells + [fmt_num(s["min"], 4), fmt_num(s["max"], 4), "%s" % s["worstEl"]])
        if self.combined_ok():
            s = self.stats["all"]
            nb = s["nBad"]
            rows.append([{"t": T("Zbiorczo (\u2265 1 metryka)", "Combined (\u2265 1 metric)"), "bold": True}, "", "%d" % tot,
                         "%d" % s["nOk"], {"t": "%d" % nb, "color": "C00000" if nb else "2E7D32", "bold": True},
                         "%.2f%%" % pct(nb, tot), "", "", ""])
        return rows

    # ---------------------------------------------------------- ustawienia
    KEYS = ("scope", "dim2", "dim3", "prefix", "nice", "over_sep", "make_sets", "mark_worst", "hide_ok")

    def to_dict(self):
        d = dict((k, getattr(self, k)) for k in self.KEYS)
        d["colors"] = {"ok": list(self.col_ok), "over": list(self.col_over), "na": list(self.col_na)}
        d["metrics"] = dict((m, {"use": c["use"], "thr_on": c["thr_on"], "thr": c["thr"], "mode": c["mode"],
                                 "nb": c["nb"], "edges": list(c["edges"]), "colors": [list(x) for x in c["colors"]]})
                            for m, c in self.cfg.items())
        return d

    def from_dict(self, d):
        assign_attrs(self, d, self.KEYS, {"scope": ("all", "displayed")})
        cols = d.get("colors") if isinstance(d.get("colors"), dict) else {}
        for nm, key in (("ok", "col_ok"), ("over", "col_over"), ("na", "col_na")):
            if isinstance(cols.get(nm), list) and len(cols[nm]) == 3:
                setattr(self, key, tuple(int(x) for x in cols[nm]))
        mets = d.get("metrics") if isinstance(d.get("metrics"), dict) else {}
        for m, c in mets.items():
            if m not in self.cfg or not isinstance(c, dict):
                continue
            for k in ("use", "thr_on"):
                if isinstance(c.get(k), bool):
                    self.cfg[m][k] = c[k]
            if c.get("mode") in ("auto", "manual"):
                self.cfg[m]["mode"] = c["mode"]
            if isinstance(c.get("nb"), int) and not isinstance(c.get("nb"), bool):
                self.cfg[m]["nb"] = max(1, min(40, c["nb"]))
            if to_float(c.get("thr")) is not None:
                self.cfg[m]["thr"] = to_float(c["thr"])
            if isinstance(c.get("edges"), list):
                self.cfg[m]["edges"] = [to_float(x) for x in c["edges"] if to_float(x) is not None]
            if isinstance(c.get("colors"), list):
                self.cfg[m]["colors"] = [tuple(max(0, min(255, int(v))) for v in x) for x in c["colors"]
                                         if isinstance(x, (list, tuple)) and len(x) == 3
                                         and all(isinstance(v, (int, float)) for v in x)]


MQ = MultiMetric()


# ===================== DELTA JAKOSCI REF <-> INF ======================
# Dwa pliki .hm: REF (baza) i INF (po zmianie ksztaltu). Wczytywanie
# SEKWENCYJNE: najpierw REF (zapamietuje metryke kazdego elementu wg ID),
# potem INF. Dla elementu INF bierzemy element REF o TYM SAMYM ID i liczymy
# POGORSZENIE D (zawsze "w gore" = gorzej):
#     Aspect Ratio, Skewness:  D = Q(INF) - Q(REF)
#     Jacobian:                D = Q(REF) - Q(INF)   (spadek = gorzej)
#     Przesuniecie [mm]:       D = maks. |xyz(INF) - xyz(REF)| wezlow elementu
# Siatka INF jest kolorowana gradientem (pasma), D ponizej progu "bez zmian"
# (w tym poprawa) -> SZARY, element bez odpowiednika -> MAGENTA.
# Tryby dodatkowe:
#  - JEDEN MODEL: podzial elementow na pasma wg WARTOSCI metryki,
#  - PODZIAL WG PROGU (jeden model): poza norma -> <prefiks>_poza_norma
#    (czerwony), reszta -> <prefiks>_w_normie (zielony),
#  - "pokaz poprawione" (magenta), nakladka REF (plik REF dolaczony do sesji
#    jako ukryty komponent), zbieranie pozostalych elementow INF,
#  - etykiety MIN / MAX (tagi) z wspolrzednymi srodka elementu,
#  - wyciszanie "bez zmian" (przezroczystosc, a gdy sie nie da - ukrycie).
# UWAGA: wczytanie plikow ZASTEPUJE model w sesji (okno pyta przed startem).
DELTA_METRICS = [("ar", "Aspect Ratio", "aspect"), ("jac", "Jacobian", "jacobian"),
                 ("skew", "Skewness", "skew"), ("disp", "", "")]


class DeltaEngine(object):
    AUTO_PREFIXES = ("dAR", "dJac", "dSkew", "dXYZ", "AR", "Jac", "Skew", "XYZ", "Q")

    def __init__(self):
        self.ref_file = ""
        self.inf_file = ""
        self.use_open_inf = False     # INF = aktualnie otwarty model (przeladowany z dysku)
        self.single_only = False      # analizuj tylko jeden model (REF)
        self.fail_mode = False        # podzial wg progu (jeden model)
        self.fail_limit = 5.0
        self.metric = "ar"            # ar | jac | skew | disp
        self.dim = "2d"               # 2d | 3d
        self.prefix = "dAR"
        self.restore_target = ""      # opcjonalnie: wszystko do jednego komponentu
        self.auto_scale = True
        self.deadband = 0.01          # prog "bez zmian" / tolerancja [mm]
        self.band_count = 10
        self.nice_round = True
        self.manual_bounds = []       # reczna skala: K+1 granic rosnaco
        self.mark_extremes = True
        self.fade_gray = True
        self.fade_level = 85
        self.show_improved = False
        self.mk_ref_comp = False      # dolacz REF jako ukryta nakladka
        self.mk_inf_comp = False      # pozostale elementy INF w komponencie
        self.clear_results()

    # ---------------------------------------------------------- tryb
    def mode(self):
        """delta (dwa pliki) | single (jeden model: pasma wg wartosci) |
        fail (jeden model: podzial wg progu)."""
        if self.fail_mode:
            return "fail"
        return "single" if self.single_only else "delta"

    def set_mode(self, mode):
        self.fail_mode = mode == "fail"
        self.single_only = mode == "single"

    def cur_metric(self):
        """Metryka OSTATNIEJ analizy (gdy jest wynik) - legendy, tagi i
        "Sprawdz delte" opisuja wynik na siatce, nie biezacy wybor w oknie."""
        return self.last_metric if (self.done and self.last_metric) else self.metric

    def cur_dim(self):
        return self.last_dim if (self.done and self.last_dim) else self.dim

    def clear_results(self):
        self.bounds, self.k, self.uniform = [], 0, True
        self.inf_q, self.single_q = {}, {}      # id -> wartosc (INF / jeden model)
        self.disp_node = {}                      # id -> wezel o najwiekszym przesunieciu
        self.unmatched_set = set()
        self.when = ""
        self.band_rgb = []
        self.clamped = 0
        self.counts = {"band": 0, "gray": 0, "unm": 0, "imp": 0}
        pm = PaletteMap()
        self.rgb_gray = pm.rgb(DELTA_GRAY_IDX)
        self.rgb_unm = pm.rgb(DELTA_UNMATCHED_IDX)
        self.rgb_imp = pm.rgb(DELTA_UNMATCHED_IDX)
        self.rgb_ok, self.rgb_bad = pm.rgb(PAL_OK), pm.rgb(3)
        self.min_info = self.max_info = None
        self.fail_info = None
        self.single = False
        self.fail_active = False
        self.last_metric = self.last_dim = ""
        self.ref_q, self.ref_xyz = {}, {}
        self.delta = {}
        self.buckets = {}
        self.comps = []
        self.gray_name, self.gray_shown, self.gray_faded = "", True, ""
        self.ref_comp, self.ref_shown, self.inf_shown = "", False, True
        self.inf_rest = ""
        self.done = False
        self.model_path = ""
        self.skipped = (0, 0)

    # ---------------------------------------------------------- nazwy
    def metric_label(self, m=None):
        m = m or self.cur_metric()
        return {"ar": "Aspect Ratio", "jac": "Jacobian", "skew": "Skewness"}.get(
            m, T("Przesuni\u0119cie [mm]", "Displacement [mm]"))

    def delta_label(self):
        m = self.cur_metric()
        if self.single:
            return self.metric_label(m)
        return {"ar": "\u0394AR", "jac": T("spadek Jacobianu", "Jacobian drop"), "skew": "\u0394Skew"}.get(
            m, T("przesuni\u0119cie [mm]", "displacement [mm]"))

    def delta_tag(self):
        base = {"ar": "AR", "jac": "Jac", "skew": "Skew", "disp": "XYZ"}[self.cur_metric()]
        return base if self.single else "d" + base

    def legend_title(self):
        m = self.cur_metric()
        if self.single:
            return "%s %s" % (self.metric_label(m), T("(warto\u015b\u0107)", "(value)"))
        return {"ar": T("Pogorszenie Aspect Ratio (INF \u2212 REF)", "Aspect Ratio worsening (INF \u2212 REF)"),
                "jac": T("Spadek Jacobianu (REF \u2212 INF)", "Jacobian drop (REF \u2212 INF)"),
                "skew": T("Pogorszenie Skewness (INF \u2212 REF)", "Skewness worsening (INF \u2212 REF)")}.get(
            m, T("Przesuni\u0119cie w\u0119z\u0142\u00f3w REF \u2192 INF [mm]", "Node displacement REF \u2192 INF [mm]"))

    def sign(self):
        return -1.0 if self.cur_metric() == "jac" else 1.0

    def fail_below(self):
        return self.cur_metric() == "jac"

    def fail_op(self):
        return "<" if self.fail_below() else ">"

    def dim_label(self):
        return "3D" if self.cur_dim() == "3d" else "2D"

    def on_metric_change(self):
        """Domyslny prefiks podaza za metryka (jesli nie zmieniony recznie)."""
        if (self.prefix or "").strip() in self.AUTO_PREFIXES or not (self.prefix or "").strip():
            self.prefix = {"ar": "dAR", "jac": "dJac", "skew": "dSkew", "disp": "dXYZ"}[self.metric]

    def clean_prefix(self, default):
        p = re.sub(r"[^A-Za-z0-9_]", "_", (self.prefix or "").strip())
        return p or default

    # ---------------------------------------------------------- odczyt z HM
    def q_attr(self):
        return {"ar": "aspect", "jac": "jacobian", "skew": "skew"}.get(self.cur_metric())

    def _elements(self, which):
        """Hurtowy odczyt elementow wybranego wymiaru (ElemData): wartosc
        metryki albo - dla przesuniecia - listy wezlow."""
        want = "3d" if self.cur_dim() == "3d" else "2d"
        BUS.progress(T("%s: odczyt element\u00f3w %s\u2026", "%s: reading %s elements\u2026", which, self.dim_label()))
        a = self.q_attr()
        return HM.read_elements(keep=lambda c: elem_dim(c) == want, names=[a] if a else [],
                                nodes=self.cur_metric() == "disp")

    def _values(self, D):
        """Slownik id -> wartosc metryki (bez brakow)."""
        vs = D.vals.get(self.q_attr()) or []
        return dict((eid, v) for eid, v in zip(D.ids, vs) if v is not None)

    def _q(self, e):
        a = self.q_attr()
        return num_attr(e, a) if a else None

    # ---------------------------------------------------------- skala
    def compute_bounds(self, maxpos):
        if not self.auto_scale and len(self.manual_bounds) >= 3:
            self.bounds = [float(x) for x in self.manual_bounds]
            self.k = len(self.bounds) - 1
            self.uniform = False
            return "manual"
        K = max(2, min(96, int(self.band_count or 10)))
        lo = max(0.0, to_float(self.deadband, 0.01))
        hi = maxpos
        if self.nice_round:
            hi = nice_ceil(hi)
        if hi <= lo + 1e-12:
            hi = lo + 0.1
        step = (hi - lo) / float(K)
        self.bounds = [lo + i * step for i in range(K + 1)]
        self.k = K
        self.uniform = True
        return "auto"

    def band_of(self, d):
        """Pasmo wartosci D (-1 = ponizej skali); D >= gornej granicy -> ostatnie."""
        bb, K = self.bounds, self.k
        if K < 1 or len(bb) != K + 1 or d < bb[0]:
            return -1
        if d >= bb[-1]:
            return K - 1
        return max(0, min(K - 1, bisect.bisect_right(bb, d) - 1))

    def band_colors(self, k):
        pm = PaletteMap()
        return [pm.rgb(i) for i in ramp_indices(k, RAMP_FULL)]

    # ---------------------------------------------------------- start
    def run(self):
        """Wybiera tryb (delta / jeden model / podzial wg progu) i uruchamia analize."""
        if not HM.ok():
            raise RuntimeError(T("Brak API HyperMesha.", "No HyperMesh API."))
        HM.refresh_palette()
        mode = self.mode()
        have_ref = bool(self.ref_file) and os.path.isfile(self.ref_file)
        if self.use_open_inf:
            p = HM.model_file()
            if not p:
                raise ValueError(T("Nie ustalono pliku otwartego modelu \u2013 zapisz model albo wska\u017c plik INF.",
                                   "The open model file is unknown \u2013 save the model or choose the INF file."))
            self.inf_file = p
        have_inf = bool(self.inf_file) and os.path.isfile(self.inf_file)
        if mode in ("single", "fail"):
            if have_ref:
                return self.run_single(self.ref_file, True)
            if have_inf:
                return self.run_single(self.inf_file, not self.use_open_inf)
            raise ValueError(T("Wska\u017c plik modelu (.hm) \u2013 REF albo INF (lub \u201eINF = otwarty model\u201d).",
                               "Choose the model file (.hm) \u2013 REF or INF (or \u201cINF = open model\u201d)."))
        if not (have_ref and have_inf):
            raise ValueError(T("Delta wymaga DW\u00d3CH modeli: REF (.hm) i INF (.hm) albo \u201eINF = otwarty model\u201d. "
                               "Dla jednego modelu wybierz tryb \u201ejeden model\u201d.",
                               "The delta needs TWO models: REF (.hm) and INF (.hm) or \u201cINF = open model\u201d. "
                               "For one model choose the \u201cone model\u201d mode."))
        return self.run_delta()

    def _begin(self, single):
        if MESH.active():
            MESH.restore()          # zdejmuje tez kolory "Wiele metryk"
        MQ.view, MQ.view_comps = "", []
        self.clear_results()
        self.single = single
        self.last_metric, self.last_dim = self.metric, self.dim

    def _load(self, path, which):
        BUS.progress(T("Wczytywanie %s: %s\u2026", "Loading %s: %s\u2026", which, os.path.basename(path)))
        ok, msg = HM.read_file(path)
        if not ok:
            raise RuntimeError(T("Nie uda\u0142o si\u0119 wczyta\u0107 %s: %s", "Could not load %s: %s", which, msg))
        HM.redraw()

    # ---------------------------------------------------------- jeden model
    def run_single(self, path, reload):
        if self.metric == "disp":
            self.metric = "ar"
            self.on_metric_change()
            BUS.status(T("Przesuni\u0119cie wymaga dw\u00f3ch modeli \u2013 dla jednego modelu u\u017cyto Aspect Ratio.",
                         "Displacement needs two models \u2013 Aspect Ratio used for one model."), "warn")
        if self.metric == "disp":
            raise ValueError(T("Przesuni\u0119cie wymaga dw\u00f3ch modeli (REF i INF).", "Displacement needs two models (REF and INF)."))
        self._begin(True)
        if reload:
            self._load(path, T("modelu", "model"))
        self.model_path = path
        D = self._elements(T("model", "model"))
        skip = D.skip
        if not D.ids:
            raise ValueError(T("Brak element\u00f3w %s w modelu (pomini\u0119to %d).", "No %s elements in the model (%d skipped).", self.dim_label(), skip))
        vals = self._values(D)
        if not vals:
            raise ValueError(T("Metryka %s niedost\u0119pna dla element\u00f3w %s.", "Metric %s unavailable for %s elements.", self.metric_label(), self.dim_label()))
        self.single_q = vals
        self.skipped = (skip, 0)
        vmin_e = min(vals, key=lambda k: vals[k])
        vmax_e = max(vals, key=lambda k: vals[k])
        vmin, vmax = vals[vmin_e], vals[vmax_e]
        P = self.clean_prefix("Q")
        groups = []
        if self.fail_mode:
            lim = to_float(self.fail_limit)
            if lim is None:
                raise ValueError(T("Pr\u00f3g musi by\u0107 liczb\u0105.", "The threshold must be a number."))
            below = self.fail_below()
            lo, hi = vmin, vmax
            if hi <= lo:
                hi = lo + 1.0
            if lim <= lo:
                lo = lim - (hi - lim) * 0.1 - 1e-6
            if lim >= hi:
                hi = lim + (lim - lo) * 0.1 + 1e-6
            self.bounds, self.k = [lo, lim, hi], 2
            bad = [k for k, v in vals.items() if (v < lim if below else v > lim)]
            badset = set(bad)
            ok = [k for k in vals if k not in badset]
            dec = dec_for((hi - lo) / 2.0)
            self.fail_info = (len(bad), len(ok), fmt_num(lim, dec), self.fail_op())
            groups.append(("%s_w_normie" % P, self.rgb_ok, ok))
            groups.append(("%s_poza_norma" % P, self.rgb_bad, bad))
            self.buckets = {"ok": ok, "bad": bad}
            self.band_rgb = [self.rgb_bad, self.rgb_ok] if below else [self.rgb_ok, self.rgb_bad]
            self.fail_active = True
            self.counts["band"] = len(vals)
        else:
            K = max(2, min(96, int(self.band_count or 10)))
            lo, hi = vmin, vmax
            if hi <= lo:
                hi = lo + 1.0
            if self.nice_round:
                hi = nice_ceil(hi) if hi > 0 else hi
            width = (hi - lo) / float(K) or 1.0
            self.bounds, self.k = [lo + i * width for i in range(K + 1)], K
            dec = dec_for(width)
            cols = self.band_colors(K)
            bk = dict((i, []) for i in range(K))
            for eid, v in vals.items():
                bk[max(0, min(K - 1, int(math.floor((v - lo) / width))))].append(eid)
            for i in range(K):
                nm = "%s_b%02d_%s_%s" % (P, i, sanit_num(self.bounds[i], dec), sanit_num(self.bounds[i + 1], dec))
                groups.append((nm, cols[i], bk[i]))
            self.buckets = bk
            self.band_rgb = cols
            self.counts["band"] = len(vals)
        self.max_info = (vmax, vmax_e)
        self.min_info = (vmin, vmin_e) if vmin_e != vmax_e else None
        self._paint(groups, D.ids)
        self._finish(P)
        return True

    # ---------------------------------------------------------- delta REF/INF
    def run_delta(self):
        self._begin(False)
        # FAZA 1: REF -> metryka (albo wspolrzedne wezlow) wg ID
        self._load(self.ref_file, "REF")
        R = self._elements("REF")
        rskip = R.skip
        if not R.ids:
            raise ValueError(T("REF: brak element\u00f3w %s (pomini\u0119to %d).", "REF: no %s elements (%d skipped).", self.dim_label(), rskip))
        if self.metric == "disp":
            BUS.progress(T("REF: wsp\u00f3\u0142rz\u0119dne w\u0119z\u0142\u00f3w\u2026", "REF: node coordinates\u2026"))
            self.ref_xyz = HM.read_nodes(n for ns in R.nodes for n in ns)
            if not self.ref_xyz:
                raise ValueError(T("Nie uda\u0142o si\u0119 odczyta\u0107 wsp\u00f3\u0142rz\u0119dnych w\u0119z\u0142\u00f3w.", "Could not read node coordinates."))
        else:
            self.ref_q = self._values(R)
            if not self.ref_q:
                raise ValueError(T("Metryka %s niedost\u0119pna w REF.", "Metric %s unavailable in REF.", self.metric_label()))
        # FAZA 2: INF -> dopasowanie po ID, delta
        self._load(self.inf_file, "INF")
        I = self._elements("INF")
        iskip, all_ids = I.skip, I.all_ids
        if not I.ids:
            raise ValueError(T("INF: brak element\u00f3w %s (pomini\u0119to %d).", "INF: no %s elements (%d skipped).", self.dim_label(), iskip))
        self.skipped = (rskip, iskip)
        sgn = self.sign()
        unmatched = []
        if self.cur_metric() == "disp":
            BUS.progress(T("INF: wsp\u00f3\u0142rz\u0119dne w\u0119z\u0142\u00f3w\u2026", "INF: node coordinates\u2026"))
            xyz = HM.read_nodes(n for ns in I.nodes for n in ns)
            for eid, ns in zip(I.ids, I.nodes):
                d, nid = self.max_disp_node(ns, xyz)
                if d is None:
                    unmatched.append(eid)
                else:
                    self.delta[eid] = d
                    self.disp_node[eid] = nid
        else:
            vs = I.vals.get(self.q_attr()) or []
            for eid, q in zip(I.ids, vs):
                if q is None:
                    continue
                self.inf_q[eid] = q
                r = self.ref_q.get(eid)
                if r is None:
                    unmatched.append(eid)
                    continue
                self.delta[eid] = sgn * (q - r)
        self.unmatched_set = set(unmatched)
        if not self.delta and not unmatched:
            raise ValueError(T("Brak warto\u015bci do por\u00f3wnania.", "No values to compare."))
        maxd_e = max(self.delta, key=lambda k: self.delta[k]) if self.delta else None
        mind_e = min(self.delta, key=lambda k: self.delta[k]) if self.delta else None
        maxpos = max(0.0, self.delta[maxd_e]) if maxd_e is not None else 0.0
        self.compute_bounds(maxpos)
        K = self.k
        tol = max(0.0, to_float(self.deadband, 0.0))
        bk = dict((i, []) for i in range(K))
        gray, imp = [], []
        hi = self.bounds[-1]
        for eid, d in self.delta.items():
            b = self.band_of(d)
            if b < 0:
                if self.show_improved and d < 0 and d <= -tol:
                    imp.append(eid)
                else:
                    gray.append(eid)
                continue
            if d - hi > 1e-12:
                self.clamped += 1
            bk[b].append(eid)
        self.buckets = bk
        self.counts = {"band": sum(len(v) for v in bk.values()), "gray": len(gray), "unm": len(unmatched), "imp": len(imp)}
        P = self.clean_prefix("dAR")
        dec = dec_for((self.bounds[-1] - self.bounds[0]) / float(K))
        cols = self.band_colors(K)
        self.band_rgb = cols
        groups = [("%s_bez_zmian" % P, self.rgb_gray, gray)]
        self.gray_name = groups[0][0]
        for i in range(K):
            nm = "%s_b%02d_%s_%s" % (P, i, sanit_num(self.bounds[i], dec), sanit_num(self.bounds[i + 1], dec))
            groups.append((nm, cols[i], bk[i]))
        if unmatched:
            groups.append(("%s_bez_odpowiednika" % P, self.rgb_unm, unmatched))
        if imp:
            groups.append(("%s_poprawione" % P, self.rgb_imp, imp))
        inf_ids = I.ids
        self._paint(groups, inf_ids)
        # pozostale elementy INF (inne wymiary) we wlasnym komponencie
        if self.mk_inf_comp and iskip > 0:
            keep = set(inf_ids)
            rest = [i for i in all_ids if i not in keep]
            if rest:
                nm = "%s_INF_pozostale" % P
                MESH.remember(rest)
                HM.ensure_comp(nm, 14)
                MESH.add_own(nm)
                HM.move_elements(rest, nm)
                self.inf_rest = nm
                self.comps.append(nm)
        if maxd_e is not None:
            self.max_info = (self.delta[maxd_e], maxd_e)
            if mind_e is not None and mind_e != maxd_e:
                self.min_info = (self.delta[mind_e], mind_e)
        # wyciszenie "bez zmian"
        if self.fade_gray and gray:
            idx = HM.comp_index()
            gid = idx.get(self.gray_name, 0)
            if gid and HM.set_transparency([gid], max(0, min(100, int(to_float(self.fade_level, 85))))):
                self.gray_faded = "transp"
            elif gid:
                HM.show_comps([gid], False)
                self.gray_shown, self.gray_faded = False, "hidden"
        # nakladka REF
        if self.mk_ref_comp:
            self._ref_overlay(P)
        self._finish(P)
        return True

    def _ref_overlay(self, P):
        """REF dolaczony do sesji jako ukryty komponent (pokaz / ukryj).
        Wszystkie encje scalone z pliku sa zapamietane i usuwane przy
        przywracaniu - model wraca dokladnie do stanu sprzed nakladki."""
        BUS.progress(T("Do\u0142\u0105czanie REF jako nak\u0142adki\u2026", "Adding REF as an overlay\u2026"))
        with HM.quiet():
            ents = HM.merge_file(self.ref_file)
            new = ents.get("Element") or []
            if not new:
                BUS.status(T("Nie uda\u0142o si\u0119 do\u0142\u0105czy\u0107 REF do sesji \u2013 analiza delty jest kompletna, bez nak\u0142adki.",
                             "Could not add REF to the session \u2013 the delta analysis is complete, without the overlay."), "warn")
                return
            nm = "%s_REF_model" % P
            cid = HM.ensure_comp(nm, 31)
            MESH.add_own(nm)
            HM.move_elements(new, nm)
            # scalone komponenty REF sa teraz puste - od razu precz z Model Browsera
            empty = [c for c in ents.get("Component", []) if HM.comp_elem_count(c) == 0]
            if empty:
                HM.delete_comps(empty)
            ents["Component"] = [c for c in ents.get("Component", []) if c not in set(empty)]
            MESH.overlay = ents
            HM.show_comps([cid], False)
        self.ref_comp, self.ref_shown = nm, False

    def max_disp(self, nids, xyz):
        """Najwieksze przesuniecie wezlow elementu wzgledem REF (None = brak
        wspolnych wezlow)."""
        return self.max_disp_node(nids, xyz)[0]

    def max_disp_node(self, nids, xyz):
        """(najwieksze przesuniecie, ID wezla) albo (None, None)."""
        best, bn = None, None
        for nid in nids:
            r = self.ref_xyz.get(nid)
            p = xyz.get(nid)
            if r is None or p is None:
                continue
            d = math.sqrt((p[0] - r[0]) ** 2 + (p[1] - r[1]) ** 2 + (p[2] - r[2]) ** 2)
            if best is None or d > best:
                best, bn = d, nid
        return best, bn

    def _elem_disp(self, e, cache=None):
        """Przesuniecie jednego elementu (sprawdzanie delty)."""
        nids = HM.elem_node_ids(e)
        return self.max_disp(nids, HM.read_nodes(nids))

    # ---------------------------------------------------------- kolorowanie
    def _paint(self, groups, all_ids):
        """Wspolne kolorowanie: zapamietanie stanu, komponenty, przeniesienie."""
        BUS.progress(T("Zapami\u0119tywanie pierwotnych komponent\u00f3w\u2026", "Remembering the original components\u2026"))
        MESH.remember(all_ids)
        MESH.owner = "delta"
        with HM.quiet():
            pm = palette_map()
            idx = HM.comp_index()
            fail = 0
            for n, (name, rgb, ids) in enumerate(groups):
                ci = pm.index_of(rgb)
                try:
                    HM.ensure_comp(name, ci, index=idx)
                except Exception:
                    fail += 1
                    continue
                MESH.add_own(name)
                self.comps.append(name)
                if ids and not HM.move_elements(ids, name):
                    fail += 1
                if n % 8 == 0:
                    BUS.progress(T("Kolorowanie: grupa %d / %d", "Coloring: group %d / %d", n + 1, len(groups)))
        # faktyczne kolory (paleta) do legendy
        self.band_rgb = [pm.snap(c) for c in self.band_rgb]
        self.rgb_gray, self.rgb_unm, self.rgb_imp = pm.snap(self.rgb_gray), pm.snap(self.rgb_unm), pm.snap(self.rgb_imp)
        self.rgb_ok, self.rgb_bad = pm.snap(self.rgb_ok), pm.snap(self.rgb_bad)
        MESH.drop_empty_own(keep=self.comps)
        self.paint_fail = fail

    def _finish(self, P):
        if self.mark_extremes and self.max_info:
            self._label_extremes(P)
        self.done = True
        self.when = now_text()
        HM.redraw()

    def _label_extremes(self, P):
        pm = palette_map()
        items = [("MAX", self.max_info, 8)] + ([("MIN", self.min_info, 2)] if self.min_info else [])
        for tag, info, col in items:
            d, eid = info[0], info[1]
            xyz = HM.elem_centroid(eid)
            info_full = (d, eid, xyz)
            if tag == "MAX":
                self.max_info = info_full
            else:
                self.min_info = info_full
            lab = "%s_%s" % (P, tag)
            body = "%s %s=%s (el. %s)" % (tag, self.delta_tag(), fmt_num(d, 4), eid)
            if HM.create_tag(eid, lab, body, 3):
                MESH.tags.append(lab)
            else:
                nm = "%s_%s_%s" % (P, tag, ("%+.4f" % d).replace("+", "plus_").replace("-", "minus_").replace(".", "p"))
                HM.ensure_comp(nm, col)
                MESH.add_own(nm)
                self.comps.append(nm)
                HM.move_elements([eid], nm)

    # ---------------------------------------------------------- widocznosc
    def toggle_gray(self):
        idx = HM.comp_index()
        gid = idx.get(self.gray_name, 0)
        if not gid:
            raise ValueError(T("Brak komponentu \u201ebez zmian\u201d z ostatniej analizy.", "No \u201cno change\u201d component from the last analysis."))
        self.gray_shown = not self.gray_shown
        HM.show_comps([gid], self.gray_shown)
        HM.redraw()
        return self.gray_shown

    def toggle_ref(self):
        idx = HM.comp_index()
        cid = idx.get(self.ref_comp, 0)
        if not cid:
            raise ValueError(T("Brak nak\u0142adki REF (zaznacz opcj\u0119 przed analiz\u0105).", "No REF overlay (tick the option before the analysis)."))
        self.ref_shown = not self.ref_shown
        HM.show_comps([cid], self.ref_shown)
        HM.redraw()
        return self.ref_shown

    def toggle_inf(self):
        if not self.done:
            raise ValueError(T("Najpierw wykonaj analiz\u0119.", "Run the analysis first."))
        self.inf_shown = not self.inf_shown
        idx = HM.comp_index()
        HM.show_all_comps(self.inf_shown)
        if self.inf_shown:
            if not self.gray_shown and self.gray_name in idx:
                HM.show_comps([idx[self.gray_name]], False)
            if not self.ref_shown and self.ref_comp in idx:
                HM.show_comps([idx[self.ref_comp]], False)
        elif self.ref_shown and self.ref_comp in idx:
            HM.show_comps([idx[self.ref_comp]], True)
        HM.redraw()
        return self.inf_shown

    # ---------------------------------------------------------- diagnostyka
    def inspect(self, eid):
        """Tekst z Q(REF), Q(INF), delta i pasmem elementu - z danych OSTATNIEJ
        analizy (bez odczytu z HyperMesha), wiec dziala takze po zmianie
        metryki / typu w oknie i nie zalezy od API atrybutow elementu."""
        eid = int(eid)
        if not self.done:
            raise ValueError(T("Najpierw wykonaj analiz\u0119.", "Run the analysis first."))
        m, dl = self.cur_metric(), self.dim_label()
        scope = T("Element %d nie nale\u017cy do analizowanego zakresu (%s, %s) albo nie ma go w modelu.",
                  "Element %d is not in the analysed scope (%s, %s) or does not exist in the model.", eid, dl, self.metric_label(m))
        if self.single:
            v = self.single_q.get(eid)
            if v is None:
                raise ValueError(scope)
            d = v
            head = "%s = %s" % (self.metric_label(m), fmt_num(v, 5))
        elif m == "disp":
            d = self.delta.get(eid)
            if d is None:
                if eid in self.unmatched_set:
                    raise ValueError(T("Element %d nie ma wsp\u00f3lnych w\u0119z\u0142\u00f3w z REF (bez odpowiednika).",
                                       "Element %d has no nodes in common with REF (unmatched).", eid))
                raise ValueError(scope)
            head = T("przesuni\u0119cie = %s mm", "displacement = %s mm", fmt_num(d, 5))
            nid = self.disp_node.get(eid)
            if nid is not None:
                head += T("   (w\u0119ze\u0142 %d)", "   (node %d)", nid)
        else:
            r, q = self.ref_q.get(eid), self.inf_q.get(eid)
            if q is None:
                raise ValueError(scope)
            if r is None:
                raise ValueError(T("Element %d nie ma odpowiednika w REF (Q(INF) = %s).",
                                   "Element %d has no counterpart in REF (Q(INF) = %s).", eid, fmt_num(q, 5)))
            d = self.sign() * (q - r)
            head = "Q(REF) = %s   Q(INF) = %s   \u0394 = %+.5f   D = %+.5f" % (fmt_num(r, 5), fmt_num(q, 5), q - r, d)
        if self.fail_active and self.fail_info:
            lim = to_float(self.bounds[1])
            bad = d < lim if self.fail_below() else d > lim
            band = T("poza norm\u0105", "out of limits") if bad else T("w normie", "within limits")
        else:
            b = self.band_of(d)
            if b < 0:
                band = T("szary \u2013 bez zmian", "gray \u2013 no change")
                if self.show_improved and d < 0 and d <= -abs(to_float(self.deadband, 0.0)):
                    band = T("poprawiony (magenta)", "improved (magenta)")
            else:
                dec = dec_for((self.bounds[-1] - self.bounds[0]) / float(self.k))
                band = T("pasmo %d / %d: %s \u2013 %s", "band %d / %d: %s \u2013 %s", b + 1, self.k,
                         fmt_num(self.bounds[b], dec), fmt_num(self.bounds[b + 1], dec))
        return T("Element %d (%s, %s)\n%s\nGrupa: %s", "Element %d (%s, %s)\n%s\nGroup: %s",
                 eid, dl, self.metric_label(m), head, band)

    # ---------------------------------------------------------- legenda
    def legend_model(self):
        """Dane legendy (okno, podglad, PPTX): pasma + wiersze dodatkowe."""
        if not self.done or self.k < 1:
            return None
        bb, K = self.bounds, self.k
        dec = dec_for((bb[-1] - bb[0]) / float(K))
        bands = []
        for i in range(K):
            if self.fail_active:
                key = ("bad" if i == 0 else "ok") if self.fail_below() else ("ok" if i == 0 else "bad")
            else:
                key = i
            n = len(self.buckets.get(key, []))
            rgb = self.band_rgb[i] if i < len(self.band_rgb) else (128, 128, 128)
            bands.append({"rgb": rgb, "lo": bb[i], "hi": bb[i + 1], "count": n,
                          "label": "%s \u2013 %s" % (fmt_num(bb[i], dec), fmt_num(bb[i + 1], dec))})
        extra = []
        if self.fail_active and self.fail_info:
            nf, nok, lim, op = self.fail_info
            extra.append({"rgb": self.rgb_bad, "label": T("poza norm\u0105 (%s %s)", "out of limits (%s %s)", op, lim), "count": nf})
            extra.append({"rgb": self.rgb_ok, "label": T("w normie", "within limits"), "count": nok})
        elif not self.single:
            g = T("bez zmian", "no change")
            if self.gray_faded == "transp":
                g += T(" (przezroczyste)", " (transparent)")
            elif self.gray_faded == "hidden":
                g += T(" (ukryte)", " (hidden)")
            extra.append({"rgb": self.rgb_gray, "label": g, "count": self.counts["gray"]})
        if self.counts.get("imp"):
            extra.append({"rgb": self.rgb_imp, "label": T("poprawione", "improved"), "count": self.counts["imp"]})
        if self.counts.get("unm"):
            extra.append({"rgb": self.rgb_unm, "label": T("bez odpowiednika w REF", "unmatched in REF"), "count": self.counts["unm"]})
        mode = T("skala automatyczna", "automatic scale") if (self.auto_scale or self.single) else T("skala r\u0119czna", "manual scale")
        return {"title": self.legend_title(), "sub": "%s \u2022 %d %s" % (mode, K, T("pasm", "bands")),
                "bands": bands, "extra": extra, "dec": dec, "clamped": self.clamped,
                "max": self.max_info, "min": self.min_info, "fail": self.fail_active}

    def stat_pairs(self):
        out = []
        for tag, info in (("MAX", self.max_info), ("MIN", self.min_info)):
            if info:
                out.append(("%s %s" % (tag, self.delta_label()), "%s  (el. %s)" % (fmt_num(info[0], 4), info[1])))
        c = self.counts
        if not self.single:
            out.append((T("W skali / bez zmian", "In scale / no change"), "%d / %d" % (c["band"], c["gray"])))
            if c.get("unm"):
                out.append((T("Bez odpowiednika w REF", "Unmatched in REF"), "%d" % c["unm"]))
        return out

    # ---------------------------------------------------------- przywracanie
    def restore(self):
        moved, deleted, fb = MESH.restore(target=(self.restore_target or "").strip())
        self.done = False
        msg = T("Przywr\u00f3cono siatk\u0119 (%d element\u00f3w, usuni\u0119to %d komponent\u00f3w).", "Mesh restored (%d elements, %d components removed).", moved, deleted)
        if fb:
            msg += T(" UWAGA: %d element\u00f3w w %s.", " WARNING: %d elements in %s.", fb, MESH.FALLBACK)
        BUS.status(msg, "warn" if fb else "ok")

    KEYS = ("ref_file", "inf_file", "use_open_inf", "single_only", "fail_mode", "fail_limit", "metric", "dim",
            "prefix", "restore_target", "auto_scale", "deadband", "band_count", "nice_round", "manual_bounds",
            "mark_extremes", "fade_gray", "fade_level", "show_improved", "mk_ref_comp", "mk_inf_comp")

    def to_dict(self):
        return dict((k, getattr(self, k)) for k in self.KEYS)

    def from_dict(self, d):
        assign_attrs(self, d, self.KEYS, {"metric": ("ar", "jac", "skew", "disp"), "dim": ("2d", "3d")})
        mb = [to_float(x) for x in self.manual_bounds] if isinstance(self.manual_bounds, list) else []
        self.manual_bounds = mb if all(x is not None for x in mb) else []
        self.band_count = max(2, min(96, int(self.band_count)))
        self.fade_level = max(0, min(100, int(self.fade_level)))


DELTA = DeltaEngine()


# ======================== RAPORT JAKOSCI SIATKI ========================
# Pelny raport jakosci siatki: dla wszystkich / wyswietlonych elementow
# wybranych wymiarow (1D / 2D / 3D) i wybranych metryk:
#  - N, min / max / srednia / mediana / odch. std., percentyle P5 / P95,
#    liczba elementow poza progiem ("do poprawy"), pelny histogram,
#    lista TOP najgorszych elementow (ID + wartosc),
#  - syntetyczny WSKAZNIK JAKOSCI 0-100 z ocena A-F (srednia udzialow
#    elementow "w normie" po metrykach z progiem),
#  - TREND: kazdy raport dopisuje wpis do pliku <baza>.qrtrend, kolejny
#    raport pokazuje zmiane wyniku wzgledem poprzedniego,
#  - opcjonalnie: podzial per komponent, tabela per element, objetosc 3D
#    i pole 2D modelu,
#  - POROWNANIE DWOCH SIATEK REF vs INF: osobne raporty _REF / _INF
#    i raport porownawczy (TXT, HTML, XLSX) z kontrola zgodnosci ID,
#    typow, topologii (wezly) i wartosci metryk.
# Wartosci czytane sa atrybutami elementow (e.aspect, e.jacobian ...) -
# zla nazwa daje brak wartosci, nigdy crash sesji. TET COLLAPSE liczony jest
# ze wspolrzednych wezlow (HM 2024 nie ma go w API).
class RepMetric(object):
    def __init__(self, key, pl, en, dpl, den, thr, wdir, bins, col, abbr):
        self.key, self.pl, self.en, self.dpl, self.den = key, pl, en, dpl, den
        self.thr, self.wdir, self.bins, self.col, self.abbr = thr, wdir, bins, col, abbr

    @property
    def label(self):
        return self.en if lang() == "en" else self.pl

    @property
    def desc(self):
        return self.den if lang() == "en" else self.dpl


# klucz, etykieta PL / EN, opis PL / EN, prog, kierunek, biny histogramu,
# kolumna CSV/XLSX, skrot. Progi i biny mozna zmienic tutaj.
REPORT_METRICS = [
    RepMetric("aspectratio", "ASPECT RATIO", "ASPECT RATIO", "wy\u017cszy = gorszy", "higher = worse",
              5.0, "above", [1, 1.5, 2, 3, 4, 5, 7, 10], "AspectRatio", "AR"),
    RepMetric("jacobian", "JACOBIAN", "JACOBIAN", "ni\u017cszy = gorszy, zakres ~0..1", "lower = worse, range ~0..1",
              0.6, "below", [0, 0.3, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0], "Jacobian", "Jac"),
    RepMetric("skew", "SKEW (sko\u015bno\u015b\u0107)", "SKEW", "wy\u017cszy = gorszy, stopnie", "higher = worse, degrees",
              60.0, "above", [0, 10, 20, 30, 45, 60, 75], "Skew", "Skew"),
    RepMetric("warpage", "WARPAGE (wypaczenie)", "WARPAGE", "wy\u017cszy = gorszy, stopnie", "higher = worse, degrees",
              10.0, "above", [0, 5, 10, 15, 30], "Warpage", "Warp"),
    RepMetric("taper", "TAPER (zw\u0119\u017cenie)", "TAPER", "wy\u017cszy = gorszy", "higher = worse",
              0.5, "above", [0, 0.1, 0.3, 0.5, 0.7], "Taper", "Tap"),
    RepMetric("minangle", "K\u0104T MIN. WEWN\u0118TRZNY", "MIN. INTERIOR ANGLE", "ni\u017cszy = gorszy, stopnie", "lower = worse, degrees",
              30.0, "below", [0, 20, 30, 45, 60], "MinAngle", "MinA"),
    RepMetric("maxangle", "K\u0104T MAX. WEWN\u0118TRZNY", "MAX. INTERIOR ANGLE", "wy\u017cszy = gorszy, stopnie", "higher = worse, degrees",
              135.0, "above", [90, 120, 135, 150, 180], "MaxAngle", "MaxA"),
    RepMetric("tetcollapse", "TET COLLAPSE", "TET COLLAPSE", "ni\u017cszy = gorszy, zakres ~0..1", "lower = worse, range ~0..1",
              0.1, "below", [0, 0.1, 0.2, 0.3, 0.5, 1.0], "TetCollapse", "Tet"),
    RepMetric("length", "MIN. D\u0141UGO\u015a\u0106 KRAW\u0118DZI", "MIN. EDGE LENGTH", "min. d\u0142ugo\u015b\u0107 kraw\u0119dzi", "min. edge length",
              0.0, "none", [], "MinLength", "Len"),
]
REP = dict((m.key, m) for m in REPORT_METRICS)
REP_ORDER = [m.key for m in REPORT_METRICS]
REP_ATTR = {"aspectratio": "aspect", "jacobian": "jacobian", "skew": "skew", "warpage": "warpage",
            "taper": "taper", "minangle": "minangle", "maxangle": "maxangle"}


def applies_metric(key, dim, shape):
    if key in ("aspectratio", "jacobian", "skew"):
        return dim in ("2d", "3d")
    if key == "length":
        return dim in ("1d", "2d", "3d")
    if key == "warpage":
        # tetra ma tylko trojkatne sciany - warpage zawsze 0 (zawyzaloby statystyki)
        return (dim == "2d" and shape == "quad") or (dim == "3d" and shape != "tet")
    if key == "taper":
        return dim == "2d" and shape == "quad"
    if key in ("minangle", "maxangle"):
        return dim == "2d"
    if key == "tetcollapse":
        return dim == "3d" and shape == "tet"
    return False


def data_name(key):
    """Nazwa zrodla danych metryki (kolumna "dataname" w raporcie)."""
    if key == "tetcollapse":
        return T("liczony z w\u0119z\u0142\u00f3w", "computed from nodes")
    if key == "length":
        return "length / shortestside"
    return REP_ATTR.get(key, key)


def stats_of(vals):
    n = len(vals)
    if not n:
        return {"n": 0}
    s = sorted(vals)
    mean = sum(s) / n
    med = s[n // 2] if n % 2 else (s[n // 2 - 1] + s[n // 2]) / 2.0
    sd = math.sqrt(sum((v - mean) ** 2 for v in s) / (n - 1)) if n > 1 else 0.0
    return {"n": n, "min": s[0], "max": s[-1], "mean": mean, "median": med, "sd": sd,
            "p5": percentile(s, 0.05), "p95": percentile(s, 0.95)}


def percentile(sorted_vals, q):
    """Percentyl (interpolacja liniowa) z listy JUZ posortowanej rosnaco."""
    n = len(sorted_vals)
    if n == 0:
        return None
    if n == 1:
        return sorted_vals[0]
    pos = q * (n - 1)
    lo = int(pos)
    if lo >= n - 1:
        return sorted_vals[-1]
    return sorted_vals[lo] + (pos - lo) * (sorted_vals[lo + 1] - sorted_vals[lo])


def histogram(vals, edges_in):
    """(krawedzie, liczniki, ponizej, powyzej); puste biny -> 8 z danych.
    Przedzialy [a, b), OSTATNI domkniety [a, b] - wartosc rowna gornej
    krawedzi (np. Jacobian = 1.0 elementu idealnego) nie jest "powyzej"."""
    if len(edges_in) < 2:
        lo, hi = min(vals), max(vals)
        if hi <= lo:
            hi = lo + 1.0
        step = (hi - lo) / 8.0
        edges = [lo + i * step for i in range(9)]
    else:
        edges = list(edges_in)
    counts = [0] * (len(edges) - 1)
    under = over = 0
    e0, eN = edges[0], edges[-1]
    for v in vals:
        if v < e0:
            under += 1
        elif v > eN:
            over += 1
        elif v == eN:
            counts[-1] += 1
        else:
            counts[bisect.bisect_right(edges, v) - 1] += 1
    return edges, counts, under, over


def is_bad(v, wdir, thr):
    return (wdir == "above" and v > thr) or (wdir == "below" and v < thr)


def worst_list(vals, ids, wdir, topn):
    pairs = list(zip(vals, ids))
    pairs.sort(key=lambda p: p[0], reverse=(wdir == "above"))
    return pairs[:topn]


def compute_score(exp):
    """Wskaznik 0-100: srednia udzialow "w normie" po metrykach z progiem."""
    parts = []
    for m in exp:
        if m["dir"] == "none" or m["n"] <= 0:
            continue
        parts.append((m["key"], 100.0 * (1.0 - float(m["bad"]) / m["n"])))
    if not parts:
        return {"has": False}
    score = sum(p[1] for p in parts) / len(parts)
    grade = "A" if score >= 95 else "B" if score >= 85 else "C" if score >= 70 else "D" if score >= 50 else "F"
    return {"has": True, "score": score, "grade": grade, "parts": parts}


def trend_load(path):
    """Historia trendu: plik JSON lines; czyta tez stary format makra Tcl."""
    out = []
    if not os.path.isfile(path):
        return out
    try:
        with io.open(path, "r", encoding="utf-8") as fh:
            for line in fh:
                t = line.strip()
                if not t or t.startswith("#"):
                    continue
                try:
                    d = json.loads(t)
                except ValueError:
                    m = re.search(r"score\s+(\S+)", t)
                    ts = re.search(r"ts\s+\{([^}]*)\}", t) or re.search(r"ts\s+(\S+)", t)
                    if not m:
                        continue
                    d = {"score": to_float(m.group(1)), "ts": ts.group(1) if ts else ""}
                if to_float(d.get("score")) is not None:
                    d["score"] = to_float(d["score"])
                    out.append(d)
    except Exception:
        pass
    return out


def trend_append(path, entry):
    try:
        new = not os.path.isfile(path) or os.path.getsize(path) == 0
        with io.open(path, "a", encoding="utf-8", newline="\n") as fh:
            if new:
                fh.write("# HM Quality Studio - trend kolejnych raportow (ts / score / total / badtot)\n")
            fh.write(json.dumps(entry) + "\n")
        return True
    except Exception:
        return False


class ReportEngine(object):
    def __init__(self):
        self.src = "current"          # current | file | compare
        self.hm_file = ""
        self.ref_file = ""
        self.inf_file = ""
        self.scope = "all"            # all | displayed
        self.out_file = ""            # baza nazwy plikow wynikowych
        self.fmt = {"txt": True, "csv": True, "xlsx": True, "html": True}
        self.per_elem = True
        self.per_comp = False
        self.num_style = "pl"         # pl: 1,23 ; | en: 1.23 ,
        self.level = "full"           # full | basic
        self.use_metric = dict((k, k in ("aspectratio", "jacobian", "skew")) for k in REP_ORDER)
        self.use_dim = {"1d": False, "2d": True, "3d": True}
        self.chk_ids = True
        self.chk_topo = True
        self.chk_vol = True
        self.chk_nodes = True         # zestawienie wezlow A vs B (wspolrzedne)
        self.node_tol = 1e-4          # tolerancja polozenia wezla [mm]
        self.cmp_tol = 1e-6
        self.topn = TOPN
        self.last = None              # wynik ostatniej analizy (slajd PPTX)
        self.last_cmp = None          # ostatnie porownanie elementow A vs B
        self.last_ncmp = None         # ostatnie zestawienie wezlow A vs B
        self.last_files = []          # zapisane pliki ostatniego raportu
        self.trend_prev = None
        self.trend_hist = []

    def dims_label(self):
        l = [d.upper() for d in ("1d", "2d", "3d") if self.use_dim.get(d)]
        return " + ".join(l) if l else "-"

    def tol(self):
        t = to_float(self.cmp_tol)
        return t if t is not None and t >= 0 else 1e-6

    def ntol(self):
        t = to_float(self.node_tol)
        return t if t is not None and t >= 0 else 1e-4

    # ---------------------------------------------------------- analiza
    def analyze(self, source="", sig=False):
        """Jeden przebieg po elementach -> slownik wynikow (run)."""
        if not HM.ok():
            raise RuntimeError(T("Brak API HyperMesha.", "No HyperMesh API."))
        dims = set(d for d in ("1d", "2d", "3d") if self.use_dim.get(d))
        keys = [k for k in REP_ORDER if self.use_metric.get(k)]
        do_topo = sig and self.chk_topo
        need_nodes = do_topo or "tetcollapse" in keys
        cache = {}

        def names_for(cfg):
            # dane HM potrzebne dla elementu danego typu (tylko te, ktore maja sens)
            if cfg in cache:
                return cache[cfg]
            d, shape = elem_dim(cfg), elem_shape(cfg)
            out = []
            for k in keys:
                if k == "tetcollapse" or not applies_metric(k, d, shape):
                    continue
                out.append(("length" if d == "1d" else "shortestside") if k == "length" else REP_ATTR[k])
            if self.per_comp:
                out.append("collector.id")
            if self.chk_vol and d in ("2d", "3d"):
                out.append("volume" if d == "3d" else "area")
            cache[cfg] = out
            return out

        BUS.progress(T("Odczyt element\u00f3w\u2026", "Reading elements\u2026"))
        D = HM.read_elements(displayed=(self.scope == "displayed"), keep=lambda c: elem_dim(c) in dims,
                             names=names_for, nodes=need_nodes)
        total = len(D.all_ids)
        if not total:
            raise ValueError(T("Brak element\u00f3w w wybranym zakresie.", "No elements in the selected scope."))
        if not D.ids:
            raise ValueError(T("Brak element\u00f3w wybranych wymiar\u00f3w (%s).", "No elements of the selected dimensions (%s).", self.dims_label()))
        skip = D.skip
        d1 = d2 = d3 = 0
        types = {}
        for cfg in D.cfg:
            d = elem_dim(cfg)
            tn = elem_type_name(cfg)
            types[tn] = types.get(tn, 0) + 1
            if d == "1d":
                d1 += 1
            elif d == "2d":
                d2 += 1
            else:
                d3 += 1
        xyz = {}
        nodes_all = None
        nnodes = HM.count(HM.ent.Node)
        if sig and self.chk_nodes:
            # zestawienie wezlow A vs B: wspolrzedne WSZYSTKICH wezlow modelu
            BUS.progress(T("Wsp\u00f3\u0142rz\u0119dne w\u0119z\u0142\u00f3w (zestawienie w\u0119z\u0142\u00f3w A vs B)\u2026", "Node coordinates (node comparison A vs B)\u2026"))
            nids = HM.entity_ids("Node") or []
            nodes_all = HM.read_nodes(nids) if nids else {}
            xyz = nodes_all
        if "tetcollapse" in keys and not xyz:
            BUS.progress(T("Wsp\u00f3\u0142rz\u0119dne w\u0119z\u0142\u00f3w (tet collapse)\u2026", "Node coordinates (tet collapse)\u2026"))
            xyz = HM.read_nodes(n for c, ns in zip(D.cfg, D.nodes) if elem_shape(c) == "tet" for n in ns[:4])
        comp_names = HM.comp_names() if self.per_comp else {}
        # siatka pokolorowana przez narzedzie: elementy siedza w grupach
        # MQ_ / dAR_ - statystyki per komponent licz dla komponentow PIERWOTNYCH
        orig = MESH.orig if (self.per_comp and MESH.orig) else {}
        if orig:
            for cid, nm in MESH.orig_names.items():
                comp_names.setdefault(cid, nm)
        values = dict((k, []) for k in keys)
        vids = dict((k, []) for k in keys)
        by_comp, comp_n = {}, {}
        comp_order = []
        rows = [] if self.per_elem else None
        sigs = {} if sig else None
        vol = area = 0.0
        voln = arean = 0
        V = D.vals
        N = len(D.ids)
        for i, eid in enumerate(D.ids):
            cfg = D.cfg[i]
            d, shape = elem_dim(cfg), elem_shape(cfg)
            cell = {}
            for k in keys:
                if not applies_metric(k, d, shape):
                    continue
                if k == "tetcollapse":
                    ns = D.nodes[i]
                    pts = [xyz.get(n) for n in ns[:4]] if len(ns) >= 4 else None
                    v = tet_collapse(pts) if pts and None not in pts else None
                elif k == "length":
                    v = V["length" if d == "1d" else "shortestside"][i]
                else:
                    v = V[REP_ATTR[k]][i]
                if v is None:
                    continue
                values[k].append(v)
                vids[k].append(eid)
                cell[k] = v
            if self.per_comp:
                cid = orig.get(eid) or V["collector.id"][i] or 0
                if cid not in comp_n:
                    comp_names.setdefault(cid, T("(brak)", "(none)"))
                    comp_order.append(cid)
                comp_n[cid] = comp_n.get(cid, 0) + 1
                for k, v in cell.items():
                    by_comp.setdefault((cid, k), []).append(v)
            if self.chk_vol:
                if d == "3d":
                    v = V["volume"][i]
                    if v is not None:
                        vol += v
                        voln += 1
                elif d == "2d":
                    v = V["area"][i]
                    if v is not None:
                        area += v
                        arean += 1
            if rows is not None:
                rows.append([eid, elem_type_name(cfg), d.upper()] + [cell.get(k) for k in keys])
            if sigs is not None:
                nodes = tuple(sorted(D.nodes[i])) if do_topo else ()
                sigs[eid] = (cfg, nodes, tuple(cell.get(k) for k in keys))
            if i % (PROGRESS_EVERY * 5) == 0:
                BUS.progress(T("Analiza: %d / %d", "Analysis: %d / %d", i, N))
        avail = [k for k in keys if values[k]]
        exp = []
        for k in avail:
            m = REP[k]
            st = stats_of(values[k])
            edges, counts, under, over = histogram(values[k], m.bins)
            bad = sum(1 for v in values[k] if is_bad(v, m.wdir, m.thr)) if m.wdir != "none" else 0
            st.update({"key": k, "label": m.label, "desc": m.desc, "dataname": data_name(k),
                       "thr": m.thr, "dir": m.wdir, "bad": bad, "badpct": pct(bad, st["n"]),
                       "edges": edges, "counts": counts, "under": under, "over": over})
            exp.append(st)
        worst = dict((k, worst_list(values[k], vids[k], REP[k].wdir, self.topn)) for k in avail)
        comp_flat = []
        for cid in comp_order:
            for k in avail:
                vs = by_comp.get((cid, k))
                if not vs:
                    continue
                st = stats_of(vs)
                m = REP[k]
                bad = None if m.wdir == "none" else sum(1 for v in vs if is_bad(v, m.wdir, m.thr))
                st.update({"comp": comp_names[cid], "cid": cid, "key": k, "label": m.label, "bad": bad,
                           "badpct": None if bad is None else pct(bad, st["n"])})
                comp_flat.append(st)
        run = {
            "source": source, "exp": exp, "score": compute_score(exp), "worst": worst,
            "total": total, "nsel": N, "d1": d1, "d2": d2, "d3": d3, "skip": skip, "types": types,
            "vol": vol if voln else None, "voln": voln, "area": area if arean else None, "arean": arean,
            "avail": avail, "keys": keys, "rows": rows, "comp_flat": comp_flat,
            "comp_meta": [(comp_names[c], comp_n[c]) for c in comp_order],
            "sig": sigs, "sig_keys": keys, "has_nodes": do_topo, "when": now_text("%Y-%m-%d %H:%M:%S"),
            "scope": self.scope, "level": self.level, "dims": self.dims_label(),
            "nodes": nodes_all, "nnodes": max(0, nnodes),
        }
        return run

    # ---------------------------------------------------------- przebieg
    def base_path(self):
        base = os.path.splitext(self.out_file)[0] if self.out_file else ""
        return base or self.out_file

    def generate(self):
        """Raport wg ustawien (biezacy model / plik / porownanie). Zwraca opis."""
        if not any(self.fmt.values()):
            raise ValueError(T("Wybierz przynajmniej jeden format wynikowy (TXT / CSV / XLSX / HTML).",
                               "Choose at least one output format (TXT / CSV / XLSX / HTML)."))
        if not any(self.use_dim.values()):
            raise ValueError(T("Zaznacz przynajmniej jeden wymiar element\u00f3w (1D / 2D / 3D).",
                               "Tick at least one element dimension (1D / 2D / 3D)."))
        if not self.base_path():
            raise ValueError(T("Wska\u017c plik wynikowy raportu.", "Choose the report output file."))
        if self.src == "compare":
            return self.generate_compare()
        source = ""
        if self.src == "file":
            if not os.path.isfile(self.hm_file):
                raise ValueError(T("Wska\u017c istniej\u0105cy plik .hm.", "Choose an existing .hm file."))
            BUS.log(T("Otwieranie: %s", "Opening: %s", self.hm_file))
            ok, msg = HM.read_file(self.hm_file)
            if not ok:
                raise RuntimeError(T("Nie uda\u0142o si\u0119 wczyta\u0107 pliku: %s", "Could not load the file: %s", msg))
            source = self.hm_file
        run = self.analyze(source or HM.model_file())
        written, failed = self.write_all(run, self.base_path())
        self.last = run
        self.last_cmp = self.last_ncmp = None
        self.last_files = list(written)
        msg = T("Gotowe. Przeanalizowano %d element\u00f3w.\nZapisane pliki:\n  %s", "Done. %d elements analysed.\nSaved files:\n  %s",
                run["nsel"], "\n  ".join(written) or "-")
        sc = run["score"]
        if sc["has"]:
            msg += "\n\n" + T("Wska\u017anik jako\u015bci: %.1f / 100 (ocena %s)", "Quality score: %.1f / 100 (grade %s)", sc["score"], sc["grade"])
            if self.trend_prev:
                msg += " (%+.1f)" % sd1(sc["score"] - self.trend_prev["score"])
        if failed:
            msg += "\n\n" + T("NIEUDANE:", "FAILED:") + "\n  " + "\n  ".join(failed)
        return msg

    def write_all(self, run, base):
        """Zapis wszystkich wybranych formatow; (zapisane, nieudane)."""
        written, failed = [], []
        trend_path = base + ".qrtrend"
        self.trend_hist = trend_load(trend_path)
        self.trend_prev = self.trend_hist[-1] if self.trend_hist else None
        entry = None
        if run["score"]["has"]:
            badtot = sum(m["bad"] for m in run["exp"] if m["dir"] != "none")
            entry = {"ts": now_text(), "score": round(run["score"]["score"], 2), "total": run["total"], "badtot": badtot}
            self.trend_hist = self.trend_hist + [entry]
        run["trend_prev"], run["trend_hist"] = self.trend_prev, self.trend_hist
        jobs = []
        if self.fmt.get("txt"):
            jobs.append((base + ".txt", lambda p: write_text(p, report_txt(run, self))))
        if self.fmt.get("csv"):
            jobs += report_csv_jobs(run, self, base)
        if self.fmt.get("xlsx"):
            jobs.append((base + ".xlsx", lambda p: report_xlsx(run, self, p)))
        if self.fmt.get("html"):
            jobs.append((base + ".html", lambda p: write_text(p, report_html(run, self))))
        for path, fn in jobs:
            BUS.progress(T("Zapis: %s", "Writing: %s", os.path.basename(path)))
            try:
                fn(path)
                written.append(path)
            except Exception as e:
                failed.append("%s (%s)" % (path, e))
        if written and entry is not None and trend_append(trend_path, entry):
            BUS.log(T("Trend dopisany do: %s", "Trend appended to: %s", trend_path))
        return written, failed

    def generate_compare(self):
        for f, lab in ((self.ref_file, "REF"), (self.inf_file, "INF")):
            if not os.path.isfile(f):
                raise ValueError(T("Wska\u017c istniej\u0105cy plik %s (.hm).", "Choose an existing %s file (.hm).", lab))
        base = self.base_path()
        runs = {}
        for lab, f in (("REF", self.ref_file), ("INF", self.inf_file)):
            BUS.log(T("=== Analiza modelu %s: %s ===", "=== Analysing model %s: %s ===", lab, f))
            ok, msg = HM.read_file(f)
            if not ok:
                raise RuntimeError(T("Nie uda\u0142o si\u0119 wczyta\u0107 %s: %s", "Could not load %s: %s", lab, msg))
            run = self.analyze(f, sig=self.chk_ids)
            self.write_all(run, "%s_%s" % (base, lab))
            runs[lab] = run
        self.last = runs["INF"]
        BUS.progress(T("Por\u00f3wnanie zbior\u00f3w element\u00f3w A vs B\u2026", "Comparing element sets A vs B\u2026"))
        cmp = compare_sets(runs["REF"], runs["INF"], self.tol(), self.chk_ids)
        BUS.progress(T("Zestawienie w\u0119z\u0142\u00f3w A vs B\u2026", "Node comparison A vs B\u2026"))
        ncmp = compare_nodes(runs["REF"], runs["INF"], self.ntol())
        self.last_cmp, self.last_ncmp = cmp, ncmp
        for lab in ("REF", "INF"):
            runs[lab]["nodes"] = None          # wspolrzedne nie sa juz potrzebne (pamiec)
        written, failed = [], []
        suf = T("porownanie", "comparison")
        jobs = []
        if self.fmt.get("txt"):
            jobs.append(("%s_%s.txt" % (base, suf), lambda p: write_text(p, compare_txt(runs["REF"], runs["INF"], cmp, self, ncmp))))
        if self.fmt.get("html"):
            jobs.append(("%s_%s.html" % (base, suf), lambda p: write_text(p, compare_html(runs["REF"], runs["INF"], cmp, self, ncmp))))
        if self.fmt.get("xlsx"):
            jobs.append(("%s_%s.xlsx" % (base, suf), lambda p: compare_xlsx(runs["REF"], runs["INF"], cmp, self, p, ncmp)))
        for path, fn in jobs:
            try:
                fn(path)
                written.append(path)
            except Exception as e:
                failed.append("%s (%s)" % (path, e))
        self.last_files = list(written)
        sr, si = runs["REF"]["score"], runs["INF"]["score"]
        msg = T("Por\u00f3wnanie zako\u0144czone.\nREF: %s / 100    INF: %s / 100    (%s)\n\nRaport por\u00f3wnawczy:\n  %s\n\nOsobne raporty zapisane z przyrostkami _REF i _INF.",
                "Comparison finished.\nREF: %s / 100    INF: %s / 100    (%s)\n\nComparison report:\n  %s\n\nSeparate reports saved with _REF and _INF suffixes.",
                score_str(sr), score_str(si), score_delta(sr, si), "\n  ".join(written) or "-")
        if ncmp["on"]:
            msg += "\n\n" + T("W\u0119z\u0142y: A %d, B %d, wsp\u00f3lne %d, tylko A %d, tylko B %d, przesuni\u0119te %d (> %g mm), maks. %s mm",
                             "Nodes: A %d, B %d, common %d, only A %d, only B %d, moved %d (> %g mm), max %s mm",
                             ncmp["nA"], ncmp["nB"], ncmp["common"], ncmp["onlyA"], ncmp["onlyB"], ncmp["moved"], ncmp["tol"], fmt_num(ncmp["dmax"], 4))
        vd = vol_differs(runs["REF"], runs["INF"], self.tol())
        msg += "\n\n" + _identity_verdict(cmp, vd, ncmp)[0]
        if failed:
            msg += "\n\n" + T("NIEUDANE:", "FAILED:") + "\n  " + "\n  ".join(failed)
        return msg

    KEYS = ("src", "hm_file", "ref_file", "inf_file", "scope", "out_file", "fmt", "per_elem", "per_comp",
            "num_style", "level", "use_metric", "use_dim", "chk_ids", "chk_topo", "chk_vol", "cmp_tol",
            "chk_nodes", "node_tol")

    def to_dict(self):
        return dict((k, getattr(self, k)) for k in self.KEYS)

    def from_dict(self, d):
        assign_attrs(self, d, self.KEYS, {"src": ("current", "file", "compare"), "scope": ("all", "displayed"),
                                         "num_style": ("pl", "en"), "level": ("full", "basic")})
        if to_float(d.get("cmp_tol")) is not None:
            self.cmp_tol = to_float(d["cmp_tol"])
        if to_float(d.get("node_tol")) is not None:
            self.node_tol = to_float(d["node_tol"])


def score_str(sc):
    return "%.1f" % sc["score"] if sc and sc.get("has") else "-"


def score_delta(a, b):
    if not (a and b and a.get("has") and b.get("has")):
        return "-"
    return "%+.1f" % sd1(b["score"] - a["score"])


def compare_sets(run_a, run_b, tol, on=True):
    """Porownanie zbiorow elementow A (REF) i B (INF) po ID."""
    res = {"on": False, "topo": False, "common": 0, "onlyA": 0, "onlyB": 0, "same": 0, "dtype": 0,
           "dtopo": 0, "dval": 0, "difftot": 0, "keys": [], "rows": [], "shown": 0, "trunc": False}
    sa, sb = run_a.get("sig"), run_b.get("sig")
    if not on or sa is None or sb is None:
        return res
    ka, kb = run_a["sig_keys"], run_b["sig_keys"]
    keys = [k for k in ka if k in kb]
    ia = [ka.index(k) for k in keys]
    ib = [kb.index(k) for k in keys]
    topo = run_a["has_nodes"] and run_b["has_nodes"]
    res.update({"on": True, "topo": topo, "keys": keys})

    def differs(a, b):
        if a is None and b is None:
            return False
        if a is None or b is None:
            return True
        if a == b:
            return False
        m = max(abs(a), abs(b))
        return m > 0 and abs(a - b) > tol * m

    for eid in sorted(set(sa) | set(sb)):
        a, b = sa.get(eid), sb.get(eid)
        if a is not None and b is None:
            res["onlyA"] += 1
            row = {"id": eid, "st": "only_a", "reasons": ["only_a"], "cfgA": a[0], "cfgB": None,
                   "dk": [], "va": [a[2][i] for i in ia], "vb": [None] * len(keys)}
        elif a is None:
            res["onlyB"] += 1
            row = {"id": eid, "st": "only_b", "reasons": ["only_b"], "cfgA": None, "cfgB": b[0],
                   "dk": [], "va": [None] * len(keys), "vb": [b[2][i] for i in ib]}
        else:
            res["common"] += 1
            reasons = []
            if a[0] != b[0]:
                reasons.append("type")
                res["dtype"] += 1
            if topo and a[1] and b[1] and a[1] != b[1]:
                reasons.append("topo")
                res["dtopo"] += 1
            va = [a[2][i] for i in ia]
            vb = [b[2][i] for i in ib]
            dk = [k for k, x, y in zip(keys, va, vb) if differs(x, y)]
            if dk:
                reasons.append("value")
                res["dval"] += 1
            if not reasons:
                res["same"] += 1
                continue
            row = {"id": eid, "st": "diff", "reasons": reasons, "cfgA": a[0], "cfgB": b[0], "dk": dk, "va": va, "vb": vb}
        res["difftot"] += 1
        if res["shown"] < DIFFMAX:
            res["rows"].append(row)
            res["shown"] += 1
        else:
            res["trunc"] = True
    return res


def reason_text(reasons, dk=()):
    names = {"only_a": T("tylko w modelu A", "only in model A"), "only_b": T("tylko w modelu B", "only in model B"),
             "type": T("inny typ elementu", "different element type"),
             "topo": T("inna topologia (w\u0119z\u0142y)", "different topology (nodes)"),
             "value": T("inne warto\u015bci metryk", "different metric values")}
    s = " + ".join(names.get(r, r) for r in reasons)
    if dk:
        s += " (" + ", ".join(REP[k].col for k in dk) + ")"
    return s


def compare_nodes(run_a, run_b, tol):
    """Zestawienie wezlow modelu A (REF) i B (INF) po ID: wspolne, tylko A,
    tylko B, wezly bez zmiany polozenia (|d| <= tol) i PRZESUNIETE, maks. /
    srednie przesuniecie, najwieksze skladowe, lista najbardziej przesunietych."""
    res = {"on": False, "nA": run_a.get("nnodes", 0), "nB": run_b.get("nnodes", 0), "common": 0, "onlyA": 0,
           "onlyB": 0, "same": 0, "moved": 0, "dmax": 0.0, "dmax_id": None, "dmean": 0.0, "axis_max": [0.0, 0.0, 0.0],
           "rows": [], "trunc": False, "tol": tol, "only_a_ids": [], "only_b_ids": []}
    A, B = run_a.get("nodes"), run_b.get("nodes")
    if A is None or B is None:
        return res
    res["on"] = True
    res["nA"], res["nB"] = len(A), len(B)
    moved, dsum = [], 0.0
    for nid, pa in A.items():
        pb = B.get(nid)
        if pb is None:
            res["onlyA"] += 1
            continue
        res["common"] += 1
        dx, dy, dz = pb[0] - pa[0], pb[1] - pa[1], pb[2] - pa[2]
        d = math.sqrt(dx * dx + dy * dy + dz * dz)
        if d <= tol:
            res["same"] += 1
            continue
        res["moved"] += 1
        dsum += d
        for k, c in enumerate((dx, dy, dz)):
            if abs(c) > res["axis_max"][k]:
                res["axis_max"][k] = abs(c)
        if d > res["dmax"]:
            res["dmax"], res["dmax_id"] = d, nid
        moved.append((d, nid, pa, pb))
    res["onlyB"] = len(B) - res["common"]
    if res["moved"]:
        res["dmean"] = dsum / res["moved"]
    moved.sort(key=lambda x: -x[0])
    res["rows"] = [{"id": nid, "d": d, "a": pa, "b": pb} for d, nid, pa, pb in moved[:DIFFMAX]]
    res["trunc"] = len(moved) > DIFFMAX
    if res["onlyA"]:
        res["only_a_ids"] = sorted(n for n in A if n not in B)[:DIFFMAX]
    if res["onlyB"]:
        res["only_b_ids"] = sorted(n for n in B if n not in A)[:DIFFMAX]
    return res


def node_rows(ncmp):
    """Wiersze tabeli zestawienia wezlow: (etykieta, A, B, flaga 'roznica')."""
    return [(T("W\u0119z\u0142y razem", "Nodes in total"), ncmp["nA"], ncmp["nB"], False),
            (T("ID w\u0119z\u0142\u00f3w w OBU modelach", "Node IDs in BOTH models"), ncmp["common"], ncmp["common"], False),
            (T("ID w\u0119z\u0142\u00f3w tylko w modelu A", "Node IDs only in model A"), ncmp["onlyA"], "", ncmp["onlyA"] > 0),
            (T("ID w\u0119z\u0142\u00f3w tylko w modelu B", "Node IDs only in model B"), "", ncmp["onlyB"], ncmp["onlyB"] > 0),
            (T("Wsp\u00f3lne \u2013 to samo po\u0142o\u017cenie (\u2264 %g mm)", "Common \u2013 same position (\u2264 %g mm)", ncmp["tol"]), ncmp["same"], ncmp["same"], False),
            (T("Wsp\u00f3lne \u2013 PRZESUNI\u0118TE (> %g mm)", "Common \u2013 MOVED (> %g mm)", ncmp["tol"]), ncmp["moved"], ncmp["moved"], ncmp["moved"] > 0)]


def node_stat_lines(ncmp):
    L = []
    if ncmp["moved"]:
        L.append(T("Maks. przesuni\u0119cie: %s mm (w\u0119ze\u0142 %s)   \u015brednie: %s mm   maks. |dx| %s, |dy| %s, |dz| %s",
                   "Max displacement: %s mm (node %s)   mean: %s mm   max |dx| %s, |dy| %s, |dz| %s",
                   fmt_num(ncmp["dmax"], 4), ncmp["dmax_id"], fmt_num(ncmp["dmean"], 4),
                   fmt_num(ncmp["axis_max"][0], 4), fmt_num(ncmp["axis_max"][1], 4), fmt_num(ncmp["axis_max"][2], 4)))
    else:
        L.append(T("\u017baden wsp\u00f3lny w\u0119ze\u0142 nie zmieni\u0142 po\u0142o\u017cenia (tolerancja %g mm).", "No common node changed position (tolerance %g mm).", ncmp["tol"]))
    return L


def nodes_differ(ncmp):
    return bool(ncmp and ncmp.get("on") and (ncmp["moved"] or ncmp["onlyA"] or ncmp["onlyB"]))


def vol_differs(run_a, run_b, tol):
    """1 = objetosc/pole rozne, 0 = zgodne, -1 = brak danych."""
    anyv = -1
    for f in ("vol", "area"):
        a, b = run_a.get(f), run_b.get(f)
        if a is None or b is None:
            continue
        anyv = 0
        m = max(abs(a), abs(b))
        if m > 0 and abs(a - b) > tol * m:
            return 1
    return anyv


REPORT = ReportEngine()


# =================== ZAPIS RAPORTU: TXT / CSV / XLSX / HTML ============
# TXT - czytelny raport tekstowy (tez do wklejenia w maila),
# CSV - tabele (separator i przecinek dziesietny wg ustawienia PL / EN),
# XLSX - skoroszyt z formatowaniem, filtrami i wykresami histogramow,
# HTML - SAMODZIELNY, interaktywny raport (dziala offline): karty, zegar
#        wskaznika, histogramy SVG z podswietleniem binow poza progiem,
#        sortowalne tabele, tryb ciemny, trend, przycisk kopiujacy komende
#        *createmark z ID najgorszych elementow (wklejasz w okno komend HM).
LINE = "=" * 64
DASH = "-" * 64


def _dirsym(d):
    return ">" if d == "above" else "<" if d == "below" else ""


def _bar(count, mx, width=40):
    if mx <= 0:
        return ""
    n = int(round(float(count) / mx * width))
    return "#" * (1 if (count > 0 and n < 1) else n)


def report_txt(run, eng):
    full = eng.level != "basic"
    L = [LINE, "  " + T("RAPORT JAKO\u015aCI SIATKI", "MESH QUALITY REPORT") + "        %s  v%s" % (APP_TITLE, VERSION), LINE,
         "  " + T("Data:    %s", "Date:    %s", run["when"])]
    L.append("  " + (T("Plik:    %s", "File:    %s", run["source"]) if run["source"] else
                     T("\u0179r\u00f3d\u0142o:  aktualny model w sesji HyperMesh", "Source:  current model in the HyperMesh session")))
    L.append("  " + T("Zakres:  %s", "Scope:   %s", T("tylko wy\u015bwietlone", "displayed only") if run["scope"] == "displayed" else T("ca\u0142a siatka", "whole mesh")))
    L.append("")
    L.append("  " + T("Element\u00f3w razem:        %d", "Total elements:         %d", run["total"]))
    if run.get("nnodes"):
        L.append("  " + T("W\u0119z\u0142\u00f3w:                %d", "Nodes:                  %d", run["nnodes"]))
    if run["d1"]:
        L.append("    - " + T("1D (belkowe):       %d", "1D (beams):         %d", run["d1"]))
    L.append("    - " + T("2D (pow\u0142okowe):     %d", "2D (shells):        %d", run["d2"]))
    L.append("    - " + T("3D (bry\u0142owe):       %d", "3D (solids):        %d", run["d3"]))
    if run["skip"]:
        L.append("    - " + T("poza zakresem (pom.): %d", "out of scope (skip.): %d", run["skip"]))
    if run["vol"] is not None:
        L.append("  " + T("Obj\u0119to\u015b\u0107 (elementy 3D): %s mm3   (zmierzono: %d)", "Volume (3D elements):   %s mm3   (measured: %d)", vol_fmt(run["vol"]), run["voln"]))
    if run["area"] is not None:
        L.append("  " + T("Pole (elementy 2D):     %s mm2   (zmierzono: %d)", "Area (2D elements):     %s mm2   (measured: %d)", vol_fmt(run["area"]), run["arean"]))
    if full:
        L += ["", "  " + T("Typy element\u00f3w:", "Element types:")]
        for tn in sorted(run["types"]):
            L.append("    %-14s %d" % (tn, run["types"][tn]))
        L += ["", LINE, "  " + T("WSKA\u0179NIK JAKO\u015aCI SIATKI (0-100)", "MESH QUALITY SCORE (0-100)"), DASH]
        sc = run["score"]
        if sc["has"]:
            L.append("  " + T("Wynik: %.1f / 100    Ocena: %s", "Score: %.1f / 100    Grade: %s", sc["score"], sc["grade"]))
            for k, v in sc["parts"]:
                L.append("    %-24s %6.1f / 100" % (REP[k].label, v))
            L.append("  " + T("(100 = brak element\u00f3w poza progami; \u015brednia z metryk z progiem)",
                              "(100 = no elements out of limits; mean over metrics with a threshold)"))
            prev = run.get("trend_prev")
            if prev:
                L.append("  " + T("Poprzednio (%s): %.1f / 100   zmiana: %+.1f", "Previously (%s): %.1f / 100   change: %+.1f",
                                  prev.get("ts", ""), prev["score"], sd1(sc["score"] - prev["score"])))
        else:
            L.append("  " + T("(brak danych do wyliczenia wska\u017anika)", "(no data to compute the score)"))
    idx = dict((m["key"], m) for m in run["exp"])
    for k in run["keys"]:
        m = REP[k]
        L += ["", LINE, "  %s   (%s)" % (m.label, m.desc), DASH]
        if k not in idx:
            L.append("  " + T("Brak element\u00f3w, dla kt\u00f3rych ta metryka ma zastosowanie (albo brak danych).",
                              "No elements this metric applies to (or no data)."))
            continue
        x = idx[k]
        L.append("  dataname:    %s" % x["dataname"])
        L.append("  " + T("Elementy:    %d", "Elements:    %d", x["n"]))
        L.append("  Min:         %s" % pnum(x["min"]))
        L.append("  Max:         %s" % pnum(x["max"]))
        L.append("  " + T("\u015arednia:     %s", "Mean:        %s", pnum(x["mean"])))
        if full:
            L.append("  " + T("Mediana:     %s", "Median:      %s", pnum(x["median"])))
            L.append("  " + T("Odch. std:   %s", "Std. dev.:   %s", pnum(x["sd"])))
            L.append("  " + T("Percentyl  5%%: %s", "Percentile  5%%: %s", pnum(x["p5"])))
            L.append("  " + T("Percentyl 95%%: %s", "Percentile 95%%: %s", pnum(x["p95"])))
            if x["dir"] != "none":
                tag = T("   <-- do poprawy", "   <-- to fix") if x["bad"] else ""
                L.append("  " + T("Elementy %s %s:  %d  (%.2f%%)%s", "Elements %s %s:  %d  (%.2f%%)%s",
                                  _dirsym(x["dir"]), pnum(x["thr"]), x["bad"], x["badpct"], tag))
        L += ["", "  " + T("Rozk\u0142ad (histogram):", "Distribution (histogram):"),
              "  " + T("Skala (min / max danych):  %s / %s", "Scale (data min / max):  %s / %s", pnum(x["min"]), pnum(x["max"]))]
        edges, counts, under, over, n = x["edges"], x["counts"], x["under"], x["over"], x["n"]
        mx = max(counts + [under, over, 0])
        if under:
            L.append("    %-18s %9d  (%5.1f%%)  %s  (min: %s)" % ("< " + pe(edges[0]), under, pct(under, n), _bar(under, mx), pnum(x["min"])))
        for i, c in enumerate(counts):
            L.append("    %-18s %9d  (%5.1f%%)  %s" % ("%s - %s" % (pe(edges[i]), pe(edges[i + 1])), c, pct(c, n), _bar(c, mx)))
        tail = "  (max: %s)" % pnum(x["max"]) if over else ""
        L.append("    %-18s %9d  (%5.1f%%)  %s%s" % ("> " + pe(edges[-1]), over, pct(over, n), _bar(over, mx), tail))
        if full and run["worst"].get(k):
            L += ["", "  " + T("TOP %d najgorszych (id: warto\u015b\u0107):", "TOP %d worst (id: value):", len(run["worst"][k]))]
            for r, (v, eid) in enumerate(run["worst"][k], 1):
                L.append("    %2d. id=%-10s %s" % (r, eid, pnum(v)))
    if run["comp_flat"]:
        L += report_comp_txt(run)
    L.append("")
    if full:
        L += [LINE, "  " + T("UWAGI:", "NOTES:"),
              "  " + T("- Warto\u015bci pobrane z HyperMesha (atrybuty element\u00f3w) wg domy\u015blnych definicji metryk;",
                       "- Values read from HyperMesh (element attributes) with the default metric definitions;"),
              "    " + T("progi i biny histogramu: REPORT_METRICS w pliku narz\u0119dzia.",
                         "thresholds and histogram bins: REPORT_METRICS in the tool file."),
              "  " + T("- Metryki liczone tylko dla element\u00f3w, dla kt\u00f3rych maj\u0105 sens",
                       "- Metrics are computed only for elements they apply to"),
              "    " + T("(np. taper / k\u0105t \u2013 tylko 2D quad, tet collapse \u2013 tylko tetra).",
                         "(e.g. taper / angle \u2013 2D quad only, tet collapse \u2013 tetra only).")]
    L.append(LINE)
    return "\n".join(L) + "\n"


def report_comp_txt(run):
    cols = [k for k in run["avail"] if REP[k].wdir != "none"]
    L = ["", "=" * 67, " " + T("PODZIA\u0141 PER KOMPONENT \u2013 liczba element\u00f3w DO POPRAWY", "PER COMPONENT \u2013 number of elements TO FIX"), "=" * 67]
    if not cols or not run["comp_meta"]:
        return L + [" " + T("(brak danych z progiem ostrze\u017cenia)", "(no data with a warning threshold)")]
    bad = dict(((m["comp"], m["key"]), m["bad"]) for m in run["comp_flat"])
    hdr = "%-22s %7s" % (T("Komponent", "Component"), "N") + "".join(" %6s" % REP[k].abbr for k in cols)
    L += [hdr, "-" * len(hdr)]
    for nm, ne in run["comp_meta"]:
        short = nm if len(nm) <= 22 else nm[:19] + "..."
        L.append("%-22s %7d" % (short, ne) + "".join(" %6s" % (bad[(nm, k)] if (nm, k) in bad else "-") for k in cols))
    L += ["", " " + T("Pe\u0142ne statystyki per komponent: arkusz/plik \u201eKomponenty\u201d (CSV/XLSX).",
                     "Full per-component statistics: 'Components' sheet/file (CSV/XLSX).")]
    return L


def vol_fmt(v):
    if v is None:
        return "-"
    if v != 0 and (abs(v) >= 1e12 or abs(v) < 1e-4):
        return "%.6g" % v
    return "%.3f" % v


# -------------------------------------------------------------- CSV
def _csv_num(v, style):
    f = to_float(v)
    if f is None:
        return "" if v is None else "%s" % v
    s = ("%d" % f) if (isinstance(v, int) and not isinstance(v, bool)) else ("%.10g" % f)
    return s.replace(".", ",") if style == "pl" else s


def _csv_field(s, sep):
    s = "%s" % s
    if sep in s or '"' in s or "\n" in s or "\r" in s:
        return '"' + s.replace('"', '""') + '"'
    return s


def write_csv(path, rows, style):
    """rows: listy komorek; liczby formatowane wg stylu (pl/en)."""
    sep = ";" if style == "pl" else ","
    with io.open(path, "w", encoding="utf-8-sig", newline="") as fh:
        for row in rows:
            fh.write(sep.join(_csv_field(_csv_num(c, style) if isinstance(c, (int, float)) else ("" if c is None else c), sep)
                              for c in row) + "\r\n")


def summary_rows(run):
    rows = [[T("Metryka", "Metric"), "dataname", "N", "Min", "Max", T("\u015arednia", "Mean"), T("Mediana", "Median"),
             T("OdchStd", "StdDev"), "P5", "P95", T("Pr\u00f3g", "Threshold"), T("Kier", "Dir"),
             T("DoPoprawy", "ToFix"), T("ProcDoPoprawy", "PctToFix")]]
    for m in run["exp"]:
        rows.append([m["label"], m["dataname"], m["n"], m["min"], m["max"], m["mean"], m["median"], m["sd"],
                     m["p5"], m["p95"], None if m["dir"] == "none" else m["thr"], _dirsym(m["dir"]),
                     m["bad"], m["badpct"]])
    return rows


def histo_rows(run):
    rows = [[T("Metryka", "Metric"), T("Od", "From"), T("Do", "To"), T("Liczba", "Count"), T("Procent", "Percent")]]
    for m in run["exp"]:
        e, n = m["edges"], m["n"]
        rows.append([m["label"], m["min"] if m["under"] else None, e[0], m["under"], pct(m["under"], n)])
        for i, c in enumerate(m["counts"]):
            rows.append([m["label"], e[i], e[i + 1], c, pct(c, n)])
        rows.append([m["label"], e[-1], m["max"] if m["over"] else None, m["over"], pct(m["over"], n)])
    return rows


def worst_rows(run):
    rows = [[T("Metryka", "Metric"), T("Lp.", "Rank"), "ID", T("Warto\u015b\u0107", "Value")]]
    for k in run["avail"]:
        for r, (v, eid) in enumerate(run["worst"].get(k, []), 1):
            rows.append([REP[k].label, r, eid, v])
    return rows


def comp_rows(run):
    rows = [[T("Komponent", "Component"), T("KompID", "CompID"), T("Metryka", "Metric"), "N", "Min", "Max",
             T("\u015arednia", "Mean"), T("Mediana", "Median"), T("OdchStd", "StdDev"), T("DoPoprawy", "ToFix"),
             T("ProcDoPoprawy", "PctToFix")]]
    for m in run["comp_flat"]:
        rows.append([m["comp"], m["cid"], m["label"], m["n"], m["min"], m["max"], m["mean"], m["median"], m["sd"],
                     m["bad"], m["badpct"]])
    return rows


def elem_rows(run):
    rows = [["ID", T("Typ", "Type"), T("Wymiar", "Dim")] + [REP[k].col for k in run["keys"]]]
    return rows + (run["rows"] or [])


def report_csv_jobs(run, eng, base):
    st = eng.num_style
    jobs = [("%s_%s.csv" % (base, T("podsumowanie", "summary")), lambda p: write_csv(p, summary_rows(run), st)),
            ("%s_%s.csv" % (base, "histogram"), lambda p: write_csv(p, histo_rows(run), st))]
    if any(run["worst"].values()):
        jobs.append(("%s_%s.csv" % (base, T("najgorsze", "worst")), lambda p: write_csv(p, worst_rows(run), st)))
    if run["rows"] is not None:
        jobs.append(("%s_%s.csv" % (base, T("elementy", "elements")), lambda p: write_csv(p, elem_rows(run), st)))
    if run["comp_flat"]:
        jobs.append(("%s_%s.csv" % (base, T("komponenty", "components")), lambda p: write_csv(p, comp_rows(run), st)))
    return jobs


# -------------------------------------------------------------- XLSX
def need_xlsx():
    if xlsxwriter is None:
        raise RuntimeError(T("Brak biblioteki XlsxWriter w Pythonie HyperMesha.", "XlsxWriter library missing in the HyperMesh Python."))


class XlsxStyles(object):
    """Wspolne formaty skoroszytow (naglowek, liczby, kolory delt)."""

    def __init__(self, wb):
        f = wb.add_format
        self.hdr = f({"bold": True, "font_color": "#FFFFFF", "bg_color": "#1C5A96", "border": 1,
                      "border_color": "#8EA9DB", "text_wrap": True, "valign": "vcenter", "align": "center"})
        self.title = f({"bold": True, "font_size": 16, "font_color": "#FFFFFF", "bg_color": "#1F3864", "align": "center", "valign": "vcenter"})
        self.band = f({"bold": True, "font_size": 12, "font_color": "#FFFFFF", "bg_color": "#1F4E79", "align": "center", "valign": "vcenter"})
        self.sec = f({"bold": True, "font_color": "#1F4E79", "bg_color": "#DDEBF7", "indent": 1, "valign": "vcenter"})
        self.note = f({"italic": True, "font_size": 9, "font_color": "#7F7F7F"})
        self.lab = f({"bold": True})
        self.txt = f({})
        self.cen = f({"align": "center"})
        self.int = f({"num_format": "#,##0"})
        self.num = f({"num_format": "0.000"})
        self.gen = f({"num_format": "General"})
        self.pct = f({"num_format": '0.00"%"'})
        self.tot = f({"bold": True, "top": 6, "top_color": "#1F4E79", "num_format": "#,##0"})
        self.tot_lab = f({"bold": True, "top": 6, "top_color": "#1F4E79", "align": "center"})
        self.tot_pct = f({"bold": True, "top": 6, "top_color": "#1F4E79", "num_format": '0.00"%"'})
        self.good = f({"font_color": "#375623", "bg_color": "#E2EFDA", "num_format": '"+"#,##0.###;"-"#,##0.###;0'})
        self.bad = f({"font_color": "#C00000", "bg_color": "#FCE4D6", "num_format": '"+"#,##0.###;"-"#,##0.###;0'})
        self.neu = f({"num_format": '"+"#,##0.###;"-"#,##0.###;0'})
        self.goodp = f({"font_color": "#375623", "bg_color": "#E2EFDA", "num_format": '"+"0.00"%";"-"0.00"%";0.00"%"'})
        self.badp = f({"font_color": "#C00000", "bg_color": "#FCE4D6", "num_format": '"+"0.00"%";"-"0.00"%";0.00"%"'})
        self.neup = f({"num_format": '"+"0.00"%";"-"0.00"%";0.00"%"'})
        self.verd_ok = f({"bold": True, "font_size": 14, "font_color": "#375623"})
        self.verd_bad = f({"bold": True, "font_size": 14, "font_color": "#C00000"})
        self.red_int = f({"font_color": "#C00000", "num_format": "#,##0"})
        self.green_int = f({"font_color": "#375623", "num_format": "#,##0"})
        self.red_num = f({"font_color": "#C00000", "num_format": "0.000"})


def _xl_write_table(ws, rows, st, widths=None, start=0, num_cols=None):
    """Tabela: pierwszy wiersz = naglowek, liczby jako liczby."""
    for r, row in enumerate(rows):
        for c, v in enumerate(row):
            if r == 0:
                ws.write_string(start, c, "%s" % v, st.hdr)
            elif v is None:
                ws.write_blank(start + r, c, None)
            elif isinstance(v, bool):
                ws.write_string(start + r, c, "%s" % v)
            elif isinstance(v, int):
                ws.write_number(start + r, c, v, st.int)
            elif isinstance(v, float):
                ws.write_number(start + r, c, v, st.gen)
            else:
                ws.write_string(start + r, c, "%s" % v)
    if widths:
        for c, w in enumerate(widths):
            ws.set_column(c, c, w)
    if len(rows) > 1:
        ws.autofilter(start, 0, start + len(rows) - 1, len(rows[0]) - 1)
    ws.freeze_panes(start + 1, 0)


XLSX_MAX_ROWS = 1048570          # z zapasem na naglowek i wiersz uwagi


def report_xlsx(run, eng, path):
    need_xlsx()
    wb = xlsxwriter.Workbook(path, {"constant_memory": False, "strings_to_numbers": False})
    st = XlsxStyles(wb)
    ws = wb.add_worksheet(T("Podsumowanie", "Summary"))
    _xl_write_table(ws, summary_rows(run), st, [24, 14, 9, 11, 11, 11, 11, 11, 11, 11, 10, 6, 11, 13])
    ws = wb.add_worksheet(T("Histogramy", "Histograms"))
    hrows = histo_rows(run)
    _xl_write_table(ws, hrows, st, [24, 11, 11, 10, 10])
    # wykresy histogramow (kolumnowe) - jeden na metryke
    wc = wb.add_worksheet(T("Wykresy", "Charts"))
    row = 1
    r0 = 1
    for m in run["exp"]:
        nb = len(m["counts"]) + 2
        ch = wb.add_chart({"type": "column"})
        sh = T("Histogramy", "Histograms")
        ch.add_series({"name": m["label"],
                       "categories": [sh, r0, 1, r0 + nb - 1, 1],
                       "values": [sh, r0, 3, r0 + nb - 1, 3],
                       "fill": {"color": "#1F4E79"}, "gap": 60})
        ch.set_title({"name": "%s (%s)" % (m["label"], m["desc"]), "name_font": {"size": 11}})
        ch.set_legend({"none": True})
        ch.set_x_axis({"name": T("od (dolna granica binu)", "from (bin lower edge)")})
        ch.set_y_axis({"name": T("liczba element\u00f3w", "element count"), "major_gridlines": {"visible": True}})
        wc.insert_chart(row, 1, ch, {"x_scale": 1.3, "y_scale": 1.0})
        row += 17
        r0 += nb
    if any(run["worst"].values()):
        ws = wb.add_worksheet(T("Najgorsze", "Worst"))
        _xl_write_table(ws, worst_rows(run), st, [24, 7, 12, 14])
    if run["comp_flat"]:
        ws = wb.add_worksheet(T("Komponenty", "Components"))
        _xl_write_table(ws, comp_rows(run), st, [28, 9, 22, 9, 11, 11, 11, 11, 11, 11, 13])
    if run["rows"] is not None:
        ws = wb.add_worksheet(T("Elementy", "Elements"))
        rows = elem_rows(run)
        if len(rows) > XLSX_MAX_ROWS:
            # limit arkusza Excela (1 048 576 wierszy) - reszta tylko w CSV
            cut = len(rows) - XLSX_MAX_ROWS
            rows = rows[:XLSX_MAX_ROWS]
            rows.append([T("\u2026 pomini\u0119to %d wierszy (limit arkusza Excela) \u2013 pe\u0142na tabela w CSV", "\u2026 %d rows omitted (Excel sheet limit) \u2013 full table in the CSV", cut)])
        _xl_write_table(ws, rows, st, [12, 11, 8] + [13] * len(run["keys"]))
    wb.close()


# -------------------------------------------------------------- HTML
HTML_CSS = """
:root{--bg:#eef2f6;--card:#ffffff;--tx:#1b2733;--mut:#5c6c7c;--acc:#1c5a96;--ok:#5BA314;--bad:#c0392b;--warn:#e67e22;--line:#d7dfe8;--th:#f3f6fa}
body.dark{--bg:#0f141a;--card:#1a2230;--tx:#e7edf4;--mut:#93a3b4;--acc:#5b9bd5;--line:#2a3648;--th:#202a3a}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--tx);font:14px/1.45 "Segoe UI",Arial,sans-serif}
header{background:#1c5a96;color:#fff;padding:14px 22px;display:flex;flex-wrap:wrap;gap:10px;align-items:center;justify-content:space-between}
header h1{font-size:19px;margin:0}
header .meta{font-size:12px;opacity:.92;margin-top:3px}
#tb{background:#ffffff22;border:1px solid #ffffff55;color:#fff;border-radius:6px;padding:5px 12px;cursor:pointer;font-size:12px}
main{max-width:1080px;margin:0 auto;padding:18px}
.cards{display:flex;flex-wrap:wrap;gap:12px;margin-bottom:16px}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:12px 16px;min-width:130px;flex:1;box-shadow:0 1px 3px rgba(16,24,40,.06)}
.card .k{font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:var(--mut)}
.card .v{font-size:24px;font-weight:600;margin-top:2px}
.card.gaugecard{display:flex;gap:14px;align-items:center}
.gv{font:700 22px "Segoe UI",Arial;fill:var(--tx)}
.gg{font:600 13px "Segoe UI",Arial;fill:var(--mut)}
section{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:14px 18px;margin-bottom:16px;box-shadow:0 1px 3px rgba(16,24,40,.06)}
section h2{margin:0 0 4px;font-size:16px;color:var(--acc)}
section .dir{font-size:12px;color:var(--mut);font-weight:400;margin-left:8px}
.grid{display:flex;flex-wrap:wrap;gap:24px;align-items:flex-start}
.k{font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:var(--mut);margin:6px 0 2px}
table{border-collapse:collapse;font-size:13px;margin-top:6px}
th,td{border:1px solid var(--line);padding:4px 9px;text-align:right}
th{background:var(--th);cursor:pointer;user-select:none;white-space:nowrap}
td:first-child,th:first-child{text-align:left}
tr:hover td{background:var(--th)}
.hl{font:11px Consolas,monospace;fill:var(--mut)}
.hc{font:11px Consolas,monospace;fill:var(--tx)}
.hb{fill:var(--acc)}
.hb.bad{fill:var(--bad)}
.hb.part{fill:var(--warn)}
.hint{font-size:11px;color:var(--mut);margin-top:6px}
.cpy{margin-top:8px;background:var(--ok);color:#fff;border:0;border-radius:6px;padding:6px 12px;cursor:pointer;font-size:12px}
footer{color:var(--mut);font-size:12px;text-align:center;padding:10px 0 26px}
.badv{color:var(--bad);font-weight:600}
.okv{color:var(--ok);font-weight:600}
@media print{#tb,.cpy{display:none}body{background:#fff}section,.card{box-shadow:none}}
"""

HTML_JS = """
document.getElementById('tb').addEventListener('click',function(){document.body.classList.toggle('dark');});
function cellVal(tr,i){var t=tr.children[i].textContent.trim();var x=parseFloat(t.replace(',','.'));return isNaN(x)?t.toLowerCase():x;}
document.querySelectorAll('table.srt').forEach(function(tb){
  var ths=tb.querySelectorAll('th');
  ths.forEach(function(th,i){
    th.addEventListener('click',function(){
      var tbody=tb.tBodies[0];var rows=Array.prototype.slice.call(tbody.rows);
      var asc=th.getAttribute('data-a')!=='1';
      ths.forEach(function(o){o.removeAttribute('data-a');});
      if(asc){th.setAttribute('data-a','1');}
      rows.sort(function(a,b){var x=cellVal(a,i),y=cellVal(b,i);
        if(typeof x==='number'&&typeof y==='number'){return asc?x-y:y-x;}
        x=String(x);y=String(y);return asc?x.localeCompare(y):y.localeCompare(x);});
      rows.forEach(function(r){tbody.appendChild(r);});
    });
  });
});
function cpy(b){var t=b.getAttribute('data-c');var ta=document.createElement('textarea');ta.value=t;document.body.appendChild(ta);ta.select();try{document.execCommand('copy');}catch(e){}
document.body.removeChild(ta);var o=b.textContent;b.textContent=b.getAttribute('data-d');setTimeout(function(){b.textContent=o;},1200);}
"""


def h_esc(s):
    return ("%s" % s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def svg_gauge(score, grade):
    s = max(0.0, min(100.0, float(score)))
    col = "#c0392b" if s < 50 else "#e67e22" if s < 70 else "#d4a017" if s < 85 else "#5BA314"
    return ('<svg class="gauge" viewBox="0 0 120 120" width="120" height="120">'
            '<circle cx="60" cy="60" r="50" fill="none" stroke="var(--line)" stroke-width="12"/>'
            '<circle cx="60" cy="60" r="50" fill="none" stroke="%s" stroke-width="12" stroke-linecap="round" '
            'pathLength="100" stroke-dasharray="%.1f 100" transform="rotate(-90 60 60)"/>'
            '<text x="60" y="58" text-anchor="middle" class="gv">%.1f</text>'
            '<text x="60" y="78" text-anchor="middle" class="gg">%s</text></svg>') % (col, s, s, h_esc(grade))


def svg_spark(hist):
    hist = [h for h in hist if to_float(h.get("score")) is not None][-60:]
    n = len(hist)
    if n < 2:
        return ""
    W, H, pad = 240, 48, 6
    sc = [float(h["score"]) for h in hist]
    lo, hi = min(sc), max(sc)
    if hi - lo < 1.0:
        mid = (hi + lo) / 2.0
        lo, hi = mid - 0.5, mid + 0.5
    pts = ["%.1f,%.1f" % (pad + i * (W - 2 * pad) / float(n - 1), H - pad - (v - lo) * (H - 2 * pad) / (hi - lo))
           for i, v in enumerate(sc)]
    lx, ly = pts[-1].split(",")
    return ('<svg class="spark" viewBox="0 0 %d %d" width="%d" height="%d"><polyline fill="none" stroke="var(--acc)" '
            'stroke-width="2" points="%s"/><circle cx="%s" cy="%s" r="3" fill="var(--ok)"/></svg>') % (W, H, W, H, " ".join(pts), lx, ly)


def _hist_cls(wdir, thr, a, b):
    if wdir == "above":
        return "hb bad" if a >= thr else "hb part" if b > thr else "hb"
    if wdir == "below":
        return "hb bad" if b <= thr else "hb part" if a < thr else "hb"
    return "hb"


def svg_hist(m):
    e, counts, n = m["edges"], m["counts"], m["n"]
    rows = []
    if m["under"]:
        rows.append(("< " + pe(e[0]), m["under"], _hist_cls(m["dir"], m["thr"], -1e300, e[0])))
    for i, c in enumerate(counts):
        rows.append(("%s - %s" % (pe(e[i]), pe(e[i + 1])), c, _hist_cls(m["dir"], m["thr"], e[i], e[i + 1])))
    rows.append(("> " + pe(e[-1]), m["over"], _hist_cls(m["dir"], m["thr"], e[-1], 1e300)))
    mx = max([1] + [r[1] for r in rows])
    rh, lw, bw, cw = 20, 150, 240, 100
    W, H = lw + bw + cw, len(rows) * rh + 6
    out = ['<svg class="hist" viewBox="0 0 %d %d" width="%d" height="%d">' % (W, H, W, H)]
    y = 3
    for lab, c, cls in rows:
        bl = float(c) / mx * (bw - 6)
        if c > 0 and bl < 2:
            bl = 2
        out.append('<text x="%d" y="%d" text-anchor="end" class="hl">%s</text>' % (lw - 6, y + 14, h_esc(lab)))
        out.append('<rect x="%d" y="%d" width="%.1f" height="%d" class="%s" rx="2"/>' % (lw, y + 3, bl, rh - 7, cls))
        out.append('<text x="%d" y="%d" class="hc">%d (%.1f%%)</text>' % (lw + int(bl) + 8, y + 14, c, pct(c, n)))
        y += rh
    out.append("</svg>")
    return "".join(out)


def html_page(title, meta, body):
    return ('<!DOCTYPE html>\n<html lang="%s"><head><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width,initial-scale=1"><title>%s</title><style>%s</style></head>\n'
            '<body>\n<header><div><h1>%s</h1><div class="meta">%s</div></div><button id="tb">\u263d %s</button></header>\n'
            '<main>\n%s</main>\n<footer>%s</footer>\n<script>%s</script></body></html>\n') % (
        lang(), h_esc(title), HTML_CSS, h_esc(title), meta, h_esc(T("Motyw", "Theme")), body,
        h_esc(T("Wygenerowano: %s \u2022 %s v%s", "Generated: %s \u2022 %s v%s", now_text("%Y-%m-%d %H:%M:%S"), APP_TITLE, VERSION)), HTML_JS)


def report_html(run, eng):
    full = eng.level != "basic"
    src = run["source"] or T("aktualny model HyperMesh", "current HyperMesh model")
    scope = T("tylko wy\u015bwietlone", "displayed only") if run["scope"] == "displayed" else T("ca\u0142a siatka", "whole mesh")
    meta = "%s: %s \u2022 %s: %s \u2022 %s" % (h_esc(T("Plik", "File")), h_esc(src), h_esc(T("Zakres", "Scope")), h_esc(scope), run["when"])
    H = ['<div class="cards">']
    sc = run["score"]
    if full and sc["has"]:
        H.append('<div class="card gaugecard">%s<div><div class="k">%s</div><div class="v">%.1f<span style="font-size:13px;color:var(--mut)"> /100</span></div>'
                 % (svg_gauge(sc["score"], sc["grade"]), h_esc(T("Wska\u017anik jako\u015bci", "Quality score")), sc["score"]))
        prev = run.get("trend_prev")
        if prev:
            dd = sc["score"] - prev["score"]
            H.append('<div class="%s">%+.1f</div>' % ("okv" if dd >= 0 else "badv", sd1(dd)))
        H.append("</div></div>")
    hint = "2D: %d \u00b7 3D: %d" % (run["d2"], run["d3"])
    if run["d1"]:
        hint = "1D: %d \u00b7 %s" % (run["d1"], hint)
    H.append('<div class="card"><div class="k">%s</div><div class="v">%d</div><div class="hint">%s</div></div>'
             % (h_esc(T("Elementy", "Elements")), run["total"], hint))
    if run["vol"] is not None:
        H.append('<div class="card"><div class="k">%s</div><div class="v" style="font-size:18px">%s</div><div class="hint">mm\u00b3 \u00b7 N: %d</div></div>'
                 % (h_esc(T("Obj\u0119to\u015b\u0107 3D", "3D volume")), vol_fmt(run["vol"]), run["voln"]))
    if run["area"] is not None:
        H.append('<div class="card"><div class="k">%s</div><div class="v" style="font-size:18px">%s</div><div class="hint">mm\u00b2 \u00b7 N: %d</div></div>'
                 % (h_esc(T("Pole 2D", "2D area")), vol_fmt(run["area"]), run["arean"]))
    if full:
        badtot = sum(m["bad"] for m in run["exp"] if m["dir"] != "none")
        H.append('<div class="card"><div class="k">%s</div><div class="v %s">%d</div></div>'
                 % (h_esc(T("Poza progami", "Out of bounds")), "badv" if badtot else "okv", badtot))
        if len(run.get("trend_hist") or []) >= 2:
            H.append('<div class="card"><div class="k">%s</div>%s</div>' % (h_esc(T("Trend (kolejne raporty)", "Trend (successive reports)")), svg_spark(run["trend_hist"])))
    H.append("</div>\n")
    if full:
        heads = [T("Metryka", "Metric"), "N", "Min", "P5", T("Mediana", "Median"), T("\u015arednia", "Mean"), "P95", "Max",
                 T("OdchStd", "StdDev"), T("Pr\u00f3g", "Threshold"), T("DoPoprawy", "ToFix"), T("ProcDoPoprawy", "PctToFix")]
    else:
        heads = [T("Metryka", "Metric"), "N", "Min", T("\u015arednia", "Mean"), "Max"]
    H.append('<section><h2>%s</h2><table class="srt"><thead><tr>%s</tr></thead><tbody>'
             % (h_esc(T("Podsumowanie", "Summary")), "".join("<th>%s</th>" % h_esc(h) for h in heads)))
    for m in run["exp"]:
        thr = ("%s %s" % (h_esc(_dirsym(m["dir"])), pnum(m["thr"]))) if m["dir"] != "none" else "-"
        cells = [h_esc(m["label"]), "%d" % m["n"]]
        if full:
            cells += [pnum(m[f]) for f in ("min", "p5", "median", "mean", "p95", "max", "sd")]
            if m["dir"] == "none":
                cells += [thr, "-", "-"]
            else:
                cl = ' class="badv"' if m["bad"] else ""
                H.append("<tr>%s<td>%s</td><td%s>%d</td><td%s>%.2f</td></tr>" % (
                    "".join("<td>%s</td>" % c for c in cells), thr, cl, m["bad"], cl, m["badpct"]))
                continue
        else:
            cells += [pnum(m[f]) for f in ("min", "mean", "max")]
        H.append("<tr>%s</tr>" % "".join("<td>%s</td>" % c for c in cells))
    H.append('</tbody></table><div class="hint">%s</div></section>\n' % h_esc(T("Kliknij nag\u0142\u00f3wek kolumny, aby sortowa\u0107.", "Click a column header to sort.")))
    for m in run["exp"]:
        k = m["key"]
        H.append('<section><h2>%s<span class="dir">%s \u2022 dataname: %s</span></h2><div class="grid">'
                 % (h_esc(m["label"]), h_esc(m["desc"]), h_esc(m["dataname"])))
        pairs = [("N", "n"), ("Min", "min"), (T("\u015arednia", "Mean"), "mean"), ("Max", "max")]
        if full:
            pairs = [("N", "n"), ("Min", "min"), ("P5", "p5"), (T("Mediana", "Median"), "median"), (T("\u015arednia", "Mean"), "mean"),
                     ("P95", "p95"), ("Max", "max"), (T("OdchStd", "StdDev"), "sd")]
        H.append('<div><div class="k">%s</div><table class="stat"><tbody>' % h_esc(T("Statystyki", "Statistics")))
        for lab, f in pairs:
            H.append("<tr><td>%s</td><td>%s</td></tr>" % (h_esc(lab), m[f] if f == "n" else pnum(m[f])))
        if full and m["dir"] != "none":
            H.append('<tr><td>%s</td><td class="%s">%d (%.2f%%)</td></tr>' % (h_esc(T("Do poprawy", "To fix")), "badv" if m["bad"] else "okv", m["bad"], m["badpct"]))
        H.append("</tbody></table></div>")
        H.append('<div><div class="k">%s</div>%s</div>' % (h_esc(T("Rozk\u0142ad", "Distribution")), svg_hist(m)))
        w = run["worst"].get(k)
        if full and w:
            H.append('<div><div class="k">%s</div><table class="srt"><thead><tr><th>%s</th><th>ID</th><th>%s</th></tr></thead><tbody>'
                     % (h_esc(T("TOP %d najgorszych", "TOP %d worst", len(w))), h_esc(T("Lp.", "Rank")), h_esc(T("Warto\u015b\u0107", "Value"))))
            for r, (v, eid) in enumerate(w, 1):
                H.append("<tr><td>%d</td><td>%s</td><td>%s</td></tr>" % (r, eid, pnum(v)))
            cmd = "*createmark elems 1 " + " ".join("%s" % eid for _, eid in w)
            H.append('</tbody></table><button class="cpy" data-d="%s" data-c="%s" onclick="cpy(this)">%s</button></div>'
                     % (h_esc(T("Skopiowano!", "Copied!")), h_esc(cmd), h_esc(T("Kopiuj *createmark", "Copy *createmark"))))
        H.append("</div></section>\n")
    if run["comp_flat"] and run["comp_meta"]:
        cols = [k for k in run["avail"] if REP[k].wdir != "none"]
        bad = dict(((m["comp"], m["key"]), m["bad"]) for m in run["comp_flat"])
        H.append('<section><h2>%s</h2><table class="srt"><thead><tr><th>%s</th><th>N</th>%s</tr></thead><tbody>'
                 % (h_esc(T("Komponenty", "Components")), h_esc(T("Komponent", "Component")),
                    "".join("<th>%s</th>" % h_esc(REP[k].col) for k in cols)))
        for nm, ne in run["comp_meta"]:
            cells = []
            for k in cols:
                b = bad.get((nm, k))
                cells.append("<td>-</td>" if b is None else '<td%s>%d</td>' % (' class="badv"' if b else "", b))
            H.append("<tr><td>%s</td><td>%d</td>%s</tr>" % (h_esc(nm), ne, "".join(cells)))
        H.append("</tbody></table></section>\n")
    return html_page(T("Raport jako\u015bci siatki", "Mesh quality report"), meta, "".join(H))


# ===================== RAPORT POROWNAWCZY REF vs INF ==================
# Wspolny raport dwoch siatek: wskazniki jakosci i werdykt, zgodnosc
# siatek (ID wspolne / tylko A / tylko B / inny typ / inna topologia /
# inne wartosci metryk, objetosc i pole), tabela delt oraz histogramy
# REF i INF obok siebie. Skoroszyt XLSX: uklad blokowy REF | INF | DELTA
# z kolorowaniem delt (zielony = zmiana korzystna, czerwony = niekorzystna
# wzgledem progu metryki) i wykresami kolumnowymi REF / INF.
def _cmp_delta(a, b, intp=False):
    if a is None or b is None:
        return "-"
    d = b - a
    return "%+d" % int(d) if intp else "%+.4g" % d


def _bad_total(run):
    return sum(m["bad"] for m in run["exp"] if m["dir"] != "none")


def _verdict(sr, si):
    if not (sr.get("has") and si.get("has")):
        return ""
    d = si["score"] - sr["score"]
    if d > 0.05:
        return T("Werdykt: INF lepszy o %.1f pkt wska\u017anika.", "Verdict: INF better by %.1f score points.", d)
    if d < -0.05:
        return T("Werdykt: REF lepszy o %.1f pkt wska\u017anika.", "Verdict: REF better by %.1f score points.", -d)
    return T("Werdykt: wska\u017aniki praktycznie r\u00f3wne.", "Verdict: scores practically equal.")


def _identity_verdict(cmp, vd, ncmp=None):
    if not cmp["on"]:
        return T("WERDYKT: nie wyznaczono \u2013 por\u00f3wnanie ID element\u00f3w by\u0142o wy\u0142\u0105czone",
                 "VERDICT: not determined \u2013 element ID comparison was disabled"), None
    if cmp["difftot"] > 0:
        return T("WERDYKT: SIATKI R\u00d3\u017bNI\u0104 SI\u0118 \u2013 liczba r\u00f3\u017cni\u0105cych si\u0119 element\u00f3w: %d",
                 "VERDICT: THE MESHES DIFFER \u2013 number of differing elements: %d", cmp["difftot"]), False
    if nodes_differ(ncmp):
        return T("WERDYKT: ELEMENTY ZGODNE, ale R\u00d3\u017bNI\u0104 SI\u0118 W\u0118Z\u0141Y \u2013 przesuni\u0119te: %d, tylko w A: %d, tylko w B: %d",
                 "VERDICT: ELEMENTS MATCH, but the NODES DIFFER \u2013 moved: %d, only in A: %d, only in B: %d",
                 ncmp["moved"], ncmp["onlyA"], ncmp["onlyB"]), False
    if vd == 1:
        return T("WERDYKT: ID, typy i warto\u015bci metryk zgodne, ale R\u00d3\u017bNI SI\u0118 OBJ\u0118TO\u015a\u0106/POLE modelu",
                 "VERDICT: IDs, types and metric values match, but the model VOLUME/AREA DIFFERS"), False
    return T("WERDYKT: SIATKI S\u0104 IDENTYCZNE w zakresie sprawdzanych kryteri\u00f3w",
             "VERDICT: THE MESHES ARE IDENTICAL within the checked criteria"), True


def _id_count_rows(ra, ri, eng, cmp):
    rows = [(T("Elementy w wybranym zakresie (razem)", "Elements in the selected scope (total)"), ra["total"], ri["total"]),
            (T("Elementy obj\u0119te sprawdzeniem", "Elements checked"), ra["nsel"], ri["nsel"])]
    for d, k in (("1d", "d1"), ("2d", "d2"), ("3d", "d3")):
        if eng.use_dim.get(d):
            rows.append(("   " + T("w tym %s", "of which %s", d.upper()), ra[k], ri[k]))
    return rows


def _cmp_counts(cmp):
    rows = [(T("ID wyst\u0119puj\u0105ce w OBU modelach", "IDs present in BOTH models"), cmp["common"], False),
            (T("ID tylko w modelu A", "IDs only in model A"), cmp["onlyA"], True),
            (T("ID tylko w modelu B", "IDs only in model B"), cmp["onlyB"], True),
            (T("Wsp\u00f3lne ID \u2013 w pe\u0142ni identyczne", "Common IDs \u2013 fully identical"), cmp["same"], False),
            (T("Wsp\u00f3lne ID \u2013 inny typ elementu", "Common IDs \u2013 different element type"), cmp["dtype"], True)]
    if cmp["topo"]:
        rows.append((T("Wsp\u00f3lne ID \u2013 inna topologia (w\u0119z\u0142y)", "Common IDs \u2013 different topology (nodes)"), cmp["dtopo"], True))
    rows.append((T("Wsp\u00f3lne ID \u2013 inne warto\u015bci metryk", "Common IDs \u2013 different metric values"), cmp["dval"], True))
    return rows


def compare_txt(ra, ri, cmp, eng, ncmp=None):
    ia = dict((m["key"], m) for m in ra["exp"])
    ii = dict((m["key"], m) for m in ri["exp"])
    L = [LINE, "  " + T("RAPORT POR\u00d3WNAWCZY SIATEK   REF vs INF", "MESH COMPARISON REPORT   REF vs INF") + "   %s v%s" % (APP_TITLE, VERSION), LINE,
         "  " + T("Data:    %s", "Date:    %s", now_text("%Y-%m-%d %H:%M:%S")),
         "  REF: %s" % ra["source"], "  INF: %s" % ri["source"], "",
         "  " + T("Elementy:  REF %d    INF %d    (%s)", "Elements:  REF %d    INF %d    (%s)", ra["total"], ri["total"], _cmp_delta(ra["total"], ri["total"], True))]
    sr, si = ra["score"], ri["score"]
    if sr["has"] and si["has"]:
        L.append("  " + T("Wska\u017anik jako\u015bci:  REF %.1f / 100    INF %.1f / 100    (%+.1f)", "Quality score:  REF %.1f / 100    INF %.1f / 100    (%+.1f)",
                          sr["score"], si["score"], sd1(si["score"] - sr["score"])))
        L.append("  " + _verdict(sr, si))
    L.append("  " + T("Poza progami \u0142\u0105cznie:  REF %d    INF %d    (%s)", "Out of bounds in total:  REF %d    INF %d    (%s)",
                      _bad_total(ra), _bad_total(ri), _cmp_delta(_bad_total(ra), _bad_total(ri), True)))
    # zgodnosc siatek
    f = "  %-40s %14s %14s %14s"
    L += ["", LINE, "  " + T("ZGODNO\u015a\u0106 SIATEK   (MODEL A = REF, MODEL B = INF)", "MESH IDENTITY   (MODEL A = REF, MODEL B = INF)"), DASH,
          f % ("", "MODEL A (REF)", "MODEL B (INF)", T("r\u00f3\u017cnica", "delta"))]
    for lab, a, b in _id_count_rows(ra, ri, eng, cmp):
        L.append(f % (lab, a, b, _cmp_delta(a, b, True)))
    vd = vol_differs(ra, ri, eng.tol())
    L.append("")
    if vd >= 0:
        for lab, k in ((T("Obj\u0119to\u015b\u0107 element\u00f3w 3D [mm\u00b3]", "3D element volume [mm\u00b3]"), "vol"), (T("Pole element\u00f3w 2D [mm\u00b2]", "2D element area [mm\u00b2]"), "area")):
            a, b = ra[k], ri[k]
            d = None if (a is None or b is None) else b - a
            L.append(f % (lab, vol_fmt(a), vol_fmt(b), "-" if d is None else ("+" if d > 0 else "") + vol_fmt(d)))
    else:
        L.append("  " + T("Uwaga: obj\u0119to\u015bci/pola nie zmierzono.", "Note: volume/area not measured."))
    L.append("")
    if cmp["on"]:
        for lab, n, _ in _cmp_counts(cmp):
            L.append("  %-40s %14s" % (lab, n))
        L.append("  %-40s %14s" % (T("ELEMENTY R\u00d3\u017bNI\u0104CE SI\u0118 \u2013 \u0141\u0104CZNIE", "DIFFERING ELEMENTS \u2013 TOTAL"), cmp["difftot"]))
    else:
        L.append("  " + T("Uwaga: por\u00f3wnanie ID element\u00f3w by\u0142o wy\u0142\u0105czone.", "Note: element ID comparison was disabled."))
    if ncmp and ncmp["on"]:
        L += ["", LINE, "  " + T("ZESTAWIENIE W\u0118Z\u0141\u00d3W   (MODEL A = REF, MODEL B = INF)   tolerancja %g mm", "NODE COMPARISON   (MODEL A = REF, MODEL B = INF)   tolerance %g mm", ncmp["tol"]), DASH,
              f % ("", "MODEL A (REF)", "MODEL B (INF)", "")]
        for lab, a, b, flag in node_rows(ncmp):
            L.append(f % (lab, a, b, "<--" if flag else ""))
        L.append("")
        for ln in node_stat_lines(ncmp):
            L.append("  " + ln)
        top = ncmp["rows"][:20]
        if top:
            L += ["", "  " + T("Najbardziej przesuni\u0119te w\u0119z\u0142y (pierwsze %d z %d):", "Most displaced nodes (first %d of %d):", len(top), ncmp["moved"])]
            for r in top:
                L.append("    ID %-9s A (%s, %s, %s)  ->  B (%s, %s, %s)   |d| = %s mm" % (
                    r["id"], pnum(r["a"][0]), pnum(r["a"][1]), pnum(r["a"][2]), pnum(r["b"][0]), pnum(r["b"][1]), pnum(r["b"][2]), fmt_num(r["d"], 4)))
        if ncmp["only_a_ids"]:
            L.append("  " + T("ID tylko w A (pierwsze %d): %s", "IDs only in A (first %d): %s", min(20, len(ncmp["only_a_ids"])), " ".join("%d" % n for n in ncmp["only_a_ids"][:20])))
        if ncmp["only_b_ids"]:
            L.append("  " + T("ID tylko w B (pierwsze %d): %s", "IDs only in B (first %d): %s", min(20, len(ncmp["only_b_ids"])), " ".join("%d" % n for n in ncmp["only_b_ids"][:20])))
    elif ncmp is not None:
        L += ["", "  " + T("Uwaga: zestawienie w\u0119z\u0142\u00f3w by\u0142o wy\u0142\u0105czone.", "Note: the node comparison was disabled.")]
    L += ["", "  " + _identity_verdict(cmp, vd, ncmp)[0]]
    rows = cmp["rows"][:30]
    if rows:
        L += ["", "  " + T("R\u00f3\u017cni\u0105ce si\u0119 elementy (pierwsze %d z %d):", "Differing elements (first %d of %d):", len(rows), cmp["difftot"])]
        for r in rows:
            L.append("    ID %-10s %-18s %s" % (r["id"], "%s / %s" % (elem_type_name(r["cfgA"]), elem_type_name(r["cfgB"])), reason_text(r["reasons"], r["dk"])))
        if cmp["difftot"] > len(rows):
            L.append("  " + T("... pe\u0142na lista w arkuszu \u201eR\u00f3\u017cnice element\u00f3w\u201d skoroszytu XLSX.", "... full list in the 'Element differences' XLSX sheet."))
    for k in REP_ORDER:
        mr, mi = ia.get(k), ii.get(k)
        if not mr and not mi:
            continue
        L += ["", LINE, "  %s   (%s)" % (REP[k].label, REP[k].desc), DASH, "  %-12s %-16s %-16s %s" % ("", "REF", "INF", T("r\u00f3\u017cnica", "delta"))]
        for lab, fld, intp in (("N", "n", True), ("Min", "min", False), (T("\u015arednia", "Mean"), "mean", False), ("Max", "max", False)):
            a = mr[fld] if mr else None
            b = mi[fld] if mi else None
            L.append("  %-12s %-16s %-16s %s" % (lab, "-" if a is None else pnum(a), "-" if b is None else pnum(b), _cmp_delta(a, b, intp)))
        base = mr or mi
        if base["dir"] != "none":
            a = "%d (%.1f%%)" % (mr["bad"], mr["badpct"]) if mr else "-"
            b = "%d (%.1f%%)" % (mi["bad"], mi["badpct"]) if mi else "-"
            L.append("  %-12s %-16s %-16s %s" % (T("DoPoprawy", "ToFix"), a, b, _cmp_delta(mr["bad"] if mr else None, mi["bad"] if mi else None, True)))
        if not mr:
            L.append("  REF: " + T("(brak w tym modelu)", "(missing in this model)"))
        if not mi:
            L.append("  INF: " + T("(brak w tym modelu)", "(missing in this model)"))
    L += ["", LINE]
    return "\n".join(L) + "\n"


def compare_html(ra, ri, cmp, eng, ncmp=None):
    ia = dict((m["key"], m) for m in ra["exp"])
    ii = dict((m["key"], m) for m in ri["exp"])
    sr, si = ra["score"], ri["score"]
    H = ['<div class="cards">']
    for lab, s in (("REF", sr), ("INF", si)):
        if s["has"]:
            extra = ""
            if lab == "INF" and sr["has"]:
                dd = si["score"] - sr["score"]
                extra = '<div class="%s">%+.1f</div>' % ("okv" if dd >= 0 else "badv", sd1(dd))
            H.append('<div class="card gaugecard">%s<div><div class="k">%s \u2022 %s</div><div class="v">%.1f</div>%s</div></div>'
                     % (svg_gauge(s["score"], s["grade"]), lab, h_esc(T("Wska\u017anik jako\u015bci", "Quality score")), s["score"], extra))
    H.append('<div class="card"><div class="k">%s</div><div class="v">%d \u2192 %d</div><div class="hint">REF \u2192 INF</div></div>'
             % (h_esc(T("Elementy", "Elements")), ra["total"], ri["total"]))
    br, bi = _bad_total(ra), _bad_total(ri)
    H.append('<div class="card"><div class="k">%s</div><div class="v">%d \u2192 %d</div><div class="%s">%+d</div></div>'
             % (h_esc(T("Poza progami", "Out of bounds")), br, bi, "okv" if bi - br <= 0 else "badv", bi - br))
    H.append("</div>\n")
    v = _verdict(sr, si)
    if v:
        H.append("<section><h2>%s</h2></section>\n" % h_esc(v))
    vd = vol_differs(ra, ri, eng.tol())
    vtxt, vok = _identity_verdict(cmp, vd, ncmp)
    H.append('<section><h2>%s</h2><p class="%s" style="font-size:16px;font-weight:700">%s</p>'
             % (h_esc(T("Zgodno\u015b\u0107 siatek (A = REF, B = INF)", "Mesh identity (A = REF, B = INF)")),
                "hint" if vok is None else ("okv" if vok else "badv"), h_esc(vtxt)))
    H.append('<table class="srt"><thead><tr><th>%s</th><th>MODEL A (REF)</th><th>MODEL B (INF)</th><th>\u0394</th></tr></thead><tbody>'
             % h_esc(T("Kryterium", "Criterion")))
    for lab, a, b in _id_count_rows(ra, ri, eng, cmp):
        H.append("<tr><td>%s</td><td>%s</td><td>%s</td><td>%s</td></tr>" % (h_esc(lab), a, b, _cmp_delta(a, b, True)))
    if vd >= 0:
        for lab, k in ((T("Obj\u0119to\u015b\u0107 element\u00f3w 3D [mm\u00b3]", "3D element volume [mm\u00b3]"), "vol"), (T("Pole element\u00f3w 2D [mm\u00b2]", "2D element area [mm\u00b2]"), "area")):
            a, b = ra[k], ri[k]
            d = None if (a is None or b is None) else b - a
            H.append("<tr><td>%s</td><td>%s</td><td>%s</td><td>%s</td></tr>" % (h_esc(lab), vol_fmt(a), vol_fmt(b), "-" if d is None else vol_fmt(d)))
    if cmp["on"]:
        for lab, n, _ in _cmp_counts(cmp):
            H.append('<tr><td>%s</td><td colspan="3">%d</td></tr>' % (h_esc(lab), n))
        H.append('<tr><td><b>%s</b></td><td colspan="3"><b>%d</b></td></tr>' % (h_esc(T("ELEMENTY R\u00d3\u017bNI\u0104CE SI\u0118 \u2013 \u0141\u0104CZNIE", "DIFFERING ELEMENTS \u2013 TOTAL")), cmp["difftot"]))
    H.append("</tbody></table>")
    rows = cmp["rows"][:500] if cmp["on"] else []
    if rows:
        H.append('<p class="hint">%s</p><table class="srt"><thead><tr><th>ID</th><th>%s</th><th>%s</th><th>%s</th><th>%s</th></tr></thead><tbody>'
                 % (h_esc(T("R\u00f3\u017cni\u0105ce si\u0119 elementy (pierwsze %d z %d):", "Differing elements (first %d of %d):", len(rows), cmp["difftot"])),
                    h_esc(T("Status", "Status")), h_esc(T("Typ w A", "Type in A")), h_esc(T("Typ w B", "Type in B")), h_esc(T("Co si\u0119 r\u00f3\u017cni", "What differs"))))
        stt = {"only_a": T("tylko w modelu A", "only in model A"), "only_b": T("tylko w modelu B", "only in model B")}
        for r in rows:
            H.append('<tr><td>%s</td><td class="badv">%s</td><td>%s</td><td>%s</td><td>%s</td></tr>'
                     % (r["id"], h_esc(stt.get(r["st"], T("r\u00f3\u017cni si\u0119", "differs"))), elem_type_name(r["cfgA"]), elem_type_name(r["cfgB"]),
                        h_esc(reason_text(r["reasons"], r["dk"]))))
        H.append("</tbody></table>")
    H.append("</section>\n")
    if ncmp and ncmp["on"]:
        H.append('<section><h2>%s<span class="dir">%s</span></h2>' % (
            h_esc(T("Zestawienie w\u0119z\u0142\u00f3w (A = REF, B = INF)", "Node comparison (A = REF, B = INF)")),
            h_esc(T("tolerancja po\u0142o\u017cenia %g mm", "position tolerance %g mm", ncmp["tol"]))))
        H.append('<table class="srt"><thead><tr><th>%s</th><th>MODEL A (REF)</th><th>MODEL B (INF)</th></tr></thead><tbody>' % h_esc(T("Kryterium", "Criterion")))
        for lab, a, b, flag in node_rows(ncmp):
            cl = ' class="badv"' if flag else ""
            H.append("<tr><td>%s</td><td%s>%s</td><td%s>%s</td></tr>" % (h_esc(lab), cl, a, cl, b))
        H.append("</tbody></table>")
        for ln in node_stat_lines(ncmp):
            H.append('<p class="hint" style="font-size:13px;color:var(--tx)">%s</p>' % h_esc(ln))
        top = ncmp["rows"][:200]
        if top:
            H.append('<p class="hint">%s</p><table class="srt"><thead><tr><th>ID</th><th>xA</th><th>yA</th><th>zA</th><th>xB</th><th>yB</th><th>zB</th><th>|d| [mm]</th></tr></thead><tbody>'
                     % h_esc(T("Najbardziej przesuni\u0119te w\u0119z\u0142y (pierwsze %d z %d):", "Most displaced nodes (first %d of %d):", len(top), ncmp["moved"])))
            for r in top:
                H.append("<tr><td>%d</td>%s%s<td class=\"badv\">%s</td></tr>" % (
                    r["id"], "".join("<td>%s</td>" % pnum(v) for v in r["a"]), "".join("<td>%s</td>" % pnum(v) for v in r["b"]), fmt_num(r["d"], 4)))
            H.append("</tbody></table>")
            cmd = "*createmark nodes 1 " + " ".join("%d" % r["id"] for r in top[:TOPN * 5])
            H.append('<button class="cpy" data-d="%s" data-c="%s" onclick="cpy(this)">%s</button>'
                     % (h_esc(T("Skopiowano!", "Copied!")), h_esc(cmd), h_esc(T("Kopiuj *createmark nodes", "Copy *createmark nodes"))))
        H.append("</section>\n")
    heads = "".join("<th>%s REF</th><th>%s INF</th><th>\u0394</th>" % (h_esc(t), h_esc(t))
                    for t in ("N", T("\u015arednia", "Mean"), "Max", T("DoPoprawy", "ToFix")))
    H.append('<section><h2>%s</h2><table class="srt"><thead><tr><th>%s</th>%s</tr></thead><tbody>'
             % (h_esc(T("Podsumowanie", "Summary")), h_esc(T("Metryka", "Metric")), heads))
    for k in REP_ORDER:
        mr, mi = ia.get(k), ii.get(k)
        if not mr and not mi:
            continue
        cells = []
        for fld, intp in (("n", True), ("mean", False), ("max", False), ("bad", True)):
            if fld == "bad" and (mr or mi)["dir"] == "none":
                cells.append("<td>-</td><td>-</td><td>-</td>")
                continue
            a = mr[fld] if mr else None
            b = mi[fld] if mi else None
            fa = "-" if a is None else ("%d" % a if intp else pnum(a))
            fb = "-" if b is None else ("%d" % b if intp else pnum(b))
            cls = ""
            if fld == "bad" and a is not None and b is not None:
                cls = ' class="%s"' % ("okv" if b - a <= 0 else "badv")
            cells.append("<td>%s</td><td>%s</td><td%s>%s</td>" % (fa, fb, cls, _cmp_delta(a, b, intp)))
        H.append("<tr><td>%s</td>%s</tr>" % (h_esc(REP[k].label), "".join(cells)))
    H.append('</tbody></table><div class="hint">%s</div></section>\n' % h_esc(T("Kliknij nag\u0142\u00f3wek kolumny, aby sortowa\u0107.", "Click a column header to sort.")))
    for k in REP_ORDER:
        mr, mi = ia.get(k), ii.get(k)
        if not mr and not mi:
            continue
        H.append('<section><h2>%s<span class="dir">%s</span></h2><div class="grid">' % (h_esc(REP[k].label), h_esc(REP[k].desc)))
        for lab, m in (("REF", mr), ("INF", mi)):
            H.append('<div><div class="k">%s</div>%s</div>' % (lab, svg_hist(m) if m else '<div class="hint">%s</div>' % h_esc(T("(brak w tym modelu)", "(missing in this model)"))))
        H.append("</div></section>\n")
    meta = "REF: %s \u2022 INF: %s \u2022 %s" % (h_esc(ra["source"]), h_esc(ri["source"]), now_text("%Y-%m-%d %H:%M:%S"))
    return html_page(T("Por\u00f3wnanie siatek REF vs INF", "Mesh comparison REF vs INF"), meta, "".join(H))


def _bin_side(lo, hi, wdir, thr):
    """1 = bin po zlej stronie progu, 0 = po dobrej, -1 = brak progu."""
    if wdir == "none":
        return -1
    mid = hi if lo is None else lo if hi is None else (lo + hi) / 2.0
    return int(mid > thr) if wdir == "above" else int(mid < thr)


def compare_xlsx(ra, ri, cmp, eng, path, ncmp=None):
    need_xlsx()
    wb = xlsxwriter.Workbook(path)
    st = XlsxStyles(wb)
    ia = dict((m["key"], m) for m in ra["exp"])
    ii = dict((m["key"], m) for m in ri["exp"])
    sh_id = T("Zgodno\u015b\u0107 siatek", "Mesh identity")
    sh_cmp = T("Por\u00f3wnanie", "Comparison")
    sh_ch = T("Wykresy", "Charts")
    sh_diff = T("R\u00f3\u017cnice element\u00f3w", "Element differences")

    def delta_fmt(d, side, kind="n"):
        if d is None:
            return st.neu
        if side < 0 or d == 0:
            return st.neup if kind == "p" else st.neu
        worse = (side == 1 and d > 0) or (side == 0 and d < 0)
        if kind == "p":
            return st.badp if worse else st.goodp
        return st.bad if worse else st.good

    # --- arkusz 1: zgodnosc siatek ---
    ws = wb.add_worksheet(sh_id)
    ws.set_column(0, 0, 46)
    ws.set_column(1, 4, 19)
    ws.merge_range(0, 0, 0, 4, T("SPRAWDZENIE ZGODNO\u015aCI SIATEK \u2013 MODEL A (REF) vs MODEL B (INF)",
                                 "MESH IDENTITY CHECK \u2013 MODEL A (REF) vs MODEL B (INF)"), st.title)
    ws.set_row(0, 30)
    ws.merge_range(1, 0, 1, 4, T("Sprawdzane wymiary: %s   |   tolerancja wzgl\u0119dna por\u00f3wnania warto\u015bci: %g",
                                 "Checked dimensions: %s   |   relative tolerance of value comparison: %g", eng.dims_label(), eng.tol()), st.note)
    vd = vol_differs(ra, ri, eng.tol())
    vtxt, vok = _identity_verdict(cmp, vd, ncmp)
    ws.merge_range(3, 0, 3, 4, vtxt, st.note if vok is None else (st.verd_ok if vok else st.verd_bad))
    ws.set_row(3, 24)
    r = 5
    ws.merge_range(r, 0, r, 4, T("LICZBY ELEMENT\u00d3W", "ELEMENT COUNTS"), st.sec)
    r += 1
    for c, h in enumerate((T("Kryterium", "Criterion"), "MODEL A (REF)", "MODEL B (INF)", T("R\u00f3\u017cnica (B - A)", "Delta (B - A)"), T("Zmiana wzgl\u0119dna", "Relative change"))):
        ws.write(r, c, h, st.hdr)
    r += 1
    for lab, a, b in _id_count_rows(ra, ri, eng, cmp):
        ws.write(r, 0, lab, st.lab)
        ws.write_number(r, 1, a, st.int)
        ws.write_number(r, 2, b, st.int)
        ws.write_number(r, 3, b - a, st.good if b == a else st.bad)
        if a:
            ws.write_number(r, 4, 100.0 * (b - a) / a, st.goodp if b == a else st.badp)
        r += 1
    r += 1
    ws.merge_range(r, 0, r, 4, T("POR\u00d3WNANIE ID ELEMENT\u00d3W (A vs B)", "ELEMENT ID COMPARISON (A vs B)"), st.sec)
    r += 1
    if cmp["on"]:
        for lab, n, flag in _cmp_counts(cmp):
            ws.write(r, 0, lab, st.lab)
            ws.write_number(r, 1, n, (st.red_int if n else st.green_int) if flag else st.int)
            r += 1
        ws.write(r, 0, T("ELEMENTY R\u00d3\u017bNI\u0104CE SI\u0118 \u2013 \u0141\u0104CZNIE", "DIFFERING ELEMENTS \u2013 TOTAL"), st.tot_lab)
        ws.write_number(r, 1, cmp["difftot"], st.tot)
        r += 1
    else:
        ws.merge_range(r, 0, r, 4, T("Uwaga: por\u00f3wnanie ID element\u00f3w by\u0142o wy\u0142\u0105czone.", "Note: element ID comparison was disabled."), st.note)
        r += 1
    r += 1
    ws.merge_range(r, 0, r, 4, T("OBJ\u0118TO\u015a\u0106 I POLE MODELU", "MODEL VOLUME AND AREA"), st.sec)
    r += 1
    if vd >= 0:
        for lab, k in ((T("Obj\u0119to\u015b\u0107 element\u00f3w 3D [mm\u00b3]", "3D element volume [mm\u00b3]"), "vol"), (T("Pole element\u00f3w 2D [mm\u00b2]", "2D element area [mm\u00b2]"), "area")):
            a, b = ra[k], ri[k]
            ws.write(r, 0, lab, st.lab)
            if a is not None:
                ws.write_number(r, 1, a, st.num)
            if b is not None:
                ws.write_number(r, 2, b, st.num)
            if a is not None and b is not None:
                ws.write_number(r, 3, b - a, st.good if abs(b - a) <= eng.tol() * max(abs(a), abs(b), 1e-30) else st.bad)
            r += 1
        ws.merge_range(r, 0, r, 4, T("Jednostki: warto\u015bci w jednostkach modelu (model w mm -> mm\u00b3 i mm\u00b2).",
                                     "Units: values in model units (model in mm -> mm\u00b3 and mm\u00b2)."), st.note)
    else:
        ws.merge_range(r, 0, r, 4, T("Uwaga: obj\u0119to\u015bci/pola nie zmierzono.", "Note: volume/area not measured."), st.note)
        r += 1
    r += 1
    ws.merge_range(r, 0, r, 4, T("ZESTAWIENIE W\u0118Z\u0141\u00d3W (A vs B)", "NODE COMPARISON (A vs B)"), st.sec)
    r += 1
    if ncmp and ncmp["on"]:
        for c, h in enumerate((T("Kryterium", "Criterion"), "MODEL A (REF)", "MODEL B (INF)")):
            ws.write(r, c, h, st.hdr)
        r += 1
        for lab, a, b, flag in node_rows(ncmp):
            ws.write(r, 0, lab, st.lab)
            for c, v in ((1, a), (2, b)):
                if v != "":
                    ws.write_number(r, c, v, st.red_int if flag else st.int)
            r += 1
        for ln in node_stat_lines(ncmp):
            ws.merge_range(r, 0, r, 4, ln, st.note)
            r += 1
    else:
        ws.merge_range(r, 0, r, 4, T("Uwaga: zestawienie w\u0119z\u0142\u00f3w by\u0142o wy\u0142\u0105czone.", "Note: the node comparison was disabled."), st.note)
        r += 1
    ws.freeze_panes(5, 0)

    # --- arkusz 2: porownanie blokowe REF | INF | DELTA ---
    ws = wb.add_worksheet(sh_cmp)
    for c, w in enumerate((24, 14, 13, 2, 24, 14, 13, 2, 14, 14, 16)):
        ws.set_column(c, c, w)
    ws.merge_range(0, 0, 0, 10, T("POR\u00d3WNANIE JAKO\u015aCI SIATKI (MESH QUALITY)", "MESH QUALITY COMPARISON"), st.title)
    ws.set_row(0, 30)
    ws.merge_range(1, 0, 1, 10, T("\u0179r\u00f3d\u0142o danych: modele .hm analizowane w HyperMesh (%s v%s)", "Data source: .hm models analysed in HyperMesh (%s v%s)", APP_TITLE, VERSION), st.note)
    ws.merge_range(3, 0, 3, 2, "REFERENCE MODEL (REF)", st.band)
    ws.merge_range(3, 4, 3, 6, "INFLUENCE MODEL (INF)", st.band)
    ws.merge_range(3, 8, 3, 10, "DELTA ANALYSIS (INF - REF)", st.band)
    ws.set_row(3, 22)
    ws.write(4, 0, T("Plik:", "File:"), st.lab)
    ws.write(4, 1, os.path.basename(ra["source"] or "-"))
    ws.write(4, 4, T("Plik:", "File:"), st.lab)
    ws.write(4, 5, os.path.basename(ri["source"] or "-"))
    ws.write(4, 8, T("Data:", "Date:"), st.lab)
    ws.write(4, 9, now_text())
    ws.write(5, 0, T("Elementy:", "Elements:"), st.lab)
    ws.write_number(5, 1, ra["total"], st.int)
    ws.write(5, 4, T("Elementy:", "Elements:"), st.lab)
    ws.write_number(5, 5, ri["total"], st.int)
    ws.write_number(5, 8, ri["total"] - ra["total"], st.neu)
    if ra["total"]:
        ws.write_number(5, 10, 100.0 * (ri["total"] - ra["total"]) / ra["total"], st.neup)
    sr, si = ra["score"], ri["score"]
    ws.write(6, 0, T("Wska\u017anik:", "Score:"), st.lab)
    ws.write(6, 4, T("Wska\u017anik:", "Score:"), st.lab)
    if sr["has"]:
        ws.write_number(6, 1, sr["score"], st.num)
        ws.write(6, 2, "(%s)" % sr["grade"], st.cen)
    if si["has"]:
        ws.write_number(6, 5, si["score"], st.num)
        ws.write(6, 6, "(%s)" % si["grade"], st.cen)
    if sr["has"] and si["has"]:
        d = si["score"] - sr["score"]
        ws.write_number(6, 8, d, st.neu if abs(d) < 0.05 else (st.good if d > 0 else st.bad))
        ws.merge_range(7, 0, 7, 10, _verdict(sr, si), st.sec)
    r = 9
    ws.merge_range(r, 0, r, 10, T("SYNTEZA \u2013 \u015brednie i warto\u015bci skrajne wg metryki (REF / INF / r\u00f3\u017cnica)",
                                  "OVERVIEW \u2013 means and extremes per metric (REF / INF / delta)"), st.sec)
    r += 1
    for c, h in ((0, T("Metryka", "Metric")), (1, T("\u015arednia", "Mean")), (2, "Max"), (5, T("\u015arednia", "Mean")), (6, "Max"),
                 (8, T("R\u00f3\u017cnica \u015bredniej", "Mean delta")), (9, T("R\u00f3\u017cnica max", "Max delta")), (10, T("Zmiana wzgl.", "Rel. change"))):
        ws.write(r, c, h, st.hdr)
    r += 1
    freeze = r
    for k in REP_ORDER:
        mr, mi = ia.get(k), ii.get(k)
        if not mr and not mi:
            continue
        side = {"above": 1, "below": 0}.get((mr or mi)["dir"], -1)
        ws.write(r, 0, REP[k].label, st.lab)
        if mr:
            ws.write_number(r, 1, mr["mean"], st.num)
            ws.write_number(r, 2, mr["max"], st.num)
        if mi:
            ws.write_number(r, 5, mi["mean"], st.num)
            ws.write_number(r, 6, mi["max"], st.num)
        if mr and mi:
            dm, dx = mi["mean"] - mr["mean"], mi["max"] - mr["max"]
            ws.write_number(r, 8, dm, delta_fmt(dm, side))
            ws.write_number(r, 9, dx, delta_fmt(dx, side))
            if mr["mean"]:
                ws.write_number(r, 10, 100.0 * dm / mr["mean"], delta_fmt(dm, side, "p"))
        r += 1
    r += 1
    chart_info = []
    for k in REP_ORDER:
        mr, mi = ia.get(k), ii.get(k)
        if not mr and not mi:
            continue
        base = mr or mi
        lbl = REP[k].label
        ws.merge_range(r, 0, r, 2, lbl, st.sec)
        ws.merge_range(r, 4, r, 6, lbl, st.sec)
        ws.merge_range(r, 8, r, 10, "DELTA VALUES", st.sec)
        r += 1
        for c, h in ((0, T("Zakres", "Range")), (1, T("Elementy", "Elements")), (2, T("Udzia\u0142 %", "Share %")),
                     (4, T("Zakres", "Range")), (5, T("Elementy", "Elements")), (6, T("Udzia\u0142 %", "Share %")),
                     (8, T("R\u00f3\u017cnica", "Delta")), (9, T("R\u00f3\u017cnica %", "Delta %")), (10, T("Zmiana wzgl.", "Rel. change"))):
            ws.write(r, c, h, st.hdr)
        r += 1
        e = base["edges"]
        specs = []
        ua, ub = (mr["under"] if mr else 0), (mi["under"] if mi else 0)
        oa, ob = (mr["over"] if mr else 0), (mi["over"] if mi else 0)
        if ua or ub:
            specs.append(("< %s" % pnum(e[0]), ua, ub, _bin_side(None, e[0], base["dir"], base["thr"])))
        for i in range(len(e) - 1):
            a = mr["counts"][i] if mr and i < len(mr["counts"]) else 0
            b = mi["counts"][i] if mi and i < len(mi["counts"]) else 0
            specs.append(("[%s - %s]" % (pnum(e[i]), pnum(e[i + 1])), a, b, _bin_side(e[i], e[i + 1], base["dir"], base["thr"])))
        if oa or ob:
            specs.append(("> %s" % pnum(e[-1]), oa, ob, _bin_side(e[-1], None, base["dir"], base["thr"])))
        na, nb_ = (mr["n"] if mr else 0), (mi["n"] if mi else 0)
        first = r
        for lab, a, b, side in specs:
            pa, pb = pct(a, na), pct(b, nb_)
            ws.write(r, 0, lab, st.cen)
            ws.write_number(r, 1, a, st.int)
            ws.write_number(r, 2, pa, st.pct)
            ws.write(r, 4, lab, st.cen)
            ws.write_number(r, 5, b, st.int)
            ws.write_number(r, 6, pb, st.pct)
            ws.write_number(r, 8, b - a, delta_fmt(b - a, side))
            ws.write_number(r, 9, pb - pa, delta_fmt(pb - pa, side, "p"))
            if a:
                ws.write_number(r, 10, 100.0 * (b - a) / a, delta_fmt(b - a, side, "p"))
            r += 1
        chart_info.append((lbl, first, r - 1))
        ws.write(r, 0, "Total", st.tot_lab)
        ws.write_number(r, 1, na, st.tot)
        ws.write_number(r, 2, 100.0 if na else 0.0, st.tot_pct)
        ws.write(r, 4, "Total", st.tot_lab)
        ws.write_number(r, 5, nb_, st.tot)
        ws.write_number(r, 6, 100.0 if nb_ else 0.0, st.tot_pct)
        ws.write_number(r, 8, nb_ - na, st.tot)
        r += 2
    ws.merge_range(r, 0, r, 10, T("Kolory delty: zielony = zmiana korzystna, czerwony = niekorzystna (wg kierunku metryki).",
                                  "Delta colors: green = favourable change, red = unfavourable (per metric direction)."), st.note)
    ws.freeze_panes(4, 0)

    # --- arkusz 3: wykresy ---
    wc = wb.add_worksheet(sh_ch)
    wc.set_column(0, 0, 26)
    wc.set_column(1, 2, 12)
    wc.write(0, 0, T("Dane do wykresu zbiorczego: indeks jako\u015bci INF wzgl\u0119dem REF (REF = 100)",
                     "Summary chart data: INF quality index relative to REF (REF = 100)"), st.lab)
    for c, h in enumerate((T("Metryka", "Metric"), "REF", "INF")):
        wc.write(1, c, h, st.hdr)
    cr = 2
    for k in REP_ORDER:
        mr, mi = ia.get(k), ii.get(k)
        if not (mr and mi) or not mr["mean"] or not mi["mean"] or mr["dir"] == "none":
            continue
        ix = 100.0 * mr["mean"] / mi["mean"] if mr["dir"] == "above" else 100.0 * mi["mean"] / mr["mean"]
        wc.write(cr, 0, REP[k].label)
        wc.write_number(cr, 1, 100.0, st.num)
        wc.write_number(cr, 2, ix, st.num)
        cr += 1
    if cr > 2:
        ch = wb.add_chart({"type": "column"})
        ch.add_series({"name": "REF", "categories": [sh_ch, 2, 0, cr - 1, 0], "values": [sh_ch, 2, 1, cr - 1, 1], "fill": {"color": "#1F4E79"}})
        ch.add_series({"name": "INF", "categories": [sh_ch, 2, 0, cr - 1, 0], "values": [sh_ch, 2, 2, cr - 1, 2], "fill": {"color": "#5BA314"}})
        ch.set_title({"name": T("Indeks jako\u015bci: INF wzgl\u0119dem REF (REF = 100, wy\u017cej = lepiej)", "Quality index: INF vs REF (REF = 100, higher = better)"), "name_font": {"size": 11}})
        ch.set_legend({"position": "bottom"})
        wc.insert_chart(0, 4, ch, {"x_scale": 1.4, "y_scale": 1.1})
    row = 22
    for lbl, r1, r2 in chart_info:
        ch = wb.add_chart({"type": "column"})
        ch.add_series({"name": "REF", "categories": [sh_cmp, r1, 0, r2, 0], "values": [sh_cmp, r1, 1, r2, 1], "fill": {"color": "#1F4E79"}, "gap": 60})
        ch.add_series({"name": "INF", "categories": [sh_cmp, r1, 0, r2, 0], "values": [sh_cmp, r1, 5, r2, 5], "fill": {"color": "#5BA314"}})
        ch.set_title({"name": "%s   (REF / INF)" % lbl, "name_font": {"size": 11}})
        ch.set_legend({"position": "bottom"})
        wc.insert_chart(row, 0, ch, {"x_scale": 1.6, "y_scale": 1.1})
        row += 18

    # --- arkusz 4: roznice elementow ---
    if cmp["on"]:
        ws = wb.add_worksheet(sh_diff)
        keys = cmp["keys"]
        nc = 6 + 3 * len(keys)
        ws.merge_range(0, 0, 0, nc - 1, T("ELEMENTY, KT\u00d3RE SI\u0118 R\u00d3\u017bNI\u0104 MI\u0118DZY MODELEM A (REF) I B (INF)",
                                          "ELEMENTS THAT DIFFER BETWEEN MODEL A (REF) AND B (INF)"), st.title)
        ws.set_row(0, 30)
        tot = cmp["difftot"]
        if tot == 0:
            note = T("Nie znaleziono \u017cadnych r\u00f3\u017cnic \u2013 siatki s\u0105 identyczne w zakresie sprawdzanych kryteri\u00f3w.",
                     "No differences found \u2013 the meshes are identical within the checked criteria.")
        elif cmp["trunc"]:
            note = T("Uwaga: znaleziono %d r\u00f3\u017cni\u0105cych si\u0119 element\u00f3w, pokazano pierwszych %d (limit arkusza).",
                     "Note: %d differing elements found, the first %d are shown (sheet limit).", tot, cmp["shown"])
        else:
            note = T("Znaleziono %d r\u00f3\u017cni\u0105cych si\u0119 element\u00f3w \u2013 lista kompletna.", "%d differing elements found \u2013 complete list.", tot)
        ws.merge_range(1, 0, 1, nc - 1, note, st.note)
        hdr = ["ID", T("Status", "Status"), T("Co si\u0119 r\u00f3\u017cni", "What differs"), T("Typ w A", "Type in A"),
               T("Typ w B", "Type in B"), T("Wymiar", "Dim")]
        for k in keys:
            hdr += ["%s - A" % REP[k].col, "%s - B" % REP[k].col, "%s - \u0394" % REP[k].col]
        for c, h in enumerate(hdr):
            ws.write(3, c, h, st.hdr)
        ws.set_row(3, 30)
        stt = {"only_a": T("tylko w modelu A", "only in model A"), "only_b": T("tylko w modelu B", "only in model B")}
        r = 4
        for row_ in cmp["rows"]:
            ws.write_number(r, 0, row_["id"], st.int)
            ws.write(r, 1, stt.get(row_["st"], T("r\u00f3\u017cni si\u0119", "differs")), st.red_int)
            ws.write(r, 2, reason_text(row_["reasons"], row_["dk"]))
            ws.write(r, 3, elem_type_name(row_["cfgA"]), st.cen)
            ws.write(r, 4, elem_type_name(row_["cfgB"]), st.cen)
            ws.write(r, 5, elem_dim(row_["cfgA"] if row_["cfgA"] is not None else row_["cfgB"]).upper(), st.cen)
            c = 6
            for k, a, b in zip(keys, row_["va"], row_["vb"]):
                if a is not None:
                    ws.write_number(r, c, a, st.num)
                if b is not None:
                    ws.write_number(r, c + 1, b, st.num)
                if a is not None and b is not None:
                    ws.write_number(r, c + 2, b - a, st.red_num if k in row_["dk"] else st.num)
                c += 3
            r += 1
        for c, w in enumerate([12, 20, 42, 12, 12, 10] + [14] * (3 * len(keys))):
            ws.set_column(c, c, w)
        if r > 4:
            ws.autofilter(3, 0, r - 1, nc - 1)
        ws.freeze_panes(4, 0)

    # --- arkusz 5: wezly ---
    if ncmp and ncmp["on"]:
        ws = wb.add_worksheet(T("W\u0119z\u0142y", "Nodes"))
        ws.merge_range(0, 0, 0, 12, T("W\u0118Z\u0141Y PRZESUNI\u0118TE MI\u0118DZY MODELEM A (REF) I B (INF)  \u2013  tolerancja %g mm",
                                      "NODES MOVED BETWEEN MODEL A (REF) AND B (INF)  \u2013  tolerance %g mm", ncmp["tol"]), st.title)
        ws.set_row(0, 30)
        if ncmp["moved"] == 0:
            note = T("\u017baden wsp\u00f3lny w\u0119ze\u0142 nie zmieni\u0142 po\u0142o\u017cenia.", "No common node changed position.")
        elif ncmp["trunc"]:
            note = T("Przesuni\u0119tych w\u0119z\u0142\u00f3w: %d \u2013 pokazano %d najbardziej przesuni\u0119tych (limit arkusza).",
                     "Moved nodes: %d \u2013 the %d most displaced are shown (sheet limit).", ncmp["moved"], len(ncmp["rows"]))
        else:
            note = T("Przesuni\u0119tych w\u0119z\u0142\u00f3w: %d \u2013 lista kompletna (malej\u0105co wg |d|).", "Moved nodes: %d \u2013 complete list (descending |d|).", ncmp["moved"])
        ws.merge_range(1, 0, 1, 12, note, st.note)
        hdr = ["ID", "xA", "yA", "zA", "xB", "yB", "zB", "dx", "dy", "dz", "|d| [mm]", T("ID tylko w A", "ID only in A"), T("ID tylko w B", "ID only in B")]
        for c, h in enumerate(hdr):
            ws.write(3, c, h, st.hdr)
        r = 4
        for row_ in ncmp["rows"]:
            ws.write_number(r, 0, row_["id"], st.int)
            for c, v in enumerate(list(row_["a"]) + list(row_["b"])):
                ws.write_number(r, 1 + c, v, st.gen)
            for c in range(3):
                ws.write_number(r, 7 + c, row_["b"][c] - row_["a"][c], st.gen)
            ws.write_number(r, 10, row_["d"], st.red_num)
            r += 1
        for i, n in enumerate(ncmp["only_a_ids"]):
            ws.write_number(4 + i, 11, n, st.int)
        for i, n in enumerate(ncmp["only_b_ids"]):
            ws.write_number(4 + i, 12, n, st.int)
        for c, w in enumerate([11, 12, 12, 12, 12, 12, 12, 11, 11, 11, 12, 14, 14]):
            ws.set_column(c, c, w)
        last = max(r, 4 + len(ncmp["only_a_ids"]), 4 + len(ncmp["only_b_ids"]))
        if last > 4:
            ws.autofilter(3, 0, last - 1, 12)
        ws.freeze_panes(4, 0)
    wb.close()


# ==================== OBRAZY: MINIATURY, TLO, KONWERSJE ===============
# HyperMesh 2024 zapisuje zrzut okna graficznego ZAWSZE na bialym tle
# (niezaleznie od tla sceny na ekranie - sprawdzone dla *jpegfilenamed
# i hw.CaptureImageTool). Wybrany kolor tla nakladamy wiec na gotowy obraz:
# piksele bialawe (tlo) -> kolor tla, a przy ciemnym tle piksele czarnawe
# (opisy, osie, legenda HM) -> biale, zeby zostaly czytelne. Kolory siatki
# i konturow zostaja nietkniete. Obrobka numpy (gdy jest) - ulamek sekundy
# nawet dla duzego zrzutu.
BG_MODES = [("none", "bez zmian", "no change", ""), ("black", "czarne", "black", "#000000"),
            ("graphite", "grafit", "graphite", "#1f1f1f"), ("navy", "granat", "navy", "#10243b"),
            ("white", "bia\u0142e", "white", "#ffffff")]


def bg_hex(mode):
    for key, _, _, hx in BG_MODES:
        if key == mode:
            return hx
    return ""


def bg_label(mode):
    for key, pl, en, _ in BG_MODES:
        if key == mode:
            return en if lang() == "en" else pl
    return mode


def need_pil():
    if Image is None:
        raise RuntimeError(T("Brak biblioteki Pillow w Pythonie HyperMesha.", "Pillow library missing in the HyperMesh Python."))


def image_size(path):
    try:
        need_pil()
        with Image.open(path) as im:
            return im.size
    except Exception:
        return (0, 0)


def convert_image(src, dst, quality=95):
    """Konwersja formatu wg rozszerzenia dst (png / jpg / bmp / tif)."""
    need_pil()
    ext = os.path.splitext(dst)[1].lower()
    with Image.open(src) as im:
        im = im.convert("RGB")
        if ext in (".jpg", ".jpeg"):
            im.save(dst, "JPEG", quality=quality, subsampling=0)
        elif ext == ".bmp":
            im.save(dst, "BMP")
        elif ext in (".tif", ".tiff"):
            im.save(dst, "TIFF")
        else:
            im.save(dst, "PNG")
    return dst


def thumb_png(path, w=THUMB_W, h=THUMB_H):
    """Miniatura pliku obrazu jako bajty PNG (albo None)."""
    try:
        need_pil()
        with Image.open(path) as im:
            im = im.convert("RGB")
            im.thumbnail((w, h), Image.LANCZOS)
            buf = io.BytesIO()
            im.save(buf, "PNG")
            return buf.getvalue()
    except Exception:
        return None


def png_from_bytes(data, fmt_hint=""):
    """Dowolny obraz (np. PPM ze starej sesji Tcl) -> bajty PNG miniatury."""
    try:
        need_pil()
        with Image.open(io.BytesIO(data)) as im:
            im = im.convert("RGB")
            im.thumbnail((THUMB_W, THUMB_H), Image.LANCZOS)
            buf = io.BytesIO()
            im.save(buf, "PNG")
            return buf.getvalue()
    except Exception:
        return None


def corners(path):
    try:
        need_pil()
        with Image.open(path) as im:
            im = im.convert("RGB")
            w, h = im.size
            if w < 10 or h < 10:
                return []
            return [im.getpixel(p) for p in ((4, 4), (w - 5, 4), (4, h - 5), (w - 5, h - 5))]
    except Exception:
        return []


def corners_match(cs, hexcol, tol=26):
    t = hex_rgb(hexcol)
    ok = sum(1 for c in cs if all(abs(c[i] - t[i]) <= tol for i in range(3)))
    return bool(cs) and ok >= (3 if len(cs) >= 4 else len(cs))


def recolor_background(path, hexcol, out=None):
    """Nak\u0142ada kolor tla na zrzut (tlo biale -> hexcol). Zwraca sciezke."""
    need_pil()
    out = out or path
    target = hex_rgb(hexcol)
    dark = is_dark(target)
    with Image.open(path) as im:
        im = im.convert("RGB")
        if np is not None:
            a = np.asarray(im).astype(np.float32)
            lo = a.min(axis=2)
            hi = a.max(axis=2)
            # piksele szare i jasne (tlo + wygladzone krawedzie siatki na tle):
            # mieszanie z kolorem tla proporcjonalnie do jasnosci - bez jasnej
            # obwodki wokol modelu na ciemnym tle. Kolory siatki (nasycone) bez zmian.
            gray = (hi - lo) <= 30
            w = np.clip((lo - 200.0) / 50.0, 0.0, 1.0) * gray
            if dark:
                black = (hi <= 47) & gray
                a[black] = 255.0
            t = np.array(target, dtype=np.float32)
            a = a * (1.0 - w[:, :, None]) + t * w[:, :, None]
            im2 = Image.fromarray(np.clip(a + 0.5, 0, 255).astype(np.uint8), "RGB")
        else:
            px = im.load()
            w, h = im.size
            for y in range(h):
                for x in range(w):
                    r, g, b = px[x, y]
                    if r >= 224 and g >= 224 and b >= 224:
                        px[x, y] = target
                    elif dark and r <= 47 and g <= 47 and b <= 47:
                        px[x, y] = (255, 255, 255)
            im2 = im
    ext = os.path.splitext(out)[1].lower()
    if ext in (".jpg", ".jpeg"):
        im2.save(out, "JPEG", quality=95, subsampling=0)
    elif ext == ".bmp":
        im2.save(out, "BMP")
    else:
        im2.save(out, "PNG")
    return out


def file_b64(path):
    with open(path, "rb") as fh:
        return base64.b64encode(fh.read()).decode("ascii")


def mime_of(path):
    return {".jpg": "jpeg", ".jpeg": "jpeg", ".png": "png", ".gif": "gif", ".bmp": "bmp"}.get(
        os.path.splitext(path)[1].lower(), "png")


# ======================== WIDOKI I ZRZUTY EKRANU =======================
# Zapamietujesz dowolna liczbe widokow (kamera: orientacja + zoom + pan).
# Przy KAZDYM widoku powstaje miniatura oraz PELNA KLATKA - dokladnie to,
# co widac na ekranie (WYSIWYG). "Generuj zrzuty" eksportuje zaznaczone
# widoki: domyslnie zapisane klatki (pozniejsze zmiany wygladu modelu nie
# psuja wczesniej zapamietanych widokow), a widoki bez klatki (dodane
# z tekstu, wczytane z sesji) sa renderowane na nowo: uklad + kamera.
# Zapis widoku: "VS: <20 liczb>" (macierz 4x4 + zakres okna, jak *viewset)
# - zgodny z makrem Tcl, wiec stare sesje HTML wczytuja sie bez zmian.
# UKLAD WYSWIETLANIA: HyperMesh zapisuje kazda akcje do pliku command*.tcl
# w folderze roboczym. Nagrywanie zapamietuje nowe linie tego pliku miedzy
# startem nagrywania a "Zapamietaj widok" i odtwarza je (po kolei, kazda
# raz) przed zrzutem w innym modelu - zrzuty wygladaja jak w modelu wzorcowym.
VIEW_PRESETS = ["REF_Aspect_Ratio", "REF_Jacobian", "REF_Skewness",
                "INF_Aspect_Ratio", "INF_Jacobian", "INF_Skewness"]
_NOISE = ("viewset", "jpegfilenamed", "pngfilenamed", "bmpfilenamed", "tifffilenamed", "setbackground",
          "sethardcopy", "hardcopy", "screencapture", "quitapplication", "hwf::setbackgroundcolor")


def view_text(nums):
    return "VS: " + " ".join("%.6f" % x for x in nums)


def parse_view(s):
    """Liczby widoku z linii (VS: / MTX: / >= 16 liczb) albo None."""
    s = (s or "").strip()
    if not s or s.startswith("#"):
        return None
    up = s.upper()
    for tag in ("VS:", "MTX:"):
        p = up.find(tag)
        if p >= 0:
            s = s[p + len(tag):].split("|")[0]
            break
    nums = [to_float(x) for x in re.findall(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", s)]
    nums = [x for x in nums if x is not None]
    return nums if len(nums) >= 16 else None


class ViewStore(object):
    def __init__(self):
        self.views = []           # rekordy: id name view thumb disp shot sel
        self.next_id = 1
        self.shot_dir = ""
        self.prefix = "shot"
        self.fmt = "png"          # png | jpg | bmp
        self.model_dir = False    # zapisuj do folderu otwartego modelu
        self.subfolder = False    # podfolder o nazwie prefiksu / pierwszej nazwy
        self.bg = "none"          # tlo zrzutu (BG_MODES)
        self.report = True        # raport HTML (sesja) przy eksporcie
        self.use_stored = True    # WYSIWYG: eksport zapisanych klatek
        self.presets = dict((p, False) for p in VIEW_PRESETS)
        self.rec_on = False
        self.rec_file = ""
        self.rec_ofs = 0
        self.disp_applied = 0

    # ---------------------------------------------------------- rekordy
    def rec(self, vid):
        for r in self.views:
            if r["id"] == vid:
                return r
        return None

    def index(self, vid):
        for i, r in enumerate(self.views):
            if r["id"] == vid:
                return i
        return -1

    def selected(self):
        return [r for r in self.views if r["sel"]]

    def new_record(self, nums, name=None, thumb=None, disp="", shot=""):
        vid = self.next_id
        self.next_id += 1
        r = {"id": vid, "name": name or T("Widok %d", "View %d", vid), "view": list(nums), "thumb": thumb,
             "disp": disp or "", "shot": shot or "", "sel": True}
        self.views.append(r)
        return r

    @staticmethod
    def shot_ok(r):
        p = r.get("shot") or ""
        return bool(p) and os.path.isfile(p) and os.path.getsize(p) > 0

    @staticmethod
    def drop_shot(r):
        p = r.get("shot") or ""
        if p and os.path.isfile(p):
            try:
                os.remove(p)
            except OSError:
                pass
        r["shot"] = ""

    # ---------------------------------------------------------- zrzut
    def capture_frame(self):
        """Pelna klatka okna graficznego (z wybranym tlem) -> sciezka albo ""."""
        HM.redraw()
        time.sleep(CAPTURE_DELAY_MS / 1000.0)
        p = tmp_file("png", "frame")
        if not HM.capture(p):
            return ""
        hx = bg_hex(self.bg)
        if hx:
            try:
                recolor_background(p, hx)
            except Exception:
                pass
        return p

    def remember(self):
        nums = HM.get_view()
        if len(nums) < 16:
            raise RuntimeError(T("Nie uda\u0142o si\u0119 pobra\u0107 widoku (czy okno graficzne jest aktywne?).",
                                 "Could not read the view (is the graphics window active?)."))
        disp = self.capture_disp_tail() if self.rec_on else ""
        shot = self.capture_frame()
        r = self.new_record(nums, thumb=thumb_png(shot) if shot else None, disp=disp, shot=shot)
        if disp:
            self.disp_applied = sum(1 for x in self.views if x["disp"])
        return r

    def apply_view(self, nums):
        ok = HM.set_view(nums)
        HM.redraw()
        HM.set_view(nums)          # ponownie - przerysowanie potrafi cofnac widok
        return ok

    def test(self, vid):
        i = self.index(vid)
        if i < 0:
            return ""
        ran, fails, behind = self.apply_disp_up_to(i)
        ok = self.apply_view(self.views[i]["view"])
        HM.redraw()
        msg = T("Zastosowano widok \u2013 sprawd\u017a ekran.", "View applied \u2013 check the screen.") if ok else \
            T("Nie uda\u0142o si\u0119 zastosowa\u0107 widoku.", "Could not apply the view.")
        if ran:
            msg += T(" Odtworzono uk\u0142ad (%d).", " Layout replayed (%d).", ran)
        if fails:
            msg += T(" Uwaga: %d komend uk\u0142adu nie wykona\u0142o si\u0119.", " Warning: %d layout commands failed.", fails)
        if behind:
            msg += T(" Uwaga: stan uk\u0142adu sesji jest ju\u017c dalej ni\u017c ten widok (wczytaj model ponownie).",
                     " Warning: the session layout is already past this view (reload the model).")
        return msg

    def refresh_thumb(self, vid):
        r = self.rec(vid)
        if r is None:
            return False
        cur = HM.get_view()
        self.apply_view(r["view"])
        shot = self.capture_frame()
        if len(cur) >= 16:
            self.apply_view(cur)
        if not shot:
            return False
        self.drop_shot(r)
        r["shot"] = shot
        r["thumb"] = thumb_png(shot)
        return True

    # ---------------------------------------------------------- edycja
    def add_from_text(self, text):
        n = 0
        for line in (text or "").splitlines():
            nums = parse_view(line)
            if nums:
                self.new_record(nums)
                n += 1
        return n

    def edit(self, vid, name, vtext, disp):
        r = self.rec(vid)
        nums = parse_view(vtext)
        if r is None or nums is None:
            return False
        if nums != r["view"]:
            self.drop_shot(r)     # zapisana klatka nie pasuje juz do danych widoku
        r["name"] = (name or "").strip() or T("Widok %d", "View %d", vid)
        r["view"] = nums
        r["disp"] = (disp or "").strip()
        return True

    def delete(self, vid):
        r = self.rec(vid)
        if r is not None:
            self.drop_shot(r)
            self.views.remove(r)

    def clear(self):
        for r in self.views:
            self.drop_shot(r)
        self.views = []
        self.disp_applied = 0

    # ---------------------------------------------------------- uklad (command.tcl)
    @staticmethod
    def find_command_file():
        base = os.getcwd()
        cands = []
        try:
            for f in os.listdir(base):
                fl = f.lower()
                if fl.startswith("command") and (fl.endswith(".tcl") or fl.endswith(".cmf")):
                    p = os.path.join(base, f)
                    cands.append((os.path.getmtime(p), p))
        except OSError:
            pass
        return max(cands)[1] if cands else ""

    def toggle_record(self, path=""):
        if self.rec_on:
            self.rec_on = False
            return True
        f = path or (self.rec_file if os.path.isfile(self.rec_file) else "") or self.find_command_file()
        if not f:
            return False
        self.rec_file = f
        self.rec_ofs = os.path.getsize(f)
        self.rec_on = True
        return True

    def capture_disp_tail(self):
        f = self.rec_file
        if not f or not os.path.isfile(f):
            return ""
        sz = os.path.getsize(f)
        if sz <= self.rec_ofs:
            return ""
        with open(f, "rb") as fh:
            fh.seek(self.rec_ofs)
            data = fh.read().decode("utf-8", "replace")
        self.rec_ofs = sz
        keep = [ln for ln in data.splitlines() if ln.strip() and not any(n in ln for n in _NOISE)]
        return "\n".join(keep)

    def apply_disp_up_to(self, idx):
        """Odtwarza przyrostowe uklady do widoku idx. (wykonane, bledy, zaDaleko)."""
        target = sum(1 for r in self.views[:idx + 1] if r["disp"])
        if target == 0:
            return 0, 0, False
        if self.disp_applied > target:
            return 0, 0, True
        ran = fails = ci = 0
        for r in self.views[:idx + 1]:
            if not r["disp"]:
                continue
            ci += 1
            if ci <= self.disp_applied:
                continue
            ok, bad = HM.eval_commands(r["disp"])
            fails += bad
            ran += 1
            self.disp_applied = ci
        HM.redraw()
        return ran, fails, False

    # ---------------------------------------------------------- eksport
    def target_dir(self):
        d = (self.shot_dir or "").strip()
        note = ""
        if self.model_dir:
            md = HM.model_dir()
            if md:
                d, note = md, T(" (folder modelu)", " (model folder)")
            else:
                note = T(" (nie wykryto folderu modelu \u2013 u\u017cyto wskazanego)", " (model folder not detected \u2013 using the chosen one)")
        return d, note

    def preset_names(self):
        return [p for p in VIEW_PRESETS if self.presets.get(p)]

    def generate(self):
        """Pliki zrzutow zaznaczonych widokow (+ raport HTML). Zwraca opis."""
        recs = self.selected()
        if not recs:
            raise ValueError(T("Zaznacz przynajmniej jeden widok.", "Select at least one view."))
        d, dnote = self.target_dir()
        if not d:
            raise ValueError(T("Wska\u017c folder zrzut\u00f3w albo zaznacz \u201efolder modelu\u201d.", "Choose a screenshot folder or tick \u201cmodel folder\u201d."))
        pfx = clean_file_name(self.prefix or "shot")
        presets = self.preset_names()
        if self.subfolder:
            d = os.path.join(d, clean_file_name(presets[0] if presets else pfx))
        os.makedirs(d, exist_ok=True)
        ext = self.fmt if self.fmt in ("png", "jpg", "bmp") else "png"
        fb = "shot" if presets else pfx
        start = next_index(d, fb)
        names, k = [], 0
        for i in range(len(recs)):
            if i < len(presets):
                names.append(presets[i])
            else:
                names.append("%s_%03d" % (fb, start + k))
                k += 1
        hx = bg_hex(self.bg)
        out, fails, stored, live, disp_fail, behind = [], 0, 0, 0, 0, False
        cur = None
        for i, r in enumerate(recs):
            BUS.progress(T("Zrzut %d / %d\u2026", "Shot %d / %d\u2026", i + 1, len(recs)))
            path = os.path.join(d, "%s.%s" % (names[i], ext))
            if self.use_stored and self.shot_ok(r):
                try:
                    convert_image(r["shot"], path)
                    out.append(path)
                    stored += 1
                    continue
                except Exception:
                    pass
            if cur is None:
                cur = HM.get_view()
            _, bad, beh = self.apply_disp_up_to(self.index(r["id"]))
            disp_fail += bad
            behind = behind or beh
            self.apply_view(r["view"])
            time.sleep(CAPTURE_DELAY_MS / 1000.0)
            tmp = tmp_file("png", "live")
            if HM.capture(tmp):
                if hx:
                    recolor_background(tmp, hx)
                convert_image(tmp, path)
                try:
                    os.remove(tmp)
                except OSError:
                    pass
                out.append(path)
                live += 1
            else:
                fails += 1
                out.append("")
        if cur and len(cur) >= 16 and live:
            self.apply_view(cur)
        rep = ""
        if self.report:
            try:
                rows = [dict(r, file=os.path.basename(p), path=p) for r, p in zip(recs, out)]
                hp = os.path.join(d, "%s_widoki.html" % pfx)
                write_text(hp, views_html(rows))
                rep = T(" + raport %s", " + report %s", os.path.basename(hp))
            except Exception as e:
                rep = T(" (raport nie zapisany: %s)", " (report not saved: %s)", e)
        msg = T("Gotowe. Zapisano %d zrzut\u00f3w w: %s%s", "Done. Saved %d screenshots in: %s%s", len([p for p in out if p]), d, dnote)
        if stored and live:
            msg += T(" (zapisane klatki: %d, renderowane: %d)", " (stored frames: %d, re-rendered: %d)", stored, live)
        elif stored:
            msg += T(" (z zapisanych klatek \u2013 WYSIWYG)", " (from stored frames \u2013 WYSIWYG)")
        if fails:
            msg += T(" UWAGA: %d zrzut\u00f3w nie powiod\u0142o si\u0119.", " WARNING: %d shots failed.", fails)
        if disp_fail:
            msg += T(" (uk\u0142ad: %d komend nie wykona\u0142o si\u0119)", " (layout: %d commands failed)", disp_fail)
        if behind:
            msg += T(" (uk\u0142ad: stan sesji jest dalej ni\u017c cz\u0119\u015b\u0107 widok\u00f3w \u2013 wczytaj model ponownie)",
                     " (layout: the session is past some views \u2013 reload the model)")
        if presets and len(presets) != len(recs):
            msg += T(" (nazw: %d, widok\u00f3w: %d)", " (names: %d, views: %d)", len(presets), len(recs))
        return msg + rep

    def save_report(self, folder):
        recs = self.selected()
        if not recs:
            raise ValueError(T("Zaznacz przynajmniej jeden widok.", "Select at least one view."))
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, "%s_widoki.html" % clean_file_name(self.prefix or "shot"))
        rows = [dict(r, file="", path=r.get("shot") or "") for r in recs]
        write_text(path, views_html(rows))
        return path

    def load_session(self, path):
        """Wczytuje sesje z pliku HTML (blok HMSESSION - takze z makra Tcl)."""
        with io.open(path, "r", encoding="utf-8", errors="replace") as fh:
            data = fh.read()
        s, e = data.find("<!--HMSESSION"), data.find("HMSESSION-->")
        if s < 0 or e < s:
            raise ValueError(T("To nie jest plik sesji (brak bloku HMSESSION).", "Not a session file (no HMSESSION block)."))
        loaded = []
        for ln in data[s:e].splitlines():
            if not ln.startswith("R\t"):
                continue
            parts = ln.split("\t")
            if len(parts) < 5:
                continue
            nums = parse_view(_ses_unesc(parts[2]))
            if not nums:
                continue
            thumb = None
            if parts[4]:
                try:
                    thumb = png_from_bytes(base64.b64decode(parts[4]))
                except Exception:
                    thumb = None
            disp = _ses_unesc(parts[5]) if len(parts) >= 6 else ""
            loaded.append((_ses_unesc(parts[1]), nums, thumb, disp))
        if not loaded:
            raise ValueError(T("Nie znaleziono widok\u00f3w w pliku sesji.", "No views found in the session file."))
        self.clear()
        for name, nums, thumb, disp in loaded:
            self.new_record(nums, name=name, thumb=thumb, disp=disp)
        self.disp_applied = 0
        return len(loaded), sum(1 for x in loaded if x[3])

    KEYS = ("shot_dir", "prefix", "fmt", "model_dir", "subfolder", "bg", "report", "use_stored", "presets")

    def to_dict(self):
        return dict((k, getattr(self, k)) for k in self.KEYS)

    def from_dict(self, d):
        assign_attrs(self, d, self.KEYS, {"fmt": ("png", "jpg", "bmp"), "bg": tuple(k for k, _, _, _ in BG_MODES)})


def _ses_esc(s):
    return ("%s" % s).replace("\\", "\\\\").replace("\t", "\\t").replace("\n", "\\n").replace("\r", "\\r")


def _ses_unesc(s):
    out, i = [], 0
    while i < len(s):
        c = s[i]
        if c == "\\" and i + 1 < len(s):
            n = s[i + 1]
            out.append({"t": "\t", "n": "\n", "r": "\r", "\\": "\\"}.get(n, n))
            i += 2
            continue
        out.append(c)
        i += 1
    return "".join(out)


def views_html(rows):
    """Samodzielny raport widokow (podglady osadzone) + blok sesji HMSESSION."""
    maxn = max([len(r["view"]) for r in rows] + [0])
    H = ['<!DOCTYPE html><html><head><meta charset="utf-8"><title>%s</title><style>' % h_esc(T("Raport widok\u00f3w", "Views report")),
         "body{font-family:Arial,sans-serif;font-size:12px;color:#222}table{border-collapse:collapse}"
         "th,td{border:1px solid #bbb;padding:4px 6px;vertical-align:top}th{background:#1c5a96;color:#fff}"
         "img{display:block;border:1px solid #ccc}code{font-family:Consolas,monospace;font-size:11px;white-space:nowrap}"
         "td.n{font-weight:bold}</style></head><body>",
         "<h2>%s</h2><p>%s &middot; %s: %d</p>" % (h_esc(T("Raport widok\u00f3w (HyperMesh)", "Views report (HyperMesh)")),
                                                  h_esc(T("Wygenerowano: %s", "Generated: %s", now_text())), h_esc(T("widok\u00f3w", "views")), len(rows)),
         '<p style="font-size:11px;color:#777">%s</p>' % h_esc(T("Ten plik mo\u017cna wczyta\u0107 z powrotem (\u201eWczytaj sesj\u0119\u2026\u201d), aby odtworzy\u0107 list\u0119 widok\u00f3w z podgl\u0105dami.",
                                                               "This file can be loaded back (\u201cLoad session\u2026\u201d) to restore the list of views with previews.")),
         "<!--HMSESSION v3", "%d" % len(rows)]
    for r in rows:
        b64 = base64.b64encode(r["thumb"]).decode("ascii") if r.get("thumb") else ""
        H.append("R\t%s\t%s\tpng\t%s\t%s" % (_ses_esc(r["name"]), _ses_esc(view_text(r["view"])), b64, _ses_esc(r.get("disp", ""))))
    H.append("HMSESSION-->")
    H.append("<table><tr><th>Nr</th><th>%s</th><th>%s</th><th>VS</th>%s</tr>" % (
        h_esc(T("Podgl\u0105d", "Preview")), h_esc(T("Nazwa", "Name")), "".join("<th>v%d</th>" % (i + 1) for i in range(maxn))))
    for i, r in enumerate(rows, 1):
        src = ""
        p = r.get("path") or ""
        if p and os.path.isfile(p):
            src = "data:image/%s;base64,%s" % (mime_of(p), file_b64(p))
        elif r.get("thumb"):
            src = "data:image/png;base64,%s" % base64.b64encode(r["thumb"]).decode("ascii")
        img = '<img src="%s" style="max-width:360px;max-height:240px">' % src if src else "<i>%s</i>" % h_esc(T("(brak podgl\u0105du)", "(no preview)"))
        nm = h_esc(r["name"]) + (" &#9881;" if r.get("disp") else "")
        H.append('<tr><td>%d</td><td>%s</td><td class="n">%s</td><td><code>%s</code></td>%s</tr>' % (
            i, img, nm, h_esc(view_text(r["view"])), "".join("<td>%.6g</td>" % v for v in r["view"])))
    H.append("</table><h3>%s</h3>" % h_esc(T("Wszystkie widoki (VS) \u2013 do skopiowania", "All views (VS) \u2013 to copy")))
    H.append('<textarea readonly rows="%d" style="width:100%%;font-family:Consolas,monospace;font-size:11px">%s</textarea>' % (
        max(3, min(20, len(rows))), "\n".join(h_esc(view_text(r["view"])) for r in rows)))
    H.append("</body></html>\n")
    return "\n".join(H)


VIEWS = ViewStore()


# ============ UKLAD SLAJDU (PRYMITYWY WSPOLNE DLA PODGLADU I PPTX) =====
# Slajd jest opisany lista prymitywow w EMU (1 mm = 36000 EMU):
#   ("rect",  x, y, w, h, wypelnienie, obrys)
#   ("text",  x, y, w, h, tekst, rozmiar_pt, pogrubienie, kolor, l/ctr/r, t/ctr/b)
#   ("image", x, y, w, h, sciezka_pliku)
#   ("table", x, y, w, h, wiersze, udzialy_kolumn, rozmiar_pt)
# Te same prymitywy rysuje PODGLAD SLAJDU (QPainter) i zapis PPTX
# (natywne ksztalty PowerPointa) - to, co widac w podgladzie, trafia na slajd.
# Kolory: "RRGGBB" albo "" (brak). Komorka tabeli: tekst albo slownik
# {"t": tekst, "bold": bool, "color": "RRGGBB", "fill": "RRGGBB", "align": l/ctr/r}.
MM = 36000.0


class SlideStyle(object):
    """Opcje ukladu slajdu (z karty PPTX / okna podgladu)."""

    LAYOUTS = ("right", "left", "full", "custom")
    DEFAULT_BOX = [0.03, 0.16, 0.62, 0.78]

    def __init__(self):
        self.layout = "right"      # right | left | full | custom (wlasny kadr)
        self.title_on = True
        self.stats_on = True
        self.box = list(self.DEFAULT_BOX)   # wlasny kadr obrazu: x, y, w, h jako ulamki slajdu

    def box_rect(self, sw, sh):
        b = self.box if len(self.box) == 4 else self.DEFAULT_BOX
        return b[0] * sw, b[1] * sh, b[2] * sw, b[3] * sh

    def set_box(self, x, y, w, h):
        w = max(0.05, min(1.0, w))
        h = max(0.05, min(1.0, h))
        x = max(0.0, min(1.0 - w, x))
        y = max(0.0, min(1.0 - h, y))
        self.box = [round(x, 4), round(y, 4), round(w, 4), round(h, 4)]


def fit_rect(x, y, w, h, iw, ih):
    """Prostokat obrazu wpisany w (x, y, w, h) z zachowaniem proporcji."""
    if iw <= 0 or ih <= 0:
        return x, y, w, h
    s = min(w / float(iw), h / float(ih))
    fw, fh = iw * s, ih * s
    return x + (w - fw) / 2.0, y + (h - fh) / 2.0, fw, fh


def cell_of(cell):
    if isinstance(cell, dict):
        d = {"t": "%s" % cell.get("t", ""), "bold": bool(cell.get("bold")), "color": cell.get("color", "262626"),
             "fill": cell.get("fill", ""), "align": cell.get("align", "")}
    else:
        d = {"t": "" if cell is None else "%s" % cell, "bold": False, "color": "262626", "fill": "", "align": ""}
    return d


def slide_ops(item, sw, sh, style):
    m = 7 * MM
    ops = []
    y0 = m
    title = (item.get("title") or "").strip()
    if style.title_on and title:
        ops.append(("text", m, 5 * MM, sw - 2 * m, 12 * MM, title, 22, True, "1C5A96", "l", "ctr"))
        ops.append(("rect", m, 17.5 * MM, sw - 2 * m, 0.7 * MM, "1C5A96", ""))
        y0 = 21 * MM
    bottom = sh - m
    note = item.get("note") or ""
    if note:
        ops.append(("text", m, sh - m - 6 * MM, sw - 2 * m, 6 * MM, note, 9, False, "666666", "l", "ctr"))
        bottom = sh - m - 7 * MM
    aw, ah = sw - 2 * m, bottom - y0
    tb = item.get("table")
    if tb:
        rows = tb["rows"]
        nr = len(rows)
        rh = min(10.0 * MM, ah / float(max(1, nr)))
        fs = 9 if nr > 14 else 10 if nr > 9 else 12
        ops.append(("table", m, y0, aw, rh * nr, rows, tb.get("fr"), fs))
        return ops
    leg = item.get("legend")
    stats = item.get("stats") if style.stats_on else []
    if style.layout == "custom":
        # wlasny kadr: obraz dokladnie tam, gdzie go przeciagnieto; legenda
        # w wiekszej wolnej strefie obok (prawo / lewo), gdy jest na nia miejsce
        bx, by, bw, bh = style.box_rect(sw, sh)
        ops += image_ops(item, bx, by, bw, bh)
        if leg or stats:
            gap = 4 * MM
            right = (m + aw) - (bx + bw) - gap
            left = bx - m - gap
            if max(right, left) >= 40 * MM:
                if right >= left:
                    lx, lw = bx + bw + gap, right
                else:
                    lx, lw = m, left
                ops += side_ops(leg, stats, lx, y0, lw, ah)
        return ops
    if style.layout != "full" and (leg or stats):
        lw = max(78 * MM, 0.28 * aw)
        gap = 5 * MM
        iw = aw - lw - gap
        if style.layout == "left":
            lx, ix = m, m + lw + gap
        else:
            ix, lx = m, m + iw + gap
        ops += image_ops(item, ix, y0, iw, ah)
        ops += side_ops(leg, stats, lx, y0, lw, ah)
    else:
        ops += image_ops(item, m, y0, aw, ah)
    return ops


def image_ops(item, x, y, w, h):
    img = item.get("img") or ""
    if not img or not os.path.isfile(img):
        return [("rect", x, y, w, h, "F2F2F2", "BFBFBF"),
                ("text", x, y, w, h, T("(brak obrazu)", "(no image)"), 14, False, "999999", "ctr", "ctr")]
    iw, ih = image_size(img)
    fx, fy, fw, fh = fit_rect(x, y, w, h, iw, ih)
    return [("image", fx, fy, fw, fh, img)]


def side_ops(leg, stats, x, y, w, h):
    ops = []
    cy = y
    stat_h = min(0.45 * h, (len(stats) + 1) * 6.6 * MM) if stats else 0
    if leg:
        t1, t2 = leg["head"]
        ops.append(("text", x, cy, w, 7 * MM, t1, 14, True, "1C5A96", "l", "ctr"))
        cy += 7 * MM
        if t2:
            ops.append(("text", x, cy, w, 6 * MM, t2, 9, False, "555555", "l", "t"))
            cy += 7 * MM
        avail = y + h - cy - (stat_h + 4 * MM if stat_h else 0)
        if leg["type"] == "bar":
            ops += bar_ops(leg, x, cy, w, avail)
        else:
            rows = leg["rows"]
            if rows:
                rh = min(8.5 * MM, avail / float(len(rows)))
                fs = max(7.0, min(12.0, rh / 12700.0 * 0.5))
                swc = min(10 * MM, 1.6 * rh)
                cw = 0.36 * w
                for r in rows:
                    pad = 0.12 * rh
                    ops.append(("rect", x, cy + pad, swc, rh - 2 * pad, hex6(r["rgb"]), "7F7F7F"))
                    ops.append(("text", x + swc + 2 * MM, cy, w - swc - 2 * MM - cw, rh, r["label"], fs, False, "262626", "l", "ctr"))
                    cnt = r.get("count")
                    if cnt is not None and cnt != "":
                        ct = "%s" % cnt
                        pc = r.get("pct")
                        if pc not in (None, ""):
                            ct = ("%s  (%.2f%%)" if 0 < pc < 1.0 else "%s  (%.1f%%)") % (cnt, pc)
                        ops.append(("text", x + w - cw, cy, cw, rh, ct, fs, False, "404040", "r", "ctr"))
                    cy += rh
    if stats:
        rows = [[T("Statystyka", "Statistic"), T("Warto\u015b\u0107", "Value")]] + [list(p) for p in stats]
        ops.append(("table", x, y + h - stat_h, w, stat_h, rows, [1.6, 1.2], 9))
    return ops


def bar_ops(leg, x, y, w, h):
    """Pionowy pasek gradientu (duzo pasm) + podzialki + wiersze dodatkowe."""
    ops = []
    extra = leg.get("rows") or []
    bh = h - len(extra) * 7.0 * MM - 3 * MM
    if bh < 20 * MM:
        bh = max(10 * MM, h * 0.6)
    bands = leg["bands"]
    K = len(bands)
    bw = 10 * MM
    lo, hi = bands[0]["lo"], bands[-1]["hi"]
    span = (hi - lo) or 1.0
    for b in bands:
        ya = y + bh - (b["lo"] - lo) / span * bh
        yz = y + bh - (b["hi"] - lo) / span * bh
        ops.append(("rect", x, yz, bw, max(1, ya - yz), hex6(b["rgb"]), ""))
    ops.append(("rect", x, y, bw, bh, "", "404040"))
    every = max(1, int(math.ceil(K / 9.0)))
    dec = leg.get("dec", 3)
    for i in list(range(0, K, every)) + [K]:
        v = bands[i]["lo"] if i < K else hi
        yy = y + bh - (v - lo) / span * bh
        ops.append(("text", x + bw + 2 * MM, yy - 3 * MM, w - bw - 2 * MM, 6 * MM, fmt_num(v, dec), 9, False, "262626", "l", "ctr"))
    cy = y + bh + 3 * MM
    for r in extra:
        ops.append(("rect", x, cy + 1 * MM, bw, 5 * MM, hex6(r["rgb"]), "7F7F7F"))
        ops.append(("text", x + bw + 2 * MM, cy, w - bw - 2 * MM, 7 * MM, "%s: %s" % (r["label"], r.get("count", "")), 10, False, "262626", "l", "ctr"))
        cy += 7 * MM
    return ops


# ===================== PREZENTACJA POWERPOINT (.pptx) ==================
# Zapis przez python-pptx (jest w Pythonie HyperMesha) - bez PowerPointa
# i bez COM. Nowa prezentacja 16:9 albo DOPISANIE slajdow na koncu
# istniejacej (np. firmowego szablonu: wybierany jest uklad z najmniejsza
# liczba pol zastepczych, pozostale pola sa usuwane). Przed dopisaniem
# powstaje kopia <nazwa>.bak.pptx; zapis idzie do pliku tymczasowego
# i dopiero potem podmienia docelowy (przerwany zapis nie psuje pliku).
# Gdy plik jest otwarty w PowerPoincie (zablokowany), prezentacja trafia
# obok z godzina w nazwie. Ksztalty dostaja nazwy "HMQS|..." - latwo je
# odnalezc w panelu zaznaczenia PowerPointa.
def need_pptx():
    if pptx is None:
        raise RuntimeError(T("Brak biblioteki python-pptx w Pythonie HyperMesha.", "python-pptx library missing in the HyperMesh Python."))


class PptxDeck(object):
    def __init__(self, path=None):
        need_pptx()
        if path:
            self.prs = pptx.Presentation(path)
        else:
            self.prs = pptx.Presentation()
            self.prs.slide_width = Emu(PPT_SLIDE_W)
            self.prs.slide_height = Emu(PPT_SLIDE_H)
        self.n = 0

    @property
    def size(self):
        return int(self.prs.slide_width), int(self.prs.slide_height)

    def blank_layout(self):
        best, bn = None, None
        for lay in self.prs.slide_layouts:
            try:
                n = len(lay.placeholders)
                nm = (lay.name or "").lower()
            except Exception:
                continue
            if nm in ("blank", "pusty", "leer", "vide"):
                n -= 0.5
            if bn is None or n < bn:
                best, bn = lay, n
        return best or self.prs.slide_layouts[0]

    def add_slide(self, ops):
        slide = self.prs.slides.add_slide(self.blank_layout())
        for ph in list(slide.placeholders):
            el = ph._element
            el.getparent().remove(el)
        self.n += 1
        for i, op in enumerate(ops, 1):
            kind = op[0]
            x, y, w, h = [int(round(v)) for v in op[1:5]]
            w, h = max(w, 1), max(h, 1)
            if kind == "rect":
                self._rect(slide, x, y, w, h, op[5], op[6], "HMQS|rect%d" % i)
            elif kind == "text":
                self._text(slide, x, y, w, h, op[5:], "HMQS|text%d" % i)
            elif kind == "image":
                if op[5] and os.path.isfile(op[5]):
                    pic = slide.shapes.add_picture(op[5], x, y, w, h)
                    pic.name = "HMQS|image%d" % i
            elif kind == "table":
                self._table(slide, x, y, w, h, op[5], op[6], op[7], "HMQS|table%d" % i)
        return slide

    @staticmethod
    def _rect(slide, x, y, w, h, fill, line, name):
        shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
        shp.name = name
        shp.shadow.inherit = False
        if fill:
            shp.fill.solid()
            shp.fill.fore_color.rgb = RGBColor.from_string(fill)
        else:
            shp.fill.background()
        if line:
            shp.line.color.rgb = RGBColor.from_string(line)
            shp.line.width = Pt(0.75)
        else:
            shp.line.fill.background()
        shp.text_frame.text = ""

    @staticmethod
    def _text(slide, x, y, w, h, spec, name):
        txt, size, bold, color, align, anchor = spec
        tb = slide.shapes.add_textbox(x, y, w, h)
        tb.name = name
        tf = tb.text_frame
        tf.word_wrap = True
        tf.auto_size = None
        for side in ("margin_left", "margin_right"):
            setattr(tf, side, Emu(int(0.8 * MM)))
        tf.margin_top = tf.margin_bottom = Emu(int(0.3 * MM))
        tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "b": MSO_ANCHOR.BOTTOM}.get(anchor, MSO_ANCHOR.MIDDLE)
        lines = ("%s" % txt).split("\n")
        for n, line in enumerate(lines):
            p = tf.paragraphs[0] if n == 0 else tf.add_paragraph()
            p.alignment = {"l": PP_ALIGN.LEFT, "r": PP_ALIGN.RIGHT}.get(align, PP_ALIGN.CENTER)
            run = p.add_run()
            run.text = line
            f = run.font
            f.size = Pt(float(size))
            f.bold = bool(bold)
            f.name = "Arial"
            f.color.rgb = RGBColor.from_string(color or "000000")

    @staticmethod
    def _table(slide, x, y, w, h, rows, fr, size, name):
        nr, nc = len(rows), max(len(r) for r in rows)
        gf = slide.shapes.add_table(nr, nc, x, y, w, h)
        gf.name = name
        tbl = gf.table
        fr = list(fr) if fr and len(fr) == nc else [1.0] * nc
        tot = float(sum(fr))
        acc = 0
        for c in range(nc):
            cw = int(w * fr[c] / tot) if c < nc - 1 else w - acc
            tbl.columns[c].width = Emu(cw)
            acc += cw
        rh = int(h / float(nr))
        for r in range(nr):
            tbl.rows[r].height = Emu(rh)
        tbl.first_row = True
        tbl.horz_banding = False
        for r, row in enumerate(rows):
            for c in range(nc):
                d = cell_of(row[c] if c < len(row) else "")
                cell = tbl.cell(r, c)
                cell.margin_left = cell.margin_right = Emu(int(1.0 * MM))
                cell.margin_top = cell.margin_bottom = Emu(int(0.3 * MM))
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                fill = "1C5A96" if r == 0 else (d["fill"] or ("F4F7FA" if r % 2 else "FFFFFF"))
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor.from_string(fill)
                tf = cell.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                al = "ctr" if r == 0 else (d["align"] or ("l" if c == 0 else "ctr"))
                p.alignment = {"l": PP_ALIGN.LEFT, "r": PP_ALIGN.RIGHT}.get(al, PP_ALIGN.CENTER)
                run = p.add_run()
                run.text = d["t"]
                f = run.font
                f.size = Pt(float(size))
                f.name = "Arial"
                f.bold = r == 0 or d["bold"]
                f.color.rgb = RGBColor.from_string("FFFFFF" if r == 0 else (d["color"] or "262626"))

    def save(self, path):
        folder = os.path.dirname(os.path.abspath(path)) or "."
        tmp = os.path.join(folder, ".~hmqs_%d.pptx" % os.getpid())
        self.prs.save(tmp)
        try:
            os.replace(tmp, path)
        except OSError:
            try:
                os.remove(tmp)
            except OSError:
                pass
            raise


def pptx_verify(path, expect_min=1):
    """Samokontrola: plik da sie otworzyc i ma slajdy. Lista problemow."""
    probs = []
    try:
        prs = pptx.Presentation(path)
        n = len(prs.slides)
        if n < expect_min:
            probs.append(T("za ma\u0142o slajd\u00f3w (%d)", "too few slides (%d)", n))
        with zipfile.ZipFile(path) as z:
            bad = z.testzip()
            if bad:
                probs.append(T("uszkodzony wpis ZIP: %s", "corrupt ZIP entry: %s", bad))
    except Exception as e:
        probs.append("%s" % e)
    return probs


def pptx_slide_size(path):
    try:
        prs = pptx.Presentation(path)
        return int(prs.slide_width), int(prs.slide_height)
    except Exception:
        return PPT_SLIDE_W, PPT_SLIDE_H


# ================= ELEMENTY PREZENTACJI (SLAJDY Z MODULOW) =============
# Slajd = slownik: kind, title, img (plik obrazu), own_img (usunac po
# eksporcie), legend, stats, table, note. Zrodla slajdow:
#  - SERIA: kazda metryka (+ widok zbiorczy) x kamera (biezaca albo
#    zaznaczone widoki z karty "Widoki i zrzuty") -> obraz + legenda +
#    tabela statystyk; opcjonalnie slajd zbiorczy z tabela wszystkich metryk,
#  - biezacy widok, widoki z karty "Widoki i zrzuty" (klatki WYSIWYG),
#    wynik delty REF/INF (obraz + legenda pasm), raport jakosci (tabela).
# deliver(): PODGLAD SLAJDU (jesli wlaczony: przegladanie, zmiana tytulow
# i ukladu, usuwanie slajdow, Utworz / Anuluj) -> zapis PPTX -> sprzatanie.
class Presenter(object):
    def __init__(self):
        self.ppt_file = ""
        self.mode = "append"          # append | new
        self.preview_on = True
        self.style = SlideStyle()
        self.title_tpl = ""
        self.summary_on = True
        self.combined_on = True
        self.cam_src = "current"      # current | shots
        self.img_fmt = "png"          # png | jpg
        self.keep_images = False
        self.leg_in_shots = True
        self.preview_cb = None        # f(items, ask) -> (ok, items)   [GUI]
        self.ask_path_cb = None       # f() -> sciezka                 [GUI]
        self.open_cb = None           # f(tekst, sciezka)              [GUI]
        self.render_cb = None         # f(items) -> lista PNG slajdow  [GUI, galeria po eksporcie]
        self.sample_path = ""         # klatka do podgladu (cache)
        self.sample_key = None
        self.last_export = None       # {"path", "n", "when", "msg", "titles", "thumbs"}

    # ---------------------------------------------------------- podglad
    def state_key(self):
        """Podpis stanu siatki i widoku - zmiana = nowa klatka do podgladu."""
        return (MQ.view, MQ.when, DELTA.done, DELTA.when, MESH.owner, VIEWS.bg, self.img_fmt,
                HM.signature() if HM.ok() else "")

    def sample_frame(self, capture=True):
        """Klatka do podgladu slajdu: ostatnia (gdy stan sie nie zmienil),
        inaczej swiezy zrzut okna graficznego - raz na stan, nie przy kazdym
        odswiezeniu podgladu."""
        key = self.state_key()
        if self.sample_path and os.path.isfile(self.sample_path) and self.sample_key == key:
            return self.sample_path
        if not capture or not HM.ok():
            return self.sample_path if (self.sample_path and os.path.isfile(self.sample_path)) else ""
        p = self.capture()
        if p:
            old = self.sample_path
            self.sample_path, self.sample_key = p, key
            if old and old != p and os.path.isfile(old):
                try:
                    os.remove(old)
                except OSError:
                    pass
        return self.sample_path

    def invalidate_frame(self):
        self.sample_key = None

    def sample_item(self, capture=True):
        """Przykladowy slajd z BIEZACEGO stanu (to, co da "Biezacy widok ->
        slajd"): obraz z klatki podgladu, legenda i statystyki aktywnego modulu."""
        img = self.sample_frame(capture)
        model = HM.model_name() if HM.ok() else "model"
        if MQ.view and MQ.analyzed:
            it = self.metric_item(MQ.view, img)
        elif DELTA.done and MESH.owner == "delta":
            it = self.delta_item(img)
        else:
            it = self.plain_item(self.fill_tpl(T("widok", "view"), "", model), img, own=False)
        it["own_img"] = False
        return it

    def plan_lines(self):
        """CO POWSTANIE z biezacych opcji - do karty "Podglad slajdu" i okna
        "Podglad na zywo" (jak w makrze ANSYS SHOTS)."""
        L = []
        p = self.target(False)
        L.append(T("PLIK: %s", "FILE: %s", p or T("(wybierzesz przy zapisie)", "(chosen when saving)")))
        if p and os.path.isfile(p) and self.mode == "append":
            L.append(T("  tryb: dopisanie slajd\u00f3w na ko\u0144cu (kopia .bak.pptx)", "  mode: append slides at the end (.bak.pptx backup)"))
        elif p and os.path.isfile(p):
            L.append(T("  tryb: NOWY plik \u2013 istniej\u0105cy zostanie zast\u0105piony", "  mode: NEW file \u2013 the existing one is replaced"))
        else:
            L.append(T("  tryb: nowa prezentacja 16:9", "  mode: new 16:9 presentation"))
        lay = {"right": T("obraz + legenda z prawej", "image + legend right"), "left": T("legenda z lewej", "legend left"),
               "full": T("sam obraz", "image only"), "custom": T("w\u0142asny kadr (%d%% x %d%% slajdu)", "custom frame (%d%% x %d%% of the slide)",
                                                               int(round(self.style.box[2] * 100)), int(round(self.style.box[3] * 100)))}.get(self.style.layout, self.style.layout)
        L.append(T("UK\u0141AD: %s; tytu\u0142: %s; tabela statystyk: %s", "LAYOUT: %s; title: %s; statistics table: %s",
                   lay, T("tak", "yes") if self.style.title_on else T("nie", "no"), T("tak", "yes") if self.style.stats_on else T("nie", "no")))
        L.append(T("  szablon tytu\u0142u: %s", "  title template: %s", (self.title_tpl or "").strip() or self.default_tpl()))
        L.append("")
        if MQ.analyzed:
            views = list(MQ.analyzed) + (["all"] if (self.combined_on and MQ.combined_ok()) else [])
            if self.cam_src == "shots":
                cams = len(VIEWS.selected()) or 1
                cam_t = T("%d zaznaczonych widok\u00f3w z karty \u201eWidoki\u201d", "%d ticked views from the \u201cViews\u201d tab", cams)
            else:
                cams, cam_t = 1, T("bie\u017c\u0105ca kamera", "current camera")
            n = len(views) * cams + (1 if self.summary_on else 0)
            L.append(T("SERIA (F7): %d slajd(\u00f3w) = %d metryk \u00d7 %s%s", "SERIES (F7): %d slide(s) = %d metrics \u00d7 %s%s",
                       n, len(views), cam_t, T(" + slajd zbiorczy", " + summary slide") if self.summary_on else ""))
            for v in views:
                L.append("   \u2022 %s \u2013 %s" % (MQ.label(v), MQ.thr_text(v) or T("widok zbiorczy", "combined view")))
            L.append(T("  na siatce teraz: %s (analiza %s, %d element\u00f3w)", "  on the mesh now: %s (analysis %s, %d elements)",
                       MQ.label(MQ.view) if MQ.view else T("(brak \u2013 kolory zdj\u0119te)", "(none \u2013 colors removed)"), MQ.when, len(MQ.elems)))
        else:
            L.append(T("SERIA (F7): brak analizy metryk \u2013 seria nie powstanie (karta \u201e1 Metryki\u201d, F5)",
                       "SERIES (F7): no metrics analysis \u2013 no series (\u201c1 Metrics\u201d tab, F5)"))
        L.append(T("DELTA REF/INF: %s", "REF/INF DELTA: %s", T("wynik na siatce (%s) \u2013 slajd dost\u0119pny", "result on the mesh (%s) \u2013 slide available", DELTA.delta_label())
                   if (DELTA.done and MESH.owner == "delta") else T("brak wyniku na siatce", "no result on the mesh")))
        L.append(T("RAPORT: %s", "REPORT: %s", T("wynik z %s \u2013 slajd z tabel\u0105 dost\u0119pny", "result from %s \u2013 table slide available", REPORT.last["when"])
                   if REPORT.last else T("brak wygenerowanego raportu", "no report generated")))
        L.append(T("WIDOKI: %d zapami\u0119tanych, %d zaznaczonych, %d z klatk\u0105 WYSIWYG", "VIEWS: %d remembered, %d ticked, %d with a WYSIWYG frame",
                   len(VIEWS.views), len(VIEWS.selected()), sum(1 for r in VIEWS.views if VIEWS.shot_ok(r))))
        if self.last_export:
            L.append("")
            L.append(T("OSTATNI EKSPORT: %s \u2013 %d slajd(\u00f3w), %s", "LAST EXPORT: %s \u2013 %d slide(s), %s",
                       os.path.basename(self.last_export["path"]), self.last_export["n"], self.last_export["when"]))
        return L

    # ---------------------------------------------------------- tytuly
    @staticmethod
    def default_tpl():
        return T("Jako\u015b\u0107 siatki \u2013 {metric} \u2013 {view}", "Mesh quality \u2013 {metric} \u2013 {view}")

    def fill_tpl(self, metric, thr, view):
        t = (self.title_tpl or "").strip() or self.default_tpl()
        dim = "+".join(d for d, on in (("2D", MQ.dim2), ("3D", MQ.dim3)) if on)
        s = (t.replace("{metric}", metric).replace("{thr}", thr).replace("{view}", view)
             .replace("{model}", HM.model_name() if HM.ok() else "model").replace("{date}", now_text("%Y-%m-%d"))
             .replace("{dim}", dim))
        s = re.sub("\\s+\u2013\\s*$", "", s)
        s = re.sub("\\s+\u2013\\s+\u2013\\s+", " \u2013 ", s)
        return s.strip()

    # ---------------------------------------------------------- obraz
    def capture(self):
        """Zrzut okna graficznego do pliku roboczego (tlo jak w module Widoki)."""
        HM.redraw()
        time.sleep(CAPTURE_DELAY_MS / 1000.0)
        png = tmp_file("png", "slide")
        if not HM.capture(png):
            return ""
        hx = bg_hex(VIEWS.bg)
        if hx:
            try:
                recolor_background(png, hx)
            except Exception:
                pass
        if self.img_fmt == "jpg":
            jp = os.path.splitext(png)[0] + ".jpg"
            try:
                convert_image(png, jp)
                os.remove(png)
                return jp
            except Exception:
                return png
        return png

    # ---------------------------------------------------------- slajdy
    @staticmethod
    def mq_legend(v):
        return {"type": "rows", "head": MQ.legend_head(v), "rows": MQ.legend_rows(v)}

    def metric_item(self, v, img, cam=""):
        return {"kind": "metric", "view": v, "img": img, "own_img": True,
                "title": self.fill_tpl(MQ.label(v), MQ.thr_text(v), cam or (HM.model_name() if HM.ok() else "")),
                "legend": self.mq_legend(v), "stats": MQ.stat_pairs(v), "table": None, "note": ""}

    def summary_item(self):
        tot = len(MQ.elems)
        note = T("Model: %s \u2022 element\u00f3w: %d (2D: %d, 3D: %d) \u2022 %s", "Model: %s \u2022 elements: %d (2D: %d, 3D: %d) \u2022 %s",
                 HM.model_name() if HM.ok() else "-", tot, MQ.n2d, MQ.n3d, MQ.when)
        if MQ.jz_note:
            note += " \u2022 Jacobian Zero: " + MQ.jz_note
        return {"kind": "summary", "img": "", "own_img": False, "legend": None, "stats": [],
                "title": T("Jako\u015b\u0107 siatki \u2013 podsumowanie metryk \u2013 %s", "Mesh quality \u2013 metrics summary \u2013 %s", HM.model_name() if HM.ok() else ""),
                "table": {"rows": MQ.summary_rows(), "fr": [2.2, 1.2, 1.1, 1.1, 1.1, 0.9, 1, 1, 1.2]}, "note": note}

    def plain_item(self, title, img, own=True):
        leg = self.mq_legend(MQ.view) if (MQ.view and self.leg_in_shots) else None
        stats = MQ.stat_pairs(MQ.view) if MQ.view else []
        return {"kind": "view", "title": title, "img": img, "own_img": own, "legend": leg, "stats": stats,
                "table": None, "note": ""}

    def delta_item(self, img):
        lm = DELTA.legend_model()
        if lm["fail"] or len(lm["bands"]) <= 16:
            rows = [{"rgb": b["rgb"], "label": b["label"], "count": b["count"], "pct": None} for b in reversed(lm["bands"])]
            if lm["fail"]:
                rows = []
            leg = {"type": "rows", "head": (lm["title"], lm["sub"]), "rows": rows + lm["extra"]}
        else:
            leg = {"type": "bar", "head": (lm["title"], lm["sub"]), "bands": lm["bands"], "rows": lm["extra"], "dec": lm["dec"]}
        return {"kind": "delta", "title": self.fill_tpl(DELTA.legend_title(), "", HM.model_name() if HM.ok() else ""),
                "img": img, "own_img": True, "legend": leg, "stats": DELTA.stat_pairs(), "table": None, "note": ""}

    def report_item(self):
        run = REPORT.last
        if not run:
            return None
        rows = [[T("Metryka", "Metric"), "N", "Min", "Max", T("\u015arednia", "Mean"), T("Pr\u00f3g", "Threshold"), T("Do poprawy", "To fix"), "%"]]
        for m in run["exp"]:
            thr = ("%s %s" % (_dirsym(m["dir"]), fmt_num(m["thr"], 3))) if m["dir"] != "none" else "\u2013"
            rows.append([m["label"], "%d" % m["n"], fmt_num(m["min"], 4), fmt_num(m["max"], 4), fmt_num(m["mean"], 4), thr,
                         {"t": "%d" % m["bad"], "color": "C00000" if m["bad"] else "2E7D32", "bold": True}, "%.2f%%" % m["badpct"]])
        note = T("Model: %s \u2022 element\u00f3w: %d", "Model: %s \u2022 elements: %d", os.path.basename(run["source"] or "") or "-", run["nsel"])
        if run["score"]["has"]:
            note += T(" \u2022 wska\u017anik jako\u015bci: %.1f / 100 (%s)", " \u2022 quality score: %.1f / 100 (%s)", run["score"]["score"], run["score"]["grade"])
        return {"kind": "report", "title": T("Raport jako\u015bci siatki \u2013 %s", "Mesh quality report \u2013 %s", os.path.splitext(os.path.basename(run["source"] or ""))[0] or "model"),
                "img": "", "own_img": False, "legend": None, "stats": [],
                "table": {"rows": rows, "fr": [2.4, 1, 1, 1, 1, 1.1, 1.1, 0.9]}, "note": note}

    @staticmethod
    def drop_items(items):
        for it in items:
            if it.get("own_img") and it.get("img") and os.path.isfile(it["img"]):
                try:
                    os.remove(it["img"])
                except OSError:
                    pass

    # ---------------------------------------------------------- zrodla slajdow
    def cameras(self):
        if self.cam_src != "shots":
            return [("", None)]
        out = [(r["name"], r["view"]) for r in VIEWS.selected()]
        if not out:
            BUS.status(T("Brak zaznaczonych widok\u00f3w na karcie \u201eWidoki i zrzuty\u201d \u2013 u\u017cyto bie\u017c\u0105cej kamery.",
                         "No views ticked on the \u201cViews & screenshots\u201d tab \u2013 using the current camera."), "warn")
            return [("", None)]
        return out

    def series_items(self):
        """Seria: kazda metryka (+ zbiorczo) x kamera -> slajdy."""
        if not MQ.analyzed or HM.signature() != MQ.sig:
            MQ.analyze()
        views = list(MQ.analyzed) + (["all"] if (self.combined_on and MQ.combined_ok()) else [])
        cams = self.cameras()
        v0 = MQ.view
        cur = HM.get_view()
        items = [self.summary_item()] if self.summary_on else []
        k, tot = 0, len(views) * len(cams)
        try:
            for v in views:
                MQ.apply_view(v)
                for name, nums in cams:
                    k += 1
                    BUS.progress(T("Zrzut %d / %d: %s %s", "Shot %d / %d: %s %s", k, tot, MQ.label(v), name))
                    if nums:
                        VIEWS.apply_view(nums)
                    items.append(self.metric_item(v, self.capture(), name))
        except Exception:
            self.drop_items(items)
            raise
        finally:
            if len(cur) >= 16:
                VIEWS.apply_view(cur)
            if v0 and v0 != MQ.view and v0 in MQ.views():
                try:
                    MQ.apply_view(v0)
                except Exception:
                    pass
        return items

    def current_item(self):
        img = self.capture()
        if not img:
            BUS.status(T("Nie uda\u0142o si\u0119 zrobi\u0107 zrzutu okna graficznego.", "Could not capture the graphics window."), "err")
        if MQ.view:
            return self.metric_item(MQ.view, img)
        return self.plain_item(self.fill_tpl(T("widok", "view"), "", HM.model_name() if HM.ok() else ""), img)

    def shots_items(self):
        recs = VIEWS.selected()
        if not recs:
            raise ValueError(T("Zaznacz widoki na karcie \u201eWidoki i zrzuty\u201d.", "Tick views on the \u201cViews & screenshots\u201d tab."))
        items, cur, live = [], None, 0
        for r in recs:
            if VIEWS.use_stored and VIEWS.shot_ok(r):
                items.append(self.plain_item(r["name"], r["shot"], own=False))
            else:
                if cur is None:
                    cur = HM.get_view()
                VIEWS.apply_view(r["view"])
                items.append(self.plain_item(r["name"], self.capture()))
                live += 1
        if live and cur and len(cur) >= 16:
            VIEWS.apply_view(cur)
        return items

    def delta_items(self):
        if MESH.owner != "delta" or not DELTA.done:
            raise ValueError(T("Na siatce nie ma wyniku delty \u2013 wykonaj analiz\u0119 na karcie \u201eDelta REF / INF\u201d.",
                               "No delta result on the mesh \u2013 run the analysis on the \u201cDelta REF / INF\u201d tab."))
        return [self.delta_item(self.capture())]

    def report_items(self):
        it = self.report_item()
        if it is None:
            raise ValueError(T("Brak wynik\u00f3w raportu \u2013 wygeneruj raport na karcie \u201eRaport jako\u015bci\u201d.",
                               "No report results \u2013 generate a report on the \u201cQuality report\u201d tab."))
        return [it]

    def summary_items(self):
        if not MQ.analyzed:
            raise ValueError(T("Najpierw wykonaj analiz\u0119 metryk.", "Run the metrics analysis first."))
        return [self.summary_item()]

    # ---------------------------------------------------------- zapis
    def target(self, ask=True):
        p = (self.ppt_file or "").strip()
        if not p and ask and self.ask_path_cb is not None:
            p = self.ask_path_cb() or ""
            if p:
                self.ppt_file = p
        if p and not p.lower().endswith(".pptx"):
            p += ".pptx"
        return p

    def slide_size(self):
        p = self.target(False)
        if self.mode == "append" and p and os.path.isfile(p):
            return pptx_slide_size(p)
        return PPT_SLIDE_W, PPT_SLIDE_H

    def deliver(self, items):
        """Podglad (jesli wlaczony) -> zapis -> galeria po eksporcie -> sprzatanie.
        Zwraca opis albo ""."""
        if not items:
            return ""
        if self.preview_on and self.preview_cb is not None:
            ok, items = self.preview_cb(items, True)
            if not ok:
                self.drop_items(items)
                BUS.status(T("Anulowano \u2013 nic nie zosta\u0142o zapisane.", "Cancelled \u2013 nothing was written."), "warn")
                return ""
        try:
            msg = self.export(items)
            if msg and self.last_export is not None:
                if self.render_cb is not None:
                    try:
                        self.last_export["thumbs"] = self.render_cb(items)
                    except Exception as e:
                        BUS.log("thumbs: %s" % e)
                if self.open_cb is not None:
                    self.open_cb(msg, self.last_export["path"])
            return msg
        finally:
            self.drop_items(items)

    def export(self, items):
        need_pptx()
        path = self.target(True)
        if not path:
            raise ValueError(T("Nie wskazano pliku prezentacji.", "No presentation file chosen."))
        append = self.mode == "append" and os.path.isfile(path)
        note = ""
        deck = None
        if append:
            try:
                deck = PptxDeck(path)
                bak = backup_file(path, "bak")
                if bak:
                    note = T(" (kopia: %s)", " (backup: %s)", os.path.basename(bak))
            except Exception as e:
                alt = "%s_%s.pptx" % (os.path.splitext(path)[0], now_text("%Y%m%d_%H%M%S"))
                if not BUS.confirm(T("Nie mo\u017cna dopisa\u0107 do:\n%s\n\n%s\n\nUtworzy\u0107 now\u0105 prezentacj\u0119:\n%s ?",
                                     "Cannot append to:\n%s\n\n%s\n\nCreate a new presentation:\n%s ?", path, e, alt), "PPTX"):
                    return ""
                path, append, deck = alt, False, None
        elif os.path.isfile(path):
            if not BUS.confirm(T("Plik istnieje i zostanie ZAST\u0104PIONY:\n%s\n\nKontynuowa\u0107?",
                                 "The file exists and will be REPLACED:\n%s\n\nContinue?", path), "PPTX"):
                return ""
        if deck is None:
            deck = PptxDeck()
        sw, sh = deck.size
        for it in items:
            deck.add_slide(slide_ops(it, sw, sh, self.style))
        BUS.progress(T("Zapis prezentacji\u2026", "Saving the presentation\u2026"))
        if file_locked(path):
            alt = "%s_%s.pptx" % (os.path.splitext(path)[0], now_text("%H%M%S"))
            note += T(" \u2013 UWAGA: %s by\u0142 zablokowany (otwarty w PowerPoincie?), zapisano jako %s",
                      " \u2013 WARNING: %s was locked (open in PowerPoint?), saved as %s", os.path.basename(path), os.path.basename(alt))
            path = alt
        deck.save(path)
        probs = pptx_verify(path, deck.n)
        if self.keep_images:
            d = os.path.splitext(path)[0] + "_obrazy"
            os.makedirs(d, exist_ok=True)
            for i, it in enumerate(items, 1):
                img = it.get("img")
                if img and os.path.isfile(img):
                    shutil.copy2(img, os.path.join(d, "%02d_%s%s" % (i, clean_file_name(it["title"], 60), os.path.splitext(img)[1])))
        msg = T("Zapisano %d slajd(\u00f3w) do %s%s", "Saved %d slide(s) to %s%s", deck.n, path, note)
        if probs:
            msg += T(" \u2013 samokontrola: %s", " \u2013 self-check: %s", "; ".join(probs))
        BUS.status(msg, "warn" if probs else "ok")
        BUS.log(msg)
        self.last_export = {"path": path, "n": deck.n, "when": now_text(), "msg": msg,
                            "titles": [it.get("title", "") for it in items], "thumbs": [], "probs": probs}
        return msg

    KEYS = ("ppt_file", "mode", "preview_on", "title_tpl", "summary_on", "combined_on", "cam_src", "img_fmt",
            "keep_images", "leg_in_shots")

    def to_dict(self):
        d = dict((k, getattr(self, k)) for k in self.KEYS)
        d.update({"layout": self.style.layout, "title_on": self.style.title_on, "stats_on": self.style.stats_on,
                  "box": list(self.style.box), "frames": [dict(f) for f in FRAME_PRESETS]})
        return d

    def from_dict(self, d):
        assign_attrs(self, d, self.KEYS, {"mode": ("append", "new"), "cam_src": ("current", "shots"),
                                         "img_fmt": ("png", "jpg")})
        assign_attrs(self.style, d, ("layout", "title_on", "stats_on"), {"layout": SlideStyle.LAYOUTS})
        box = d.get("box")
        if isinstance(box, list) and len(box) == 4 and all(isinstance(v, (int, float)) for v in box):
            self.style.set_box(*[float(v) for v in box])
        fr = d.get("frames")
        if isinstance(fr, list):
            del FRAME_PRESETS[:]
            for f in fr:
                if isinstance(f, dict) and isinstance(f.get("name"), str) and isinstance(f.get("box"), list) and len(f["box"]) == 4:
                    FRAME_PRESETS.append({"name": f["name"], "box": [float(v) for v in f["box"]]})


# Zapisane kadry wlasne (nazwa + prostokat) - jak "zapisane kadry" w ANSYS SHOTS
FRAME_PRESETS = []
QUICK_FRAMES = [("half_l", "1/2 lewa", "1/2 left", [0.02, 0.16, 0.47, 0.78]),
                ("half_r", "1/2 prawa", "1/2 right", [0.51, 0.16, 0.47, 0.78]),
                ("half_t", "1/2 g\u00f3ra", "1/2 top", [0.03, 0.16, 0.94, 0.39]),
                ("half_b", "1/2 d\u00f3\u0142", "1/2 bottom", [0.03, 0.57, 0.94, 0.39]),
                ("full", "pe\u0142ny", "full", [0.02, 0.16, 0.96, 0.80]),
                ("big_l", "2/3 lewa", "2/3 left", [0.02, 0.16, 0.64, 0.78])]


PRESENT = Presenter()


# ========================= USTAWIENIA UZYTKOWNIKA ======================
# Jeden plik JSON w katalogu uzytkownika (.hm_quality_studio.json): jezyk,
# progi i legendy metryk, opcje delty, raportu, widokow i prezentacji.
# Plik jest niezalezny od modelu - ustawienia wracaja przy kazdym starcie.
def settings_path():
    home = os.environ.get("USERPROFILE") or os.path.expanduser("~")
    return os.path.join(home, SETTINGS_FILE)


GUI_PREFS = {"tab": 0, "live_open": False, "live_geo": [], "win_geo": [], "log_open": False}


def save_settings():
    d = {"version": VERSION, "lang": lang(), "mq": MQ.to_dict(), "delta": DELTA.to_dict(),
         "report": REPORT.to_dict(), "views": VIEWS.to_dict(), "ppt": PRESENT.to_dict(), "gui": dict(GUI_PREFS)}
    p = settings_path()
    tmp = p + ".tmp"
    try:
        # zapis do pliku tymczasowego i podmiana: przerwany zapis nie psuje ustawien
        with io.open(tmp, "w", encoding="utf-8") as fh:
            json.dump(d, fh, indent=1, ensure_ascii=True)
        os.replace(tmp, p)
        return True
    except Exception:
        try:
            os.remove(tmp)
        except OSError:
            pass
        return False


def load_settings():
    p = settings_path()
    if not os.path.isfile(p):
        return False
    try:
        with io.open(p, "r", encoding="utf-8") as fh:
            d = json.load(fh)
    except Exception:
        return False
    if not isinstance(d, dict):
        return False
    set_lang(d.get("lang", lang()))
    for key, obj in (("mq", MQ), ("delta", DELTA), ("report", REPORT), ("views", VIEWS), ("ppt", PRESENT)):
        try:
            part = d.get(key)
            obj.from_dict(part if isinstance(part, dict) else {})
        except Exception:
            pass
    g = d.get("gui")
    if isinstance(g, dict):
        for k in GUI_PREFS:
            if k in g and isinstance(g[k], type(GUI_PREFS[k])):
                GUI_PREFS[k] = g[k]
    return True




# ================== GUI: WSPOLNE KONTROLKI I RYSOWANIE =================
# Okna sa w PyQt5 - tym samym Qt, na ktorym stoi interfejs HyperMesha 2024,
# wiec okno narzedzia jest NIEMODALNE (mozna obracac model, zapamietywac
# widoki) i nie blokuje HyperMesha. Plik laduje sie tez bez Qt (tryb
# wsadowy): klasy okien dziedzicza wtedy po "object" i po prostu nie sa
# uzywane.
# WAZNE (Qt): przyciski radiowe w JEDNYM rodzicu sa automatycznie wzajemnie
# wykluczajace - kazdy zestaw radiowy dostaje tu wlasna QButtonGroup
# (radio_group), inaczej wybor "2D / 3D" odznaczal metryke (blad z 2.1).
# Legendy sa opisane prostymi prymitywami (prostokat, tekst, trojkat) -
# ten sam opis rysuje okno (QPainter) i eksport SVG (zapis legendy do
# PowerPointa / dokumentacji), wiec plik SVG wyglada jak okno.
_Q = qt()
if _Q:
    QtCore, QtGui, QtWidgets = _Q
    _QWidget, _QDialog, _QMainWindow = QtWidgets.QWidget, QtWidgets.QDialog, QtWidgets.QMainWindow
else:                                    # pragma: no cover - tryb bez Qt
    QtCore = QtGui = QtWidgets = None
    _QWidget = _QDialog = _QMainWindow = object


def qcolor(c):
    if isinstance(c, (tuple, list)):
        return QtGui.QColor(int(c[0]), int(c[1]), int(c[2]))
    return QtGui.QColor(c if str(c).startswith("#") else "#" + str(c))


def app_instance():
    return QtWidgets.QApplication.instance()


def hm_main_window():
    """Glowne okno HyperMesha (rodzic naszych okien) albo None."""
    app = app_instance()
    if app is None:
        return None
    for w in app.topLevelWidgets():
        try:
            if isinstance(w, QtWidgets.QMainWindow) and (
                    "HyperMesh" in w.windowTitle() or w.objectName() == "Unity Main Window"):
                return w
        except Exception:
            pass
    return None


def screen_rect(widget=None):
    """Dostepny obszar ekranu (bez paska zadan) dla okna albo glowny ekran."""
    app = app_instance()
    if app is None:
        return QtCore.QRect(0, 0, 1600, 900)
    try:
        return app.desktop().availableGeometry(widget) if widget is not None else app.desktop().availableGeometry()
    except Exception:
        scr = app.primaryScreen()
        return scr.availableGeometry() if scr else QtCore.QRect(0, 0, 1600, 900)


def styled_button(text, color=None, big=False, parent=None):
    b = QtWidgets.QPushButton(text, parent)
    if color:
        dark = QtGui.QColor(color).darker(125).name()
        b.setStyleSheet("QPushButton{background:%s;color:white;font-weight:bold;padding:%s;border:1px solid %s;border-radius:3px}"
                        "QPushButton:hover{background:%s}QPushButton:disabled{background:#9aa6b2;color:#eef}"
                        % (color, "6px 14px" if big else "3px 10px", dark, dark))
        if big:
            f = b.font()
            f.setPointSizeF(f.pointSizeF() + 1.5)
            b.setFont(f)
    return b


def note_label(text, color="#555555"):
    lab = QtWidgets.QLabel(text)
    lab.setWordWrap(True)
    lab.setStyleSheet("color:%s" % color)
    f = lab.font()
    f.setPointSizeF(max(7.0, f.pointSizeF() - 0.5))
    lab.setFont(f)
    return lab


def step_label(num, text):
    """Naglowek kroku "1  Tekst" - numerowane kroki prowadza przez karte."""
    lab = QtWidgets.QLabel("<span style='background:%s;color:white;padding:1px 7px;border-radius:8px;font-weight:bold'>%s</span>"
                           "&nbsp; <b>%s</b>" % (HDR_COLOR, num, h_esc(text)))
    lab.setTextFormat(QtCore.Qt.RichText)
    return lab


def tip(widget, text):
    """Podpowiedz (dymek) - takze dla ukladow: dla wszystkich ich widgetow."""
    if isinstance(widget, QtWidgets.QLayout):
        for i in range(widget.count()):
            w = widget.itemAt(i).widget()
            if w is not None:
                w.setToolTip(text)
    else:
        widget.setToolTip(text)
    return widget


def hrow(*items, **kw):
    """Poziomy uklad: widgety, liczby (odstep) albo None (rozpychacz)."""
    lay = QtWidgets.QHBoxLayout()
    lay.setContentsMargins(0, 0, 0, 0)
    lay.setSpacing(kw.get("spacing", 6))
    for it in items:
        if it is None:
            lay.addStretch(1)
        elif isinstance(it, int):
            lay.addSpacing(it)
        elif isinstance(it, QtWidgets.QLayout):
            lay.addLayout(it)
        else:
            lay.addWidget(it)
    return lay


def vcol(*items, **kw):
    """Pionowy uklad: widgety, uklady, liczby (odstep) albo None (rozpychacz)."""
    lay = QtWidgets.QVBoxLayout()
    lay.setContentsMargins(0, 0, 0, 0)
    lay.setSpacing(kw.get("spacing", 4))
    for it in items:
        if it is None:
            lay.addStretch(1)
        elif isinstance(it, int):
            lay.addSpacing(it)
        elif isinstance(it, QtWidgets.QLayout):
            lay.addLayout(it)
        else:
            lay.addWidget(it)
    return lay


def group(title, layout=None):
    """Ramka z tytulem. Uklad dostaje minimalne marginesy - bez tego
    uklad z zerowymi marginesami (np. hrow) wchodzi pod tytul ramki."""
    g = QtWidgets.QGroupBox(amp(title))
    g.setStyleSheet("QGroupBox{font-weight:bold;margin-top:8px}QGroupBox::title{subcontrol-origin:margin;left:8px;padding:0 3px}")
    if layout is not None:
        m = layout.contentsMargins()
        top = g.fontMetrics().height() // 2 + 8
        layout.setContentsMargins(max(m.left(), 9), max(m.top(), top), max(m.right(), 9), max(m.bottom(), 8))
        g.setLayout(layout)
    return g


def radio_group(parent, *buttons):
    """QButtonGroup dla zestawu przyciskow radiowych - zestawy w jednym
    rodzicu nie wykluczaja sie wtedy nawzajem."""
    bg = QtWidgets.QButtonGroup(parent)
    bg.setExclusive(True)
    for b in buttons:
        bg.addButton(b)
    return bg


def clear_layout(lay):
    """Usuwa widgety z ukladu. hide() od razu, bo deleteLater() zadziala
    dopiero w petli zdarzen - inaczej stare widgety przeswituja spod nowych."""
    while lay.count():
        it = lay.takeAt(0)
        w = it.widget()
        if w is not None:
            w.hide()
            w.deleteLater()
        elif it.layout() is not None:
            clear_layout(it.layout())


def amp(text):
    """Qt traktuje '&' w etykietach przyciskow / kart jako skrot klawiszowy."""
    return text.replace("&", "&&")


def ask_open_file(parent, title, filt, start=""):
    p, _ = QtWidgets.QFileDialog.getOpenFileName(parent, title, start or "", filt)
    return p


def ask_save_file(parent, title, filt, start="", confirm=True):
    opts = QtWidgets.QFileDialog.Options()
    if not confirm:
        opts |= QtWidgets.QFileDialog.DontConfirmOverwrite
    p, _ = QtWidgets.QFileDialog.getSaveFileName(parent, title, start or "", filt, "", opts)
    return p


def ask_dir(parent, title, start=""):
    return QtWidgets.QFileDialog.getExistingDirectory(parent, title, start or "")


def msg_box(parent, title, text, level="info"):
    icon = {"warn": QtWidgets.QMessageBox.Warning, "err": QtWidgets.QMessageBox.Critical}.get(level, QtWidgets.QMessageBox.Information)
    box = QtWidgets.QMessageBox(icon, title, text, QtWidgets.QMessageBox.Ok, parent)
    box.exec_()


def yes_no(parent, title, text):
    r = QtWidgets.QMessageBox.question(parent, title, text, QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                                       QtWidgets.QMessageBox.No)
    return r == QtWidgets.QMessageBox.Yes


def ask_text(parent, title, label, default=""):
    s, ok = QtWidgets.QInputDialog.getText(parent, title, label, QtWidgets.QLineEdit.Normal, default)
    return s if ok else None


class PaletteDialog(_QDialog):
    """Wybor koloru z palety HyperMesha (64 kolory) - kolor na siatce bedzie
    DOKLADNIE taki, jak w legendzie. "Inny kolor..." dobiera najblizszy."""

    def __init__(self, parent, rgb=None, title=None):
        super(PaletteDialog, self).__init__(parent)
        self.setWindowTitle(title or T("Kolor z palety HyperMesha", "Color from the HyperMesh palette"))
        self.result_rgb = None
        pm = palette_map()
        lay = QtWidgets.QVBoxLayout(self)
        lay.addWidget(note_label(T("Komponent w HyperMeshu ma kolor z palety 64 kolor\u00f3w \u2013 wybierz jeden z nich.",
                                   "A HyperMesh component uses a color from the 64-color palette \u2013 pick one.")))
        grid = QtWidgets.QGridLayout()
        grid.setSpacing(3)
        cur = pm.index_of(rgb) if rgb else None
        for i in range(1, 65):
            c = pm.rgb(i)
            b = QtWidgets.QPushButton("%d" % i)
            b.setFixedSize(38, 26)
            fg = "#ffffff" if is_dark(c) else "#000000"
            bd = "3px solid #1c5a96" if i == cur else "1px solid #7f7f7f"
            b.setStyleSheet("QPushButton{background:%s;color:%s;border:%s;font-size:8pt}" % (rgb_hex(c), fg, bd))
            b.setToolTip("%d: RGB %d, %d, %d" % ((i,) + tuple(c)))
            b.clicked.connect(lambda _=False, c=c: self._pick(c))
            grid.addWidget(b, (i - 1) // 8, (i - 1) % 8)
        lay.addLayout(grid)
        other = QtWidgets.QPushButton(T("Inny kolor\u2026 (najbli\u017cszy z palety)", "Other color\u2026 (nearest from the palette)"))
        other.clicked.connect(lambda: self._other(rgb))
        cancel = QtWidgets.QPushButton(T("Anuluj", "Cancel"))
        cancel.clicked.connect(self.reject)
        lay.addLayout(hrow(other, None, cancel))

    def _pick(self, c):
        self.result_rgb = tuple(c)
        self.accept()

    def _other(self, rgb):
        c = QtWidgets.QColorDialog.getColor(qcolor(rgb or (128, 128, 128)), self)
        if c.isValid():
            self._pick(palette_map().snap((c.red(), c.green(), c.blue())))

    @staticmethod
    def ask(parent, rgb):
        d = PaletteDialog(parent, rgb)
        return d.result_rgb if d.exec_() else None


class ColorButton(QtWidgets.QPushButton if _Q else object):
    """Przycisk-probka koloru; klik otwiera palete HM."""

    def __init__(self, rgb, on_change=None, parent=None):
        super(ColorButton, self).__init__(parent)
        self.rgb = tuple(rgb)
        self.on_change = on_change
        self.setFixedSize(40, 20)
        self.clicked.connect(self._click)
        self.refresh()

    def refresh(self):
        self.setStyleSheet("QPushButton{background:%s;border:1px solid #555}" % rgb_hex(self.rgb))
        self.setToolTip(T("Kliknij, aby zmieni\u0107 kolor", "Click to change the color"))

    def set_rgb(self, rgb):
        self.rgb = tuple(rgb)
        self.refresh()

    def _click(self):
        c = PaletteDialog.ask(self, self.rgb)
        if c:
            self.set_rgb(c)
            if self.on_change:
                self.on_change(c)


# ------------------------------------------------------------- legendy
LEGEND_THEMES = {   # tlo, tekst, tekst przyciemniony, obrys, akcent, tlo w SVG
    "dark": ("#41566b", "#ffffff", "#cfe0f5", "#1c2733", "#ffd28a", True),
    "white": ("#ffffff", "#000000", "#444444", "#888888", "#b35a00", True),
    "none": ("#ffffff", "#000000", "#444444", "#888888", "#b35a00", False),
}


class Prims(object):
    """Lista prymitywow legendy w pikselach + pomiar tekstu (QFontMetrics)."""

    def __init__(self, family="Arial"):
        self.items = []
        self.family = family
        self._fm = {}

    def font(self, px, bold=False):
        f = QtGui.QFont(self.family)
        f.setPixelSize(int(px))
        f.setBold(bold)
        return f

    def text_w(self, s, px, bold=False):
        key = (px, bold)
        if key not in self._fm:
            self._fm[key] = QtGui.QFontMetrics(self.font(px, bold))
        return self._fm[key].width(s)

    def rect(self, x, y, w, h, fill, line=""):
        self.items.append(("rect", x, y, w, h, fill, line))

    def text(self, x, y, s, px, color, bold=False, anchor="w"):
        self.items.append(("text", x, y, s, px, color, bold, anchor))

    def poly(self, pts, fill):
        self.items.append(("poly", pts, fill))

    def bbox(self):
        x1 = y1 = 0
        for it in self.items:
            if it[0] == "rect":
                x1, y1 = max(x1, it[1] + it[3]), max(y1, it[2] + it[4])
            elif it[0] == "text":
                w = self.text_w(it[3], it[4], it[6])
                right = it[1] + (w if it[7] == "w" else 0 if it[7] == "e" else w / 2.0)
                x1, y1 = max(x1, right), max(y1, it[2] + it[4] * 0.7)
            else:
                for px, py in it[1]:
                    x1, y1 = max(x1, px), max(y1, py)
        return int(x1) + 12, int(y1) + 10

    def paint(self, p):
        for it in self.items:
            if it[0] == "rect":
                _, x, y, w, h, fill, line = it
                p.setPen(qcolor(line) if line else QtCore.Qt.NoPen)
                p.setBrush(qcolor(fill) if fill else QtCore.Qt.NoBrush)
                p.drawRect(QtCore.QRectF(x, y, w, h))
            elif it[0] == "text":
                _, x, y, s, px, color, bold, anchor = it
                p.setFont(self.font(px, bold))
                p.setPen(qcolor(color))
                w = self.text_w(s, px, bold)
                x0 = x if anchor == "w" else x - w if anchor == "e" else x - w / 2.0
                p.drawText(QtCore.QRectF(x0 - 1, y - px, w + 4, px * 2), QtCore.Qt.AlignVCenter | QtCore.Qt.AlignLeft, s)
            else:
                p.setPen(QtCore.Qt.NoPen)
                p.setBrush(qcolor(it[2]))
                p.drawPolygon(QtGui.QPolygonF([QtCore.QPointF(a, b) for a, b in it[1]]))

    def svg(self, bg=None):
        w, h = self.bbox()
        out = ['<?xml version="1.0" encoding="UTF-8"?>',
               '<svg xmlns="http://www.w3.org/2000/svg" width="%d" height="%d" viewBox="0 0 %d %d">' % (w, h, w, h)]
        if bg:
            out.append('<rect x="0" y="0" width="%d" height="%d" fill="%s"/>' % (w, h, bg))
        for it in self.items:
            if it[0] == "rect":
                _, x, y, rw, rh, fill, line = it
                out.append('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" fill="%s" stroke="%s"/>'
                           % (x, y, rw, rh, fill or "none", line or "none"))
            elif it[0] == "text":
                _, x, y, s, px, color, bold, anchor = it
                ta = {"w": "start", "e": "end"}.get(anchor, "middle")
                out.append('<text x="%.1f" y="%.1f" fill="%s" font-family="Arial" font-size="%dpx" font-weight="%s" text-anchor="%s">%s</text>'
                           % (x, y + px * 0.35, color, px, "bold" if bold else "normal", ta, h_esc(s)))
            else:
                out.append('<polygon points="%s" fill="%s"/>' % (" ".join("%.1f,%.1f" % p for p in it[1]), it[2]))
        out.append("</svg>")
        return "\n".join(out) + "\n"


def mq_legend_prims(v, theme):
    """Legenda widoku "Wiele metryk" (wiersze: kolor, zakres, liczba i %)."""
    bg, fg, fgd, ln, acc, _ = LEGEND_THEMES.get(theme, LEGEND_THEMES["dark"])
    P = Prims()
    title, sub = MQ.legend_head(v)
    y = 18
    P.text(12, y, title, 15, fg, True)
    y += 20
    P.text(12, y, sub, 12, fgd)
    y += 22
    rows = MQ.legend_rows(v)
    rh, sw = 24, 34
    wl = max([P.text_w(r["label"], 12) for r in rows] + [120])
    xl = 14 + sw + 10
    xr = xl + wl + 26 + max([P.text_w("%d  (%.2f%%)" % (r["count"], r["pct"]), 12) for r in rows] + [60])
    for r in rows:
        P.rect(14, y - 9, sw, rh - 6, rgb_hex(r["rgb"]), ln)
        P.text(xl, y + 3, r["label"], 12, fg)
        P.text(xr, y + 3, "%d  (%.2f%%)" % (r["count"], r["pct"]), 12, fg, anchor="e")
        y += rh
    y += 6
    if v != "all":
        s = MQ.stats[v]
        P.text(12, y, T("min: %s (el. %s)   max: %s (el. %s)", "min: %s (el. %s)   max: %s (el. %s)",
                        fmt_num(s["min"], 4), s["minEl"], fmt_num(s["max"], 4), s["maxEl"]), 11, fgd)
        y += 17
        P.text(12, y, T("\u017ar\u00f3d\u0142o: %s", "source: %s", MQ.source_text(v)), 11, fgd)
        y += 17
    P.text(12, y, T("element\u00f3w: %d (2D: %d, 3D: %d)", "elements: %d (2D: %d, 3D: %d)", len(MQ.elems), MQ.n2d, MQ.n3d), 11, fgd)
    return P


def delta_legend_prims(theme):
    """Legenda delty: pionowy pasek gradientu, podzialki, MIN / MAX, wiersze."""
    bg, fg, fgd, ln, acc, _ = LEGEND_THEMES.get(theme, LEGEND_THEMES["dark"])
    lm = DELTA.legend_model()
    P = Prims()
    if not lm:
        return P
    y = 18
    P.text(12, y, lm["title"], 15, fg, True)
    y += 20
    P.text(12, y, lm["sub"], 12, fgd)
    bands = lm["bands"]
    x1, x2 = 104, 142
    top, bot = y + 22, y + 22 + 400
    lo, hi = bands[0]["lo"], bands[-1]["hi"]
    span = (hi - lo) or 1.0
    for b in bands:
        ya = bot - (b["lo"] - lo) / span * (bot - top)
        yb = bot - (b["hi"] - lo) / span * (bot - top)
        P.rect(x1, yb, x2 - x1, ya - yb, rgb_hex(b["rgb"]))
    P.rect(x1, top, x2 - x1, bot - top, "", ln)
    K = len(bands)
    every = max(1, int(math.ceil(K / 9.0)))
    for i in list(range(0, K, every)) + [K]:
        v = bands[i]["lo"] if i < K else hi
        yy = bot - (v - lo) / span * (bot - top)
        P.rect(x2, yy, 6, 0.8, fg)
        P.text(x2 + 10, yy, fmt_num(v, lm["dec"] + 1), 12, fg)
    prev = None
    for tag, info in (("MAX", lm["max"]), ("MIN", lm["min"])):
        if not info:
            continue
        f = max(0.0, min(1.0, (info[0] - lo) / span))
        yy = bot - f * (bot - top)
        if prev is not None and abs(yy - prev) < 28:
            continue
        P.poly([(x1 - 13, yy - 6), (x1 - 13, yy + 6), (x1 - 3, yy)], fg)
        P.text(x1 - 17, yy - 6, tag, 11, fg, True, "e")
        P.text(x1 - 17, yy + 7, fmt_num(info[0], lm["dec"] + 1), 11, fg, True, "e")
        prev = yy
    yy = bot + 22
    for r in lm["extra"]:
        P.rect(x1, yy - 8, x2 - x1, 16, rgb_hex(r["rgb"]), ln)
        P.text(x2 + 10, yy, "%s: %d" % (r["label"], r["count"]), 12, fg)
        yy += 22
    if lm["clamped"]:
        P.text(x1, yy, T("\u25b2 powy\u017cej zakresu: %d el. (kolor g\u00f3rnego pasma)", "\u25b2 above range: %d el. (top band color)", lm["clamped"]), 12, acc)
        yy += 22
    yy += 6
    for tag, info in (("MAX", lm["max"]), ("MIN", lm["min"])):
        if not info:
            continue
        P.text(12, yy, "%s %s = %s   el. %s" % (tag, DELTA.delta_label(), fmt_num(info[0], 4), info[1]), 12, fg, True)
        yy += 17
        xyz = info[2] if len(info) > 2 else None
        if xyz:
            P.text(24, yy, T("\u015brodek: (%g, %g, %g)", "center: (%g, %g, %g)", xyz[0], xyz[1], xyz[2]), 11, fgd)
            yy += 16
    c = DELTA.counts
    P.text(12, yy, T("Elementy w skali: %d", "Elements in scale: %d", c["band"]), 11, fgd)
    return P


class LegendCanvas(_QWidget):
    def __init__(self, parent=None):
        super(LegendCanvas, self).__init__(parent)
        self.prims = None
        self.bg = "#41566b"

    def set_prims(self, prims, bg):
        self.prims, self.bg = prims, bg
        w, h = prims.bbox() if prims else (300, 200)
        self.setFixedSize(max(w, 260), max(h, 120))
        self.update()

    def paintEvent(self, ev):
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing, True)
        p.setRenderHint(QtGui.QPainter.TextAntialiasing, True)
        p.fillRect(self.rect(), qcolor(self.bg))
        if self.prims:
            self.prims.paint(p)
        p.end()


# ------------------------------------------------------------- akcje
class ActionRunner(object):
    """Uruchamia akcje z okna: blokada przyciskow, pasek postepu, przerwanie,
    czytelne bledy; po akcji odswieza karty i podglad slajdu."""

    def __init__(self, owner):
        self.owner = owner
        self.busy = False

    def run(self, fn, done=None, busy=True):
        if self.busy:
            BUS.status(T("Trwa inna operacja \u2013 poczekaj albo kliknij \u201ePrzerwij\u201d.", "Another operation is running \u2013 wait or click \u201cStop\u201d."), "warn")
            return None
        if busy:
            self.busy = True
            self.owner.set_busy(True)
        BUS.cancel = False
        res = None
        try:
            res = fn()
            if done:
                done(res)
        except Cancelled as e:
            BUS.status("%s" % e, "warn")
        except (ValueError, RuntimeError) as e:
            BUS.status("%s" % e, "err")
        except Exception as e:
            BUS.status(T("B\u0141\u0104D: %s (szczeg\u00f3\u0142y w dzienniku)", "ERROR: %s (details in the log)", e), "err")
            BUS.log(traceback.format_exc())
        finally:
            if busy:
                self.busy = False
                self.owner.set_busy(False)
            try:
                self.owner.refresh_all()
            except Exception:
                pass
        return res


# ================= GUI: LEGENDA, EDYTOR LEGENDY, PALETA ================
# Okno legendy (niemodalne, zawsze nad oknem narzedzia): przelacznik
# widokow bez ponownego czytania wartosci, motyw tla (ciemne / biale / bez
# tla), zapis SVG, edytor legendy i slajd PPTX biezacego widoku.
# Edytor legendy: tryb RECZNY (liczba kolorow, granice i kolor kazdego
# pasma, "rozloz rowno", kolory domyslne, wczytanie z automatu) i tryb
# AUTOMATYCZNY (liczba pasm, "ladne" liczby, podglad granic).
class LegendWindow(_QWidget):
    def __init__(self, studio, mode="mq"):
        super(LegendWindow, self).__init__(studio, QtCore.Qt.Tool)
        self.studio = studio
        self.mode = mode           # mq | delta
        self.theme = "dark"
        self.setWindowTitle(T("Legenda", "Legend"))
        lay = QtWidgets.QVBoxLayout(self)
        lay.setContentsMargins(6, 6, 6, 6)
        self.top = QtWidgets.QHBoxLayout()
        lay.addLayout(self.top)
        self.canvas = LegendCanvas()
        lay.addWidget(self.canvas)
        bot = QtWidgets.QHBoxLayout()
        self.theme_rbs = {}
        bot.addWidget(QtWidgets.QLabel(T("T\u0142o:", "Background:")))
        rbs = []
        for key, pl, en in (("dark", "ciemne", "dark"), ("white", "bia\u0142e", "white"), ("none", "bez t\u0142a", "none")):
            rb = QtWidgets.QRadioButton(T(pl, en))
            rb.setChecked(key == self.theme)
            rb.toggled.connect(lambda on, k=key: on and self.set_theme(k))
            self.theme_rbs[key] = rb
            rbs.append(rb)
            bot.addWidget(rb)
        self.theme_group = radio_group(self, *rbs)
        bot.addStretch(1)
        b_svg = QtWidgets.QPushButton(T("Zapisz SVG", "Save SVG"))
        b_svg.clicked.connect(self.save_svg)
        bot.addWidget(b_svg)
        if mode == "mq":
            b_ed = QtWidgets.QPushButton(T("Edytuj legend\u0119\u2026", "Edit legend\u2026"))
            b_ed.clicked.connect(lambda: self.studio.open_editor(MQ.view if MQ.view != "all" else ""))
        else:
            b_ed = QtWidgets.QPushButton(T("R\u0119czna skala\u2026", "Manual scale\u2026"))
            b_ed.clicked.connect(self.studio.tab_delta.edit_scale)
        bot.addWidget(b_ed)
        b_ppt = styled_button(T("Slajd PPTX\u2026", "PPTX slide\u2026"), HDR_COLOR)
        b_ppt.clicked.connect(self.to_pptx)
        bot.addWidget(b_ppt)
        lay.addLayout(bot)
        self.view_btns = {}
        self.refresh()

    def set_theme(self, key):
        self.theme = key
        rb = self.theme_rbs.get(key)
        if rb is not None and not rb.isChecked():
            rb.blockSignals(True)          # wywolanie z kodu: tylko zaznacz
            rb.setChecked(True)
            rb.blockSignals(False)
        self.refresh()

    def refresh(self):
        clear_layout(self.top)
        self.view_btns = {}
        if self.mode == "mq":
            if not MQ.view:
                self.hide()
                return
            self.top.addWidget(QtWidgets.QLabel(T("Poka\u017c:", "Show:")))
            for v in MQ.views():
                b = QtWidgets.QPushButton(MQ.tag(v))
                b.setCheckable(True)
                b.setChecked(v == MQ.view)
                b.clicked.connect(lambda _=False, v=v: self.studio.tab_mq.show_view(v))
                self.top.addWidget(b)
                self.view_btns[v] = b
            self.top.addStretch(1)
            prims = mq_legend_prims(MQ.view, self.theme)
            self.setWindowTitle("%s \u2013 %s" % (T("Legenda", "Legend"), MQ.label(MQ.view)))
        else:
            if not DELTA.done:
                self.hide()
                return
            prims = delta_legend_prims(self.theme)
            self.setWindowTitle("%s \u2013 %s" % (T("Legenda", "Legend"), DELTA.delta_label()))
        self.canvas.set_prims(prims, LEGEND_THEMES[self.theme][0])
        self.adjustSize()

    def save_svg(self):
        prims = mq_legend_prims(MQ.view, self.theme) if self.mode == "mq" else delta_legend_prims(self.theme)
        th = LEGEND_THEMES[self.theme]
        tag = MQ.tag(MQ.view) if self.mode == "mq" else DELTA.delta_tag()
        base = "legenda_%s_%s.svg" % (tag, now_text("%Y%m%d-%H%M%S"))
        d = HM.model_dir()
        path = os.path.join(d, base) if d and writable_dir(d) else ask_save_file(self, T("Zapisz legend\u0119", "Save legend"), "SVG (*.svg)", base)
        if not path:
            return
        try:
            write_text(path, prims.svg(th[0] if th[5] else None))
            BUS.status(T("Zapisano legend\u0119: %s", "Legend saved: %s", path))
        except Exception as e:
            BUS.status(T("Nie zapisano SVG: %s", "SVG not saved: %s", e), "err")

    def to_pptx(self):
        if self.mode == "mq":
            self.studio.tab_ppt.current_to_pptx()
        else:
            self.studio.tab_ppt.delta_to_pptx()


class LegendEditor(_QDialog):
    def __init__(self, studio, m):
        super(LegendEditor, self).__init__(studio)
        self.studio = studio
        self.setWindowTitle(T("Legenda i przedzia\u0142y kolor\u00f3w", "Legend and color bands"))
        self.setModal(False)
        self.lay = QtWidgets.QVBoxLayout(self)
        self.body = None
        self.load(m)

    # ---------------------------------------------------------- stan edytora
    def load(self, m):
        c = MQ.cfg[m]
        self.m = m
        self.thr_on = bool(c["thr_on"])
        self.thr = "%s" % fmt_num(c["thr"], 6)
        self.mode = c["mode"]
        self.nb = int(c["nb"])
        E = manual_edges(c["edges"], self.thr_on, to_float(self.thr), MQ_DEF[m].dir)
        if len(E) < 2:
            E = self.preview_auto()
        self.E = [float(x) for x in E]
        k = len(self.E) - 1
        self.C = [tuple(x) for x in c["colors"]] if len(c["colors"]) == k else default_band_colors(MQ_DEF[m].dir, k, self.thr_on)
        self.far = self.E[-1] if MQ_DEF[m].dir == "above" else self.E[0]
        self.build()

    def direction(self):
        return MQ_DEF[self.m].dir

    def preview_auto(self):
        st = MQ.stats.get(self.m) or {}
        thr = to_float(self.thr, 0.0)
        vmin, vmax = st.get("min"), st.get("max")
        if vmin is None:
            span = abs(thr) if thr else 1.0
            vmin, vmax = (thr, thr + span) if self.direction() == "above" else (thr - span, thr)
        return auto_edges(self.direction(), self.thr_on, thr if self.thr_on else None, self.nb, vmin, vmax, MQ.nice)

    # ---------------------------------------------------------- budowa okna
    def build(self):
        if self.body is not None:
            self.lay.removeWidget(self.body)
            self.body.hide()               # deleteLater dziala dopiero w petli zdarzen
            self.body.deleteLater()
        self.body = QtWidgets.QWidget()
        self.lay.addWidget(self.body)
        v = QtWidgets.QVBoxLayout(self.body)
        v.setContentsMargins(0, 0, 0, 0)
        hdr = QtWidgets.QLabel(T("Legenda i przedzia\u0142y kolor\u00f3w", "Legend and color bands"))
        hdr.setStyleSheet("background:%s;color:white;font-weight:bold;padding:6px 10px" % HDR_COLOR)
        v.addWidget(hdr)
        row = QtWidgets.QHBoxLayout()
        row.addWidget(QtWidgets.QLabel(T("Metryka:", "Metric:")))
        mrbs = []
        for mm in MQ_ORDER:
            rb = QtWidgets.QRadioButton(MQ_DEF[mm].label)
            rb.setChecked(mm == self.m)
            rb.toggled.connect(lambda on, mm=mm: on and mm != self.m and self.load(mm))
            row.addWidget(rb)
            mrbs.append(rb)
        self.metric_group = radio_group(self.body, *mrbs)    # osobna grupa: nie koliduje z trybem legendy
        row.addStretch(1)
        v.addLayout(row)
        self.cb_thr = QtWidgets.QCheckBox(T("Podzia\u0142 wg progu \u2013 w normie:", "Split by threshold \u2013 within limits:"))
        self.cb_thr.setChecked(self.thr_on)
        self.cb_thr.toggled.connect(lambda on: self.rebuild_keep())
        self.e_thr = QtWidgets.QLineEdit(self.thr)
        self.e_thr.setMaximumWidth(90)
        self.e_thr.editingFinished.connect(self.rebuild_keep)
        strict = MQ_DEF[self.m].strict
        op = QtWidgets.QLabel(("<" if strict else "\u2264") if self.direction() == "above" else (">" if strict else "\u2265"))
        op.setStyleSheet("font-weight:bold")
        hint = note_label(T("(wy\u017csze = gorsze)", "(higher = worse)") if self.direction() == "above" else T("(ni\u017csze = gorsze)", "(lower = worse)"))
        v.addLayout(hrow(self.cb_thr, op, self.e_thr, hint, None))
        self.rb_auto = QtWidgets.QRadioButton(T("automatyczny (z warto\u015bci skrajnych)", "automatic (from extreme values)"))
        self.rb_man = QtWidgets.QRadioButton(T("r\u0119czny (pe\u0142na kontrola)", "manual (full control)"))
        self.mode_group = radio_group(self.body, self.rb_auto, self.rb_man)
        (self.rb_auto if self.mode == "auto" else self.rb_man).setChecked(True)
        self.rb_auto.toggled.connect(lambda on: on and self.set_mode("auto"))
        self.rb_man.toggled.connect(lambda on: on and self.set_mode("manual"))
        v.addLayout(hrow(QtWidgets.QLabel(T("Tryb legendy:", "Legend mode:")), self.rb_auto, self.rb_man, None))
        st = MQ.stats.get(self.m)
        if st and st.get("min") is not None:
            lab = QtWidgets.QLabel(T("Dane z ostatniej analizy: min %s, max %s", "Data from the last analysis: min %s, max %s",
                                     fmt_num(st["min"], 4), fmt_num(st["max"], 4)))
            lab.setStyleSheet("color:%s" % HDR_COLOR)
            v.addWidget(lab)
        box = QtWidgets.QFrame()
        box.setFrameShape(QtWidgets.QFrame.StyledPanel)
        g = QtWidgets.QGridLayout(box)
        if self.mode == "auto":
            self.sp_nb = QtWidgets.QSpinBox()
            self.sp_nb.setRange(1, 40)
            self.sp_nb.setValue(self.nb)
            self.cb_nice = QtWidgets.QCheckBox(T("\u0142adne liczby", "nice numbers"))
            self.cb_nice.setChecked(MQ.nice)
            b_prev = QtWidgets.QPushButton(T("Przelicz podgl\u0105d", "Recompute preview"))
            b_prev.clicked.connect(self.rebuild_keep)
            b_man = QtWidgets.QPushButton(T("Przejd\u017a do r\u0119cznego z tymi pasmami", "Switch to manual with these bands"))
            b_man.clicked.connect(self.auto_to_manual)
            g.addLayout(hrow(QtWidgets.QLabel(T("Liczba pasm:", "Number of bands:")), self.sp_nb, self.cb_nice, b_prev, None), 0, 0)
            g.addLayout(hrow(b_man, None), 1, 0)
            E = self.preview_auto()
            C = default_band_colors(self.direction(), len(E) - 1, self.thr_on)
            editable = False
        else:
            self.sp_k = QtWidgets.QSpinBox()
            self.sp_k.setRange(1, 40)
            self.sp_k.setValue(len(self.E) - 1)
            b_k = QtWidgets.QPushButton(T("Ustaw", "Set"))
            b_k.clicked.connect(self.set_k)
            self.e_far = QtWidgets.QLineEdit(fmt_num(self.far, 6))
            self.e_far.setMaximumWidth(80)
            b_sp = QtWidgets.QPushButton(T("Roz\u0142\u00f3\u017c", "Spread"))
            b_sp.clicked.connect(self.spread)
            b_dc = QtWidgets.QPushButton(T("Kolory domy\u015blne", "Default colors"))
            b_dc.clicked.connect(self.default_colors)
            b_fa = QtWidgets.QPushButton(T("Wczytaj z automatu", "Load from automatic"))
            b_fa.clicked.connect(self.from_auto)
            g.addLayout(hrow(QtWidgets.QLabel(T("Liczba kolor\u00f3w (pasm):", "Number of colors (bands):")), self.sp_k, b_k, 12,
                             QtWidgets.QLabel(T("Rozk\u0142ad r\u00f3wno do:", "Spread evenly to:")), self.e_far, b_sp, None), 0, 0)
            g.addLayout(hrow(b_dc, b_fa, None), 1, 0)
            E, C = self.E, self.C
            editable = True
        v.addWidget(box)
        # siatka granic i kolorow - najwieksze wartosci u gory
        self.edge_edits = {}
        inner = QtWidgets.QWidget()
        gl = QtWidgets.QGridLayout(inner)
        gl.setVerticalSpacing(2)
        K = len(E) - 1
        lock = (0 if self.direction() == "above" else K) if self.thr_on else -1
        r = 0
        for i in range(K, -1, -1):
            lab = QtWidgets.QLabel(T("pr\u00f3g", "threshold") if i == lock else T("granica", "edge"))
            lab.setStyleSheet("color:#555")
            ed = QtWidgets.QLineEdit(fmt_num(E[i], 6))
            ed.setMaximumWidth(100)
            ed.setAlignment(QtCore.Qt.AlignRight)
            if not editable or i == lock:
                ed.setReadOnly(True)
                ed.setStyleSheet("background:#eeeeee")
            self.edge_edits[i] = ed
            gl.addWidget(lab, r, 0, QtCore.Qt.AlignRight)
            gl.addWidget(ed, r, 1)
            r += 1
            if i > 0:
                b = i - 1
                cb = ColorButton(C[b], (lambda c, b=b: self.set_color(b, c)) if editable else None)
                cb.setEnabled(editable)
                rank = b + 1 if self.direction() == "above" else K - b
                extra = ""
                if rank == K:
                    extra = " (%s)" % T("najgorsze", "worst")
                elif rank == 1 and self.thr_on:
                    extra = " (%s)" % T("najbli\u017cej normy", "closest to limit")
                gl.addWidget(cb, r, 1, QtCore.Qt.AlignCenter)
                gl.addWidget(QtWidgets.QLabel("%s %d%s" % (T("pasmo", "band"), rank, extra)), r, 2)
                r += 1
        sa = QtWidgets.QScrollArea()
        sa.setWidget(inner)
        sa.setWidgetResizable(True)
        scr = screen_rect(self)
        sa.setMinimumHeight(min(inner.sizeHint().height() + 8, int(scr.height() * 0.4)))
        v.addWidget(sa)
        self.c_ok = ColorButton(MQ.col_ok, lambda c: setattr(MQ, "col_ok", c))
        self.c_over = ColorButton(MQ.col_over, lambda c: setattr(MQ, "col_over", c))
        self.c_na = ColorButton(MQ.col_na, lambda c: setattr(MQ, "col_na", c))
        self.cb_over = QtWidgets.QCheckBox(T("poza skal\u0105 = osobna grupa", "beyond scale = separate group"))
        self.cb_over.setChecked(MQ.over_sep)
        v.addLayout(hrow(QtWidgets.QLabel(T("w normie", "within limits")), self.c_ok, QtWidgets.QLabel(T("poza skal\u0105", "beyond scale")),
                         self.c_over, QtWidgets.QLabel(T("n/d", "n/a")), self.c_na, 10, self.cb_over, None))
        v.addWidget(note_label(T("Granice rosn\u0105 od do\u0142u do g\u00f3ry. Przy podziale wg progu granica po stronie normy jest r\u00f3wna progowi. "
                                 "Kliknij kolor, aby wybra\u0107 go z palety HyperMesha.",
                                 "Edges grow from bottom to top. With the threshold split, the edge on the within-limits side equals "
                                 "the threshold. Click a color to pick it from the HyperMesh palette.")))
        b_apply_rc = styled_button(T("Zastosuj i przekoloruj", "Apply and recolor"), RUN_COLOR)
        b_apply_rc.clicked.connect(lambda: self.apply(True))
        b_apply = QtWidgets.QPushButton(T("Zastosuj", "Apply"))
        b_apply.clicked.connect(lambda: self.apply(False))
        b_close = QtWidgets.QPushButton(T("Zamknij", "Close"))
        b_close.clicked.connect(self.close)
        v.addLayout(hrow(b_apply_rc, b_apply, None, b_close))
        self.adjustSize()

    # ---------------------------------------------------------- operacje
    def read_edges(self):
        return [to_float(self.edge_edits[i].text()) for i in sorted(self.edge_edits)]

    def read_state(self):
        self.thr_on = self.cb_thr.isChecked()
        self.thr = self.e_thr.text().strip()
        if self.mode == "auto":
            self.nb = self.sp_nb.value()
            MQ.nice = self.cb_nice.isChecked()
        else:
            E = self.read_edges()
            if all(e is not None for e in E):
                self.E = E
        MQ.over_sep = self.cb_over.isChecked()

    def sync_thr(self):
        T0 = to_float(self.thr)
        if not self.thr_on or T0 is None or self.mode != "manual":
            return
        E = self.E
        if self.direction() == "above":
            F = [T0] + [e for e in E[1:] if e > T0]
        else:
            F = [e for e in E[:-1] if e < T0] + [T0]
        if len(F) < 2:
            step = abs(T0) * 0.25 if T0 else 0.25
            F = [T0, T0 + step] if self.direction() == "above" else [T0 - step, T0]
        if len(self.C) != len(F) - 1:
            self.C = default_band_colors(self.direction(), len(F) - 1, True)
        self.E = F

    def rebuild_keep(self):
        self.read_state()
        self.sync_thr()
        self.build()

    def set_mode(self, mode):
        self.read_state()
        self.mode = mode
        self.sync_thr()
        self.build()

    def set_k(self):
        self.read_state()
        K = self.sp_k.value()
        lo, hi = self.E[0], self.E[-1]
        if hi <= lo:
            hi = lo + 1.0
        self.E = [clean10(lo + (hi - lo) * i / float(K)) for i in range(K + 1)]
        self.C = default_band_colors(self.direction(), K, self.thr_on)
        self.build()

    def spread(self):
        self.read_state()
        far = to_float(self.e_far.text())
        if far is None:
            return
        K = len(self.E) - 1
        lo, hi = (self.E[0], far) if self.direction() == "above" else (far, self.E[-1])
        if hi <= lo:
            BUS.status(T("Zakres do roz\u0142o\u017cenia musi by\u0107 rosn\u0105cy.", "The range to spread must be increasing."), "err")
            return
        self.far = far
        self.E = [clean10(lo + (hi - lo) * i / float(K)) for i in range(K + 1)]
        self.build()

    def default_colors(self):
        self.read_state()
        self.C = default_band_colors(self.direction(), len(self.E) - 1, self.thr_on)
        self.build()

    def from_auto(self):
        self.read_state()
        self.E = self.preview_auto()
        self.C = default_band_colors(self.direction(), len(self.E) - 1, self.thr_on)
        self.build()

    def auto_to_manual(self):
        self.read_state()
        self.E = self.preview_auto()
        self.C = default_band_colors(self.direction(), len(self.E) - 1, self.thr_on)
        self.mode = "manual"
        self.build()

    def set_color(self, b, c):
        self.C[b] = tuple(c)

    def apply(self, recolor):
        self.read_state()
        m = self.m
        if self.thr_on and to_float(self.thr) is None:
            BUS.status(T("Pr\u00f3g musi by\u0107 liczb\u0105.", "The threshold must be a number."), "err")
            return
        c = MQ.cfg[m]
        if self.mode == "manual":
            E = self.read_edges()
            for i, e in enumerate(E):
                if e is None:
                    BUS.status(T("Granica nr %d nie jest liczb\u0105.", "Edge #%d is not a number.", i + 1), "err")
                    return
                if i and e <= E[i - 1]:
                    BUS.status(T("Granice musz\u0105 rosn\u0105\u0107 (problem przy granicy nr %d).", "Edges must increase (problem at edge #%d).", i + 1), "err")
                    return
            c["edges"] = E
            c["colors"] = list(self.C)
        else:
            c["nb"] = self.nb
        c["thr_on"] = self.thr_on
        if to_float(self.thr) is not None:
            c["thr"] = to_float(self.thr)
        c["mode"] = self.mode
        save_settings()
        self.studio.tab_mq.load_from_engine()
        if m in MQ.analyzed:
            MQ.recompute()
            if recolor:
                def work():
                    if MQ.make_sets:
                        MQ.build_sets()
                    MQ.apply_view(m if MQ.view != "all" else "all")
                self.studio.runner.run(work)
            else:
                BUS.status(T("Zapisano legend\u0119 %s (siatka bez zmian \u2013 u\u017cyj \u201eZastosuj i przekoloruj\u201d).",
                             "Legend %s saved (mesh unchanged \u2013 use \u201cApply and recolor\u201d).", MQ.label(m)))
            self.studio.refresh_all()
        else:
            BUS.status(T("Zapisano legend\u0119 %s \u2013 zadzia\u0142a przy nast\u0119pnej analizie.", "Legend %s saved \u2013 used by the next analysis.", MQ.label(m)))
        self.load(m)


class DeltaScaleEditor(_QDialog):
    """Reczna skala delty: dolny prog kazdego pasma + gorna granica."""

    def __init__(self, parent):
        super(DeltaScaleEditor, self).__init__(parent)
        self.setWindowTitle(T("R\u0119czna skala \u2013 progi pasm", "Manual scale \u2013 band limits"))
        K = max(2, min(96, int(DELTA.band_count or 10)))
        bb = list(DELTA.bounds) if len(DELTA.bounds) == K + 1 else (
            list(DELTA.manual_bounds) if len(DELTA.manual_bounds) == K + 1 else None)
        if bb is None:
            lo = max(0.0, to_float(DELTA.deadband, 0.01))
            hi = lo + 0.35
            bb = [lo + i * (hi - lo) / K for i in range(K + 1)]
        v = QtWidgets.QVBoxLayout(self)
        hdr = QtWidgets.QLabel(T("Dolny pr\u00f3g ka\u017cdego pasma + g\u00f3rna granica", "Lower limit of each band + upper limit"))
        hdr.setStyleSheet("background:#2b579a;color:white;font-weight:bold;padding:6px 10px")
        v.addWidget(hdr)
        inner = QtWidgets.QWidget()
        g = QtWidgets.QGridLayout(inner)
        for c, h in enumerate((T("Pasmo", "Band"), T("Kolor", "Color"), T("od", "from"))):
            lab = QtWidgets.QLabel(h)
            lab.setStyleSheet("font-weight:bold")
            g.addWidget(lab, 0, c)
        g.addWidget(QtWidgets.QLabel(T("g\u00f3ra", "top")), 1, 0)
        g.addWidget(QtWidgets.QLabel(T("do", "to")), 1, 1)
        self.e_hi = QtWidgets.QLineEdit("%.6g" % bb[-1])
        g.addWidget(self.e_hi, 1, 2)
        cols = DELTA.band_colors(K)
        self.edits = [None] * K
        for r, i in enumerate(range(K - 1, -1, -1), start=2):
            g.addWidget(QtWidgets.QLabel("%d" % (i + 1)), r, 0)
            sw = QtWidgets.QLabel()
            sw.setFixedSize(34, 16)
            sw.setStyleSheet("background:%s;border:1px solid #555" % rgb_hex(cols[i]))
            g.addWidget(sw, r, 1)
            e = QtWidgets.QLineEdit("%.6g" % bb[i])
            g.addWidget(e, r, 2)
            self.edits[i] = e
        sa = QtWidgets.QScrollArea()
        sa.setWidget(inner)
        sa.setWidgetResizable(True)
        sa.setMinimumHeight(min(inner.sizeHint().height() + 8, 420))
        v.addWidget(sa)
        v.addWidget(note_label(T("Pasma przylegaj\u0105 do siebie: g\u00f3rna granica pasma = dolny pr\u00f3g nast\u0119pnego. Warto\u015bci musz\u0105 rosn\u0105\u0107.",
                                 "Bands are adjacent: the top of a band = the lower limit of the next one. Values must increase.")))
        ok = styled_button(T("Zastosuj", "Apply"), RUN_COLOR)
        ok.clicked.connect(self.apply)
        cancel = QtWidgets.QPushButton(T("Anuluj", "Cancel"))
        cancel.clicked.connect(self.reject)
        v.addLayout(hrow(ok, None, cancel))

    def apply(self):
        vals = [to_float(e.text()) for e in self.edits] + [to_float(self.e_hi.text())]
        for i, x in enumerate(vals):
            if x is None:
                BUS.status(T("Warto\u015b\u0107 nr %d nie jest liczb\u0105.", "Value #%d is not a number.", i + 1), "err")
                return
            if i and x <= vals[i - 1]:
                BUS.status(T("Warto\u015bci musz\u0105 rosn\u0105\u0107 (pasmo %d).", "Values must increase (band %d).", i), "err")
                return
        DELTA.manual_bounds = vals
        DELTA.band_count = len(vals) - 1
        DELTA.auto_scale = False
        BUS.status(T("Zapisano r\u0119czn\u0105 skal\u0119 (%d pasm) \u2013 zadzia\u0142a przy nast\u0119pnej analizie.",
                     "Manual scale saved (%d bands) \u2013 used by the next analysis.", len(vals) - 1))
        self.accept()


# ========================= GUI: PODGLAD SLAJDU =========================
# Jak w makrze ANSYS SHOTS - podglad jest w CZTERECH miejscach i wszedzie
# rysuje te same prymitywy, ktore trafia do pliku .pptx (slide_ops):
#  1. INTERAKTYWNE plotno na karcie "Prezentacja PPTX": uklad "wlasny kadr"
#     pozwala przeciagnac obraz myszka (polozenie) i zmienic jego rozmiar
#     (uchwyty w rogach); szybkie dopasowania (1/2 lewa, pelny...), zapisane
#     kadry pod wlasna nazwa; przeciagniecie obrazu przy innym ukladzie
#     przelacza uklad na "wlasny kadr",
#  2. karta "Podglad slajdu" (tylko do odczytu): przykladowy slajd z
#     biezacego stanu + "CO POWSTANIE" (plik, tryb, seria, delta, raport),
#     odswiezana przy wejsciu na karte i po kazdej zmianie opcji,
#  3. okno "Podglad na zywo" (F6): niemodalne, mozna je postawic obok okna
#     narzedzia; odswieza sie SAMO po kazdej zmianie opcji (LIVE_DELAY_MS),
#  4. PODGLAD WSZYSTKICH SLAJDOW przed zapisem (Utworz / Anuluj) - zmiana
#     tytulow, ukladu, usuwanie slajdow - i GALERIA wyeksportowanych slajdow
#     po zapisie (z otwarciem prezentacji / folderu).
LIVE_DELAY_MS = 250           # odswiezenie podgladu po ostatniej zmianie opcji
PREVIEW_MARGIN = 10           # margines plotna wokol karty slajdu [px]
HANDLE = 9                    # uchwyt zmiany rozmiaru kadru [px]


def paint_slide(p, ops, s, img_cache, ox=0.0, oy=0.0):
    """Rysuje prymitywy slajdu na QPainter w skali s (px / EMU) z przesunieciem."""
    for op in ops:
        kind = op[0]
        X, Y, W, H = ox + op[1] * s, oy + op[2] * s, op[3] * s, op[4] * s
        r = QtCore.QRectF(X, Y, W, H)
        if kind == "rect":
            p.setPen(qcolor(op[6]) if op[6] else QtCore.Qt.NoPen)
            p.setBrush(qcolor(op[5]) if op[5] else QtCore.Qt.NoBrush)
            p.drawRect(r)
        elif kind == "text":
            txt, size, bold, color, align, anchor = op[5:11]
            f = QtGui.QFont("Arial")
            f.setPixelSize(max(4, int(round(size * 12700 * s))))
            f.setBold(bool(bold))
            p.setFont(f)
            p.setPen(qcolor(color))
            fl = {"l": QtCore.Qt.AlignLeft, "r": QtCore.Qt.AlignRight}.get(align, QtCore.Qt.AlignHCenter)
            fl |= {"t": QtCore.Qt.AlignTop, "b": QtCore.Qt.AlignBottom}.get(anchor, QtCore.Qt.AlignVCenter)
            p.drawText(r.adjusted(2, 0, -2, 0), int(fl) | QtCore.Qt.TextWordWrap, txt)
        elif kind == "image":
            img = img_cache.get(op[5])
            if img is None:
                img = QtGui.QImage(op[5])
                img_cache[op[5]] = img
            p.fillRect(r, qcolor("#e8ecf0"))
            if not img.isNull():
                p.drawImage(r, img)
        elif kind == "table":
            rows, fr, size = op[5], op[6], op[7]
            nr, nc = len(rows), max(len(x) for x in rows)
            fr = list(fr) if fr and len(fr) == nc else [1.0] * nc
            tot = float(sum(fr))
            rh = H / float(nr)
            f = QtGui.QFont("Arial")
            f.setPixelSize(max(4, int(round(size * 12700 * s))))
            for ri, row in enumerate(rows):
                cx = X
                for ci in range(nc):
                    cw = W * fr[ci] / tot
                    d = cell_of(row[ci] if ci < len(row) else "")
                    fill = "#1C5A96" if ri == 0 else ("#" + d["fill"] if d["fill"] else ("#F4F7FA" if ri % 2 else "#FFFFFF"))
                    cr = QtCore.QRectF(cx, Y + ri * rh, cw, rh)
                    p.setPen(qcolor("#BFBFBF"))
                    p.setBrush(qcolor(fill))
                    p.drawRect(cr)
                    f.setBold(ri == 0 or d["bold"])
                    p.setFont(f)
                    p.setPen(qcolor("#FFFFFF" if ri == 0 else "#" + (d["color"] or "262626")))
                    al = "ctr" if ri == 0 else (d["align"] or ("l" if ci == 0 else "ctr"))
                    fl = {"l": QtCore.Qt.AlignLeft, "r": QtCore.Qt.AlignRight}.get(al, QtCore.Qt.AlignHCenter) | QtCore.Qt.AlignVCenter
                    p.drawText(cr.adjusted(3, 0, -3, 0), int(fl), d["t"])
                    cx += cw


def render_slide_image(ops, sw, sh, width=1280):
    """Obraz slajdu (QImage) z prymitywow - do galerii po eksporcie."""
    w = int(width)
    h = int(round(w * sh / float(sw)))
    img = QtGui.QImage(w, h, QtGui.QImage.Format_RGB32)
    img.fill(QtCore.Qt.white)
    p = QtGui.QPainter(img)
    p.setRenderHint(QtGui.QPainter.Antialiasing, True)
    p.setRenderHint(QtGui.QPainter.SmoothPixmapTransform, True)
    p.setRenderHint(QtGui.QPainter.TextAntialiasing, True)
    paint_slide(p, ops, w / float(sw), {})
    p.end()
    return img


def render_items_png(items, width=1280):
    """Pliki PNG wszystkich slajdow (folder roboczy) - galeria po eksporcie."""
    sw, sh = PRESENT.slide_size()
    out = []
    d = work_dir("exports")
    stamp = now_text("%Y%m%d_%H%M%S")
    for i, it in enumerate(items, 1):
        try:
            img = render_slide_image(slide_ops(it, sw, sh, PRESENT.style), sw, sh, width)
            path = os.path.join(d, "slide_%s_%02d.png" % (stamp, i))
            if img.save(path, "PNG"):
                out.append(path)
        except Exception as e:
            BUS.log("render slide %d: %s" % (i, e))
    return out


class SlideCanvas(_QWidget):
    """Plotno slajdu (proporcje prezentacji, karta ze cieniem). W trybie
    interaktywnym pozwala przeciagac i skalowac kadr obrazu (uklad "wlasny
    kadr"); przeciagniecie obrazu przy innym ukladzie przelacza na wlasny."""

    def __init__(self, parent, sw, sh, min_w=360, interactive=False):
        super(SlideCanvas, self).__init__(parent)
        self.sw, self.sh = sw, sh
        self.ops = []
        self.cache = {}
        self.interactive = interactive
        self.rebuild_cb = None        # f() -> ops (po zmianie kadru)
        self.changed_cb = None        # f() po zakonczeniu przeciagania
        self.badge = ""
        self.drag = None
        self.setMinimumSize(int(min_w), int(min_w * sh / float(sw)))
        self.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.setMouseTracking(True)

    def set_slide_size(self, sw, sh):
        self.sw, self.sh = sw, sh
        self.update()

    def set_ops(self, ops):
        self.ops = ops or []
        self.update()

    def clear_cache(self):
        self.cache = {}

    # ---------------------------------------------------------- geometria
    def bounds(self):
        """(x, y, w, h) karty slajdu w pikselach plotna (wpisanej z marginesem)."""
        W, H = self.width(), self.height()
        aw, ah = max(40, W - 2 * PREVIEW_MARGIN), max(30, H - 2 * PREVIEW_MARGIN)
        s = min(aw / float(self.sw), ah / float(self.sh))
        w, h = self.sw * s, self.sh * s
        return PREVIEW_MARGIN + (aw - w) / 2.0, PREVIEW_MARGIN + (ah - h) / 2.0, w, h

    def box_px(self):
        sx, sy, sw, sh = self.bounds()
        b = PRESENT.style.box
        return QtCore.QRectF(sx + b[0] * sw, sy + b[1] * sh, b[2] * sw, b[3] * sh)

    def image_px(self):
        """Prostokat obrazu z prymitywow (dowolny uklad) - do przelaczenia na wlasny kadr."""
        sx, sy, sw, sh = self.bounds()
        s = sw / float(self.sw)
        for op in self.ops:
            if op[0] == "image" or (op[0] == "rect" and op[5] == "F2F2F2"):
                return QtCore.QRectF(sx + op[1] * s, sy + op[2] * s, op[3] * s, op[4] * s)
        return None

    # ---------------------------------------------------------- rysowanie
    def paintEvent(self, ev):
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing, True)
        p.setRenderHint(QtGui.QPainter.SmoothPixmapTransform, True)
        p.setRenderHint(QtGui.QPainter.TextAntialiasing, True)
        p.fillRect(self.rect(), qcolor("#eef2f6"))
        sx, sy, sw, sh = self.bounds()
        p.setPen(QtCore.Qt.NoPen)
        p.setBrush(qcolor("#d7dce4"))
        p.drawRect(QtCore.QRectF(sx + 4, sy + 4, sw, sh))
        p.setBrush(QtCore.Qt.white)
        p.drawRect(QtCore.QRectF(sx, sy, sw, sh))
        p.setClipRect(QtCore.QRectF(sx, sy, sw, sh))
        paint_slide(p, self.ops, sw / float(self.sw), self.cache, sx, sy)
        p.setClipping(False)
        p.setPen(qcolor("#9aa6b2"))
        p.setBrush(QtCore.Qt.NoBrush)
        p.drawRect(QtCore.QRectF(sx, sy, sw, sh))
        if self.interactive and PRESENT.style.layout == "custom":
            r = self.box_px()
            pen = QtGui.QPen(qcolor("#1c5a96"), 1.5, QtCore.Qt.DashLine)
            p.setPen(pen)
            p.drawRect(r)
            p.setPen(qcolor("#1c5a96"))
            p.setBrush(QtCore.Qt.white)
            for c in self._handles(r):
                p.drawRect(c)
        if self.badge:
            f = QtGui.QFont("Arial")
            f.setPixelSize(11)
            f.setBold(True)
            p.setFont(f)
            fm = QtGui.QFontMetrics(f)
            tw = fm.width(self.badge) + 12
            p.setPen(QtCore.Qt.NoPen)
            p.setBrush(qcolor("#f2f5f9"))
            p.drawRect(QtCore.QRectF(self.width() - tw - 6, 4, tw, 18))
            p.setPen(qcolor("#6e7682"))
            p.drawText(QtCore.QRectF(self.width() - tw - 6, 4, tw, 18), int(QtCore.Qt.AlignCenter), self.badge)
        p.end()

    @staticmethod
    def _handles(r):
        h = HANDLE
        return [QtCore.QRectF(r.left() - h / 2.0, r.top() - h / 2.0, h, h), QtCore.QRectF(r.right() - h / 2.0, r.top() - h / 2.0, h, h),
                QtCore.QRectF(r.left() - h / 2.0, r.bottom() - h / 2.0, h, h), QtCore.QRectF(r.right() - h / 2.0, r.bottom() - h / 2.0, h, h)]

    # ---------------------------------------------------------- mysz
    def _hit(self, pos):
        if not self.interactive:
            return None
        if PRESENT.style.layout == "custom":
            r = self.box_px()
            for i, c in enumerate(self._handles(r)):
                if c.adjusted(-3, -3, 3, 3).contains(pos):
                    return "h%d" % i
            if r.contains(pos):
                return "move"
            return None
        r = self.image_px()
        return "switch" if (r is not None and r.contains(pos)) else None

    def mousePressEvent(self, ev):
        if ev.button() != QtCore.Qt.LeftButton:
            return
        hit = self._hit(ev.pos())
        if hit is None:
            return
        if hit == "switch":
            # inny uklad: przeciagniecie obrazu = przejscie na wlasny kadr od jego biezacego prostokata
            r = self.image_px()
            sx, sy, sw, sh = self.bounds()
            PRESENT.style.layout = "custom"
            PRESENT.style.set_box((r.left() - sx) / sw, (r.top() - sy) / sh, r.width() / sw, r.height() / sh)
            hit = "move"
            self._rebuild()
        self.drag = (hit, ev.pos(), list(PRESENT.style.box))
        self.setCursor(QtCore.Qt.ClosedHandCursor if hit == "move" else QtCore.Qt.SizeFDiagCursor)

    def mouseMoveEvent(self, ev):
        if self.drag is None:
            hit = self._hit(ev.pos())
            if hit in ("move", "switch"):
                self.setCursor(QtCore.Qt.OpenHandCursor)
            elif hit in ("h0", "h3"):
                self.setCursor(QtCore.Qt.SizeFDiagCursor)
            elif hit in ("h1", "h2"):
                self.setCursor(QtCore.Qt.SizeBDiagCursor)
            else:
                self.setCursor(QtCore.Qt.ArrowCursor)
            return
        mode, p0, b0 = self.drag
        sx, sy, sw, sh = self.bounds()
        dx, dy = (ev.pos().x() - p0.x()) / sw, (ev.pos().y() - p0.y()) / sh
        x, y, w, h = b0
        if mode == "move":
            x, y = x + dx, y + dy
        elif mode == "h3":
            w, h = w + dx, h + dy
        elif mode == "h0":
            x, y, w, h = x + dx, y + dy, w - dx, h - dy
        elif mode == "h1":
            y, w, h = y + dy, w + dx, h - dy
        elif mode == "h2":
            x, w, h = x + dx, w - dx, h + dy
        if ev.modifiers() & QtCore.Qt.ShiftModifier and mode != "move":
            # Shift: proporcje jak obraz na slajdzie (16:9 slajdu -> kadr 4:3 obrazu HM)
            h = w * (self.sw / float(self.sh)) * 0.625
        PRESENT.style.set_box(x, y, w, h)
        self._rebuild()

    def mouseReleaseEvent(self, ev):
        if self.drag is None:
            return
        self.drag = None
        self.setCursor(QtCore.Qt.ArrowCursor)
        if self.changed_cb:
            self.changed_cb()

    def mouseDoubleClickEvent(self, ev):
        if self.interactive and self._hit(ev.pos()) in ("move", "switch"):
            PRESENT.style.layout = "custom"
            PRESENT.style.box = list(SlideStyle.DEFAULT_BOX)
            self._rebuild()
            if self.changed_cb:
                self.changed_cb()

    def _rebuild(self):
        if self.rebuild_cb:
            self.ops = self.rebuild_cb() or []
        self.update()


class LayoutBar(_QWidget):
    """Pasek ukladu slajdu: uklad (radio w grupie), tytul, tabela statystyk,
    szybkie dopasowania kadru, zapisane kadry, dane kadru w mm. Wspolny dla
    karty PPTX i okna podgladu - zmiany ida do PRESENT.style i wolaja on_change."""

    def __init__(self, parent, on_change, compact=False):
        super(LayoutBar, self).__init__(parent)
        self.on_change = on_change
        self._loading = False
        v = QtWidgets.QVBoxLayout(self)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(3)
        self.rb = {}
        rbs = [QtWidgets.QLabel(T("Uk\u0142ad:", "Layout:"))]
        for key, pl, en in (("right", "obraz + legenda z prawej", "image + legend right"), ("left", "legenda z lewej", "legend left"),
                            ("full", "sam obraz", "image only"), ("custom", "w\u0142asny kadr (przeci\u0105gaj)", "custom frame (drag)")):
            rb = QtWidgets.QRadioButton(T(pl, en))
            rb.toggled.connect(lambda on, k=key: on and self._set("layout", k))
            self.rb[key] = rb
            rbs.append(rb)
        self.group = radio_group(self, *self.rb.values())
        self.cb_title = QtWidgets.QCheckBox(T("tytu\u0142", "title"))
        self.cb_title.toggled.connect(lambda on: self._set("title_on", on))
        self.cb_stats = QtWidgets.QCheckBox(T("tabela statystyk", "statistics table"))
        self.cb_stats.toggled.connect(lambda on: self._set("stats_on", on))
        v.addLayout(hrow(*(rbs + [12, self.cb_title, self.cb_stats, None])))
        tip(self.rb["custom"], T("Obraz mo\u017cna przeci\u0105ga\u0107 na podgl\u0105dzie (po\u0142o\u017cenie) i skalowa\u0107 uchwytami w rogach (Shift = proporcje). "
                                 "Dwuklik = kadr domy\u015blny. Legenda trafia w wi\u0119ksz\u0105 woln\u0105 stref\u0119 obok obrazu.",
                                 "Drag the image on the preview (position) and scale it with the corner handles (Shift = keep proportions). "
                                 "Double-click = default frame. The legend goes to the larger free area next to the image."))
        # szybkie dopasowania + zapisane kadry (widoczne przy wlasnym kadrze)
        self.quick = QtWidgets.QWidget()
        q = QtWidgets.QHBoxLayout(self.quick)
        q.setContentsMargins(0, 0, 0, 0)
        q.setSpacing(4)
        q.addWidget(QtWidgets.QLabel(T("Kadr:", "Frame:")))
        for key, pl, en, box in QUICK_FRAMES:
            b = QtWidgets.QPushButton(T(pl, en))
            b.setMinimumWidth(56)
            b.clicked.connect(lambda _=False, bx=box: self._quick(bx))
            q.addWidget(b)
        b_c = QtWidgets.QPushButton(T("\u015brodek", "center"))
        b_c.clicked.connect(self._center)
        q.addWidget(b_c)
        q.addSpacing(8)
        self.cmb = QtWidgets.QComboBox()
        self.cmb.setMinimumWidth(130)
        self.cmb.activated.connect(self._preset_pick)
        b_save = QtWidgets.QPushButton(T("+ Zapisz kadr\u2026", "+ Save frame\u2026"))
        b_save.clicked.connect(self._preset_save)
        b_del = QtWidgets.QPushButton(T("Usu\u0144", "Delete"))
        b_del.clicked.connect(self._preset_del)
        q.addWidget(self.cmb)
        q.addWidget(b_save)
        q.addWidget(b_del)
        q.addStretch(1)
        v.addWidget(self.quick)
        self.info = note_label("", HDR_COLOR)
        v.addWidget(self.info)
        self.load()

    def load(self):
        self._loading = True
        st = PRESENT.style
        self.rb.get(st.layout, self.rb["right"]).setChecked(True)
        self.cb_title.setChecked(st.title_on)
        self.cb_stats.setChecked(st.stats_on)
        self._fill_presets()
        self._loading = False
        self._sync()

    def _fill_presets(self):
        self.cmb.clear()
        self.cmb.addItem(T("-- zapisane kadry --", "-- saved frames --"))
        for f in FRAME_PRESETS:
            self.cmb.addItem(f["name"])

    def _sync(self):
        st = PRESENT.style
        custom = st.layout == "custom"
        self.quick.setVisible(custom)
        if custom:
            sw, sh = PRESENT.slide_size()
            b = st.box
            self.info.setText(T("Kadr: X %.1f mm, Y %.1f mm, szer. %.1f mm (%d%%), wys. %.1f mm (%d%%) \u2013 przeci\u0105gaj obraz na podgl\u0105dzie",
                                "Frame: X %.1f mm, Y %.1f mm, width %.1f mm (%d%%), height %.1f mm (%d%%) \u2013 drag the image on the preview",
                                b[0] * sw / MM, b[1] * sh / MM, b[2] * sw / MM, int(round(b[2] * 100)), b[3] * sh / MM, int(round(b[3] * 100))))
        else:
            self.info.setText(T("Przeci\u0105gni\u0119cie obrazu na podgl\u0105dzie prze\u0142\u0105cza uk\u0142ad na \u201ew\u0142asny kadr\u201d.",
                                "Dragging the image on the preview switches the layout to \u201ccustom frame\u201d."))

    def _set(self, key, val):
        if self._loading:
            return
        setattr(PRESENT.style, key, val)
        self._sync()
        if self.on_change:
            self.on_change()

    def _quick(self, box):
        PRESENT.style.layout = "custom"
        PRESENT.style.set_box(*box)
        self.load()
        if self.on_change:
            self.on_change()

    def _center(self):
        b = PRESENT.style.box
        PRESENT.style.layout = "custom"
        PRESENT.style.set_box((1.0 - b[2]) / 2.0, b[1], b[2], b[3])
        self.load()
        if self.on_change:
            self.on_change()

    def _preset_pick(self, idx):
        if idx <= 0 or idx - 1 >= len(FRAME_PRESETS):
            return
        self._quick(FRAME_PRESETS[idx - 1]["box"])
        BUS.status(T("Zastosowano kadr \u201e%s\u201d.", "Frame \u201c%s\u201d applied.", FRAME_PRESETS[idx - 1]["name"]))

    def _preset_save(self):
        name = ask_text(self, T("Zapisz kadr", "Save frame"), T("Nazwa kadru:", "Frame name:"), T("Kadr %d", "Frame %d", len(FRAME_PRESETS) + 1))
        if not name or not name.strip():
            return
        name = name.strip()
        for f in FRAME_PRESETS:
            if f["name"].lower() == name.lower():
                f["box"] = list(PRESENT.style.box)
                break
        else:
            FRAME_PRESETS.append({"name": name, "box": list(PRESENT.style.box)})
        PRESENT.style.layout = "custom"
        save_settings()
        self.load()
        self.cmb.setCurrentIndex([f["name"] for f in FRAME_PRESETS].index(name) + 1)
        BUS.status(T("Zapisano kadr \u201e%s\u201d.", "Frame \u201c%s\u201d saved.", name))

    def _preset_del(self):
        idx = self.cmb.currentIndex()
        if idx <= 0 or idx - 1 >= len(FRAME_PRESETS):
            BUS.status(T("Wybierz zapisany kadr z listy.", "Choose a saved frame from the list."), "warn")
            return
        f = FRAME_PRESETS.pop(idx - 1)
        save_settings()
        self.load()
        BUS.status(T("Usuni\u0119to kadr \u201e%s\u201d.", "Frame \u201c%s\u201d deleted.", f["name"]))

    def refresh_from_style(self):
        """Po przeciagnieciu kadru na plotnie (styl juz zmieniony)."""
        self.load()


class SlidePreview(_QDialog):
    """Podglad WSZYSTKICH slajdow przed zapisem: przegladanie, tytuly, uklad
    (takze wlasny kadr myszka), usuwanie slajdow, Utworz / Anuluj."""

    def __init__(self, parent, items, ask=True):
        super(SlidePreview, self).__init__(parent)
        self.setWindowTitle(T("Podgl\u0105d slajd\u00f3w przed zapisem", "Slide preview before saving") if ask else T("Podgl\u0105d slajdu", "Slide preview"))
        self.items = list(items)
        self.dropped = []
        self.idx = 0
        self.sw, self.sh = PRESENT.slide_size()
        scr = screen_rect(parent or self)
        cw = min(int(scr.width() * 0.58), 1100)
        if cw * self.sh / float(self.sw) > scr.height() * 0.55:
            cw = int(scr.height() * 0.55 * self.sw / float(self.sh))
        outer = QtWidgets.QHBoxLayout(self)
        left = QtWidgets.QVBoxLayout()
        self.canvas = SlideCanvas(self, self.sw, self.sh, min_w=cw, interactive=True)
        self.canvas.rebuild_cb = self._ops
        self.canvas.changed_cb = self._box_changed
        left.addWidget(self.canvas, 1)
        b_prev = QtWidgets.QPushButton("\u25c0")
        b_next = QtWidgets.QPushButton("\u25b6")
        b_prev.clicked.connect(lambda: self.go(-1))
        b_next.clicked.connect(lambda: self.go(1))
        self.lab_n = QtWidgets.QLabel()
        self.lab_n.setMinimumWidth(110)
        self.lab_n.setAlignment(QtCore.Qt.AlignCenter)
        self.b_drop = QtWidgets.QPushButton(T("Usu\u0144 ten slajd", "Remove this slide"))
        self.b_drop.clicked.connect(self.drop)
        left.addLayout(hrow(b_prev, self.lab_n, b_next, None, self.b_drop))
        self.e_title = QtWidgets.QLineEdit()
        self.e_title.textEdited.connect(lambda s: self.set_title(False))
        left.addLayout(hrow(QtWidgets.QLabel(T("Tytu\u0142 slajdu:", "Slide title:")), self.e_title))
        self.bar = LayoutBar(self, self.draw)
        left.addWidget(self.bar)
        outer.addLayout(left, 1)
        right = QtWidgets.QVBoxLayout()
        self.summary = QtWidgets.QPlainTextEdit()
        self.summary.setReadOnly(True)
        self.summary.setMinimumWidth(300)
        right.addWidget(self.summary)
        if ask:
            b_ok = styled_button(T("\u25b6 Utw\u00f3rz prezentacj\u0119", "\u25b6 Create presentation"), GO_COLOR, big=True)
            b_ok.clicked.connect(self.accept)
            b_c = QtWidgets.QPushButton(T("Anuluj", "Cancel"))
            b_c.clicked.connect(self.reject)
            right.addLayout(hrow(b_ok, b_c))
        else:
            b_c = QtWidgets.QPushButton(T("Zamknij", "Close"))
            b_c.clicked.connect(self.reject)
            right.addLayout(hrow(None, b_c))
        outer.addLayout(right)
        QtWidgets.QShortcut(QtGui.QKeySequence(QtCore.Qt.Key_Left), self, lambda: self.go(-1))
        QtWidgets.QShortcut(QtGui.QKeySequence(QtCore.Qt.Key_Right), self, lambda: self.go(1))
        self.draw()

    def _ops(self):
        return slide_ops(self.items[self.idx], self.sw, self.sh, PRESENT.style) if self.items else []

    def _box_changed(self):
        self.bar.refresh_from_style()
        save_settings()

    def go(self, d):
        if self.items:
            self.idx = (self.idx + d) % len(self.items)
            self.draw()

    def set_title(self, redraw=True):
        if self.items:
            self.items[self.idx]["title"] = self.e_title.text()
            if redraw:
                self.draw()
            else:
                self.canvas.set_ops(self._ops())
                self._summary()

    def drop(self):
        if len(self.items) <= 1:
            return
        self.dropped.append(self.items.pop(self.idx))
        self.idx = min(self.idx, len(self.items) - 1)
        self.draw()

    def draw(self):
        it = self.items[self.idx]
        self.lab_n.setText(T("slajd %d / %d", "slide %d / %d", self.idx + 1, len(self.items)))
        if self.e_title.text() != it.get("title", ""):
            self.e_title.setText(it.get("title", ""))
        self.canvas.set_ops(self._ops())
        self.b_drop.setEnabled(len(self.items) > 1)
        self._summary()

    def _summary(self):
        p = PRESENT.target(False)
        L = [T("PLIK:", "FILE:"), "  " + (p or T("(wybierzesz przy zapisie)", "(chosen when saving)"))]
        if p and os.path.isfile(p) and PRESENT.mode == "append":
            L.append(T("  tryb: DOPISANIE slajd\u00f3w na ko\u0144cu (kopia .bak.pptx)", "  mode: APPEND slides at the end (.bak.pptx backup)"))
        elif p and os.path.isfile(p):
            L.append(T("  tryb: NOWY plik \u2013 istniej\u0105cy zostanie ZAST\u0104PIONY", "  mode: NEW file \u2013 the existing one will be REPLACED"))
        else:
            L.append(T("  tryb: nowa prezentacja 16:9", "  mode: new 16:9 presentation"))
        L += ["", T("SLAJDY (%d):", "SLIDES (%d):", len(self.items))]
        for i, x in enumerate(self.items):
            L.append("%s %d. %s" % ("\u25b6" if i == self.idx else " ", i + 1, x.get("title", "")))
        L += ["", T("Obrazy: pliki zrzut\u00f3w (WYSIWYG \u2013 dok\u0142adnie te klatki trafi\u0105 na slajdy). Legenda i tabele: natywne, edytowalne kszta\u0142ty PowerPointa. "
                    "Strza\u0142ki \u25c0 \u25b6 (tak\u017ce klawisze) prze\u0142\u0105czaj\u0105 slajdy; uk\u0142ad i kadr dotycz\u0105 wszystkich slajd\u00f3w.",
                    "Images: the captured frames (WYSIWYG \u2013 exactly these frames go to the slides). Legend and tables: native, editable PowerPoint shapes. "
                    "Arrows \u25c0 \u25b6 (also keys) switch slides; the layout and frame apply to all slides.")]
        self.summary.setPlainText("\n".join(L))

    @staticmethod
    def run(parent, items, ask=True):
        d = SlidePreview(parent, items, ask)
        ok = d.exec_() == QtWidgets.QDialog.Accepted
        save_settings()
        Presenter.drop_items(d.dropped)
        return ok, d.items


class PreviewPanel(_QWidget):
    """Podglad tylko do odczytu: przykladowy slajd z biezacego stanu + "CO
    POWSTANIE". Uzywany przez karte "Podglad slajdu" i okno "Podglad na zywo"."""

    def __init__(self, parent, studio, min_w=420):
        super(PreviewPanel, self).__init__(parent)
        self.studio = studio
        self.item = None
        v = QtWidgets.QVBoxLayout(self)
        v.setContentsMargins(0, 0, 0, 0)
        sw, sh = PRESENT.slide_size()
        self.canvas = SlideCanvas(self, sw, sh, min_w=min_w, interactive=False)
        self.canvas.badge = T("tylko podgl\u0105d", "preview only")
        v.addWidget(self.canvas, 3)
        self.ctx = note_label("", "#556")
        v.addWidget(self.ctx)
        self.text = QtWidgets.QPlainTextEdit()
        self.text.setReadOnly(True)
        self.text.setMaximumHeight(190)
        self.text.setStyleSheet("font-size:8.5pt")
        v.addWidget(self.text, 1)

    def refresh(self, capture=True):
        """Nowy obraz z biezacych opcji; capture=False - bez zrzutu z HM
        (np. w trakcie innej operacji)."""
        sw, sh = PRESENT.slide_size()
        self.canvas.set_slide_size(sw, sh)
        try:
            self.item = PRESENT.sample_item(capture)
            ops = slide_ops(self.item, sw, sh, PRESENT.style)
        except Exception as e:
            self.item, ops = None, []
            BUS.log("preview: %s" % e)
        self.canvas.clear_cache()
        self.canvas.set_ops(ops)
        src = T("brak klatki \u2013 zrzut powstanie przy eksporcie", "no frame \u2013 the shot is taken at export")
        if self.item and self.item.get("img") and os.path.isfile(self.item["img"]):
            src = T("klatka z okna graficznego HM (odnawiana po zmianie widoku / analizy)", "frame from the HM graphics window (renewed after a view / analysis change)")
        kind = {"metric": T("slajd metryki", "metric slide"), "delta": T("slajd delty", "delta slide"), "view": T("slajd widoku", "view slide")}.get(
            (self.item or {}).get("kind", ""), "")
        self.ctx.setText(T("Przyk\u0142adowy slajd: %s \u2013 to, co da \u201eBie\u017c\u0105cy widok \u2192 slajd\u201d. Obraz: %s. Proporcje: %s.",
                           "Sample slide: %s \u2013 what \u201cCurrent view \u2192 slide\u201d produces. Image: %s. Proportions: %s.",
                           kind or "-", src, T("16:9 (nowa) / z istniej\u0105cego pliku", "16:9 (new) / from the existing file")))
        try:
            self.text.setPlainText("\n".join([T("CO POWSTANIE Z BIE\u017b\u0104CYCH OPCJI:", "WHAT THE CURRENT OPTIONS PRODUCE:")] + PRESENT.plan_lines()))
        except Exception as e:
            self.text.setPlainText("%s" % e)


class LivePreviewWindow(_QWidget):
    """Okno "Podglad na zywo" (F6): niemodalne, obok okna narzedzia, odswiezane
    samo po kazdej zmianie opcji (z opoznieniem). Tylko do odczytu."""

    def __init__(self, studio):
        super(LivePreviewWindow, self).__init__(studio, QtCore.Qt.Tool | QtCore.Qt.WindowTitleHint | QtCore.Qt.WindowCloseButtonHint)
        self.studio = studio
        self.setWindowTitle(T("Podgl\u0105d na \u017cywo", "Live preview"))
        v = QtWidgets.QVBoxLayout(self)
        v.setContentsMargins(6, 4, 6, 6)
        v.addWidget(note_label(T("Od\u015bwie\u017ca si\u0119 samo po ka\u017cdej zmianie opcji. Tylko do odczytu. F6 / Esc \u2013 zamknij.",
                                 "Refreshes by itself after every option change. Read only. F6 / Esc \u2013 close."), "#556"))
        self.panel = PreviewPanel(self, studio, min_w=360)
        v.addWidget(self.panel, 1)
        geo = GUI_PREFS.get("live_geo") or []
        placed = False
        if len(geo) == 4:
            r = QtCore.QRect(*[int(x) for x in geo])
            if screen_rect(studio).intersects(r):
                self.setGeometry(r)
                placed = True
        if not placed:
            g = studio.frameGeometry()
            scr = screen_rect(studio)
            w, h = 620, 520
            x = g.right() + 8
            if x + w > scr.right():
                x = max(scr.left(), scr.right() - w - 8)
            self.setGeometry(x, max(scr.top(), min(g.top() + 40, scr.bottom() - h)), w, h)
        QtWidgets.QShortcut(QtGui.QKeySequence(QtCore.Qt.Key_Escape), self, self.close)
        QtWidgets.QShortcut(QtGui.QKeySequence(QtCore.Qt.Key_F6), self, self.close)

    def refresh(self, capture=True):
        self.panel.refresh(capture)
        it = self.panel.item
        self.setWindowTitle(T("Podgl\u0105d na \u017cywo: %s", "Live preview: %s", it.get("title", "")) if it else T("Podgl\u0105d na \u017cywo", "Live preview"))

    def closeEvent(self, ev):
        g = self.geometry()
        GUI_PREFS["live_geo"] = [g.x(), g.y(), g.width(), g.height()]
        GUI_PREFS["live_open"] = bool(getattr(self.studio, "_closing", False))
        save_settings()
        self.studio.live = None
        super(LivePreviewWindow, self).closeEvent(ev)


class ExportDoneDialog(_QDialog):
    """Po eksporcie: "Eksport gotowy" + galeria wyeksportowanych slajdow
    (dokladnie to, co trafilo do pliku) + Otworz prezentacje / folder / dziennik."""

    def __init__(self, parent, info):
        super(ExportDoneDialog, self).__init__(parent)
        self.setWindowTitle(T("Eksport gotowy", "Export finished"))
        self.info = info
        self.thumbs = [p for p in (info.get("thumbs") or []) if os.path.isfile(p)]
        self.idx = 0
        v = QtWidgets.QVBoxLayout(self)
        head = QtWidgets.QLabel(T("<b>Eksport gotowy.</b> %d slajd(\u00f3w) \u2192 %s", "<b>Export finished.</b> %d slide(s) \u2192 %s",
                                  info.get("n", 0), h_esc(info.get("path", ""))))
        head.setTextFormat(QtCore.Qt.RichText)
        head.setWordWrap(True)
        v.addWidget(head)
        if info.get("probs"):
            v.addWidget(note_label(T("Samokontrola pliku: %s", "File self-check: %s", "; ".join(info["probs"])), WARN_COLOR))
        self.img = QtWidgets.QLabel()
        self.img.setAlignment(QtCore.Qt.AlignCenter)
        self.img.setStyleSheet("background:#eef2f6;border:1px solid #c8d0da")
        scr = screen_rect(parent or self)
        self.tw = min(900, int(scr.width() * 0.5))
        self.img.setFixedSize(self.tw, int(self.tw * 9 / 16.0))
        v.addWidget(self.img)
        b_prev = QtWidgets.QPushButton("\u25c0")
        b_next = QtWidgets.QPushButton("\u25b6")
        b_prev.clicked.connect(lambda: self.go(-1))
        b_next.clicked.connect(lambda: self.go(1))
        self.lab = QtWidgets.QLabel()
        self.lab.setAlignment(QtCore.Qt.AlignCenter)
        v.addLayout(hrow(b_prev, self.lab, b_next))
        b_open = styled_button(T("Otw\u00f3rz prezentacj\u0119", "Open the presentation"), GO_COLOR)
        b_open.clicked.connect(lambda: open_path(info.get("path", "")))
        b_dir = QtWidgets.QPushButton(T("Otw\u00f3rz folder", "Open folder"))
        b_dir.clicked.connect(lambda: open_path(os.path.dirname(info.get("path", "")) or "."))
        b_log = QtWidgets.QPushButton(T("Dziennik", "Log"))
        b_log.clicked.connect(lambda: parent.b_log.setChecked(True) if parent is not None and hasattr(parent, "b_log") else None)
        b_ok = QtWidgets.QPushButton(T("Zamknij", "Close"))
        b_ok.clicked.connect(self.accept)
        v.addLayout(hrow(b_open, b_dir, b_log, None, b_ok))
        QtWidgets.QShortcut(QtGui.QKeySequence(QtCore.Qt.Key_Left), self, lambda: self.go(-1))
        QtWidgets.QShortcut(QtGui.QKeySequence(QtCore.Qt.Key_Right), self, lambda: self.go(1))
        self.show_thumb()

    def go(self, d):
        if self.thumbs:
            self.idx = (self.idx + d) % len(self.thumbs)
            self.show_thumb()

    def show_thumb(self):
        titles = self.info.get("titles") or []
        if not self.thumbs:
            self.img.setText(T("(brak podgl\u0105du slajd\u00f3w)", "(no slide previews)"))
            self.lab.setText("")
            return
        pm = QtGui.QPixmap(self.thumbs[self.idx])
        if not pm.isNull():
            self.img.setPixmap(pm.scaled(self.img.size(), QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation))
        t = titles[self.idx] if self.idx < len(titles) else ""
        self.lab.setText(T("slajd %d / %d \u2013 %s", "slide %d / %d \u2013 %s", self.idx + 1, len(self.thumbs), t))


# ============================ GUI: DIAGNOSTYKA ==========================
class DiagnosticsDialog(_QDialog):
    """Samokontrola odczytu z HyperMesha (API, most Tcl, nazwy danych metryk
    z wartosciami, zgodnosc odczytu hurtowego z pojedynczym, wzorce geometrii,
    zrzut, biblioteki) - "Weryfikacja obliczen przez makro"."""

    def __init__(self, studio):
        super(DiagnosticsDialog, self).__init__(studio)
        self.studio = studio
        self.setWindowTitle(T("Diagnostyka i weryfikacja odczytu metryk", "Diagnostics and metric read verification"))
        v = QtWidgets.QVBoxLayout(self)
        v.addWidget(note_label(T("Sprawdza, czy makro poprawnie czyta metryki z HyperMesha: dost\u0119pno\u015b\u0107 API i mostu Tcl, nazw\u0119 danych "
                                 "ka\u017cdej metryki (kandydaci: aspect / aspectratio, skew / skewness \u2026) z warto\u015bciami na pr\u00f3bce element\u00f3w "
                                 "ka\u017cdego typu, zgodno\u015b\u0107 odczytu hurtowego (mark=1) z odczytem pojedynczym po ID, kolejno\u015b\u0107 odpowiedzi, "
                                 "wzorce geometrii (Jacobian Zero, tet collapse), zrzut okna i biblioteki.",
                                 "Checks that the macro reads metrics from HyperMesh correctly: API and Tcl bridge availability, the data name "
                                 "of each metric (candidates: aspect / aspectratio, skew / skewness \u2026) with values on a sample of elements of "
                                 "every type, bulk (mark=1) vs single-ID read consistency, reply order, geometry references (Jacobian Zero, "
                                 "tet collapse), capture and libraries.")))
        self.tb = QtWidgets.QTextBrowser()
        self.tb.setMinimumSize(720, 420)
        v.addWidget(self.tb)
        b_run = styled_button(T("\u25b6 Uruchom diagnostyk\u0119", "\u25b6 Run diagnostics"), RUN_COLOR)
        b_run.clicked.connect(self.run)
        b_copy = QtWidgets.QPushButton(T("Kopiuj do schowka", "Copy to clipboard"))
        b_copy.clicked.connect(lambda: app_instance().clipboard().setText(self.tb.toPlainText()))
        b_close = QtWidgets.QPushButton(T("Zamknij", "Close"))
        b_close.clicked.connect(self.close)
        v.addLayout(hrow(b_run, b_copy, None, b_close))
        self.rows = []

    def run(self):
        def work():
            return HM.diagnostics()

        def done(rows):
            self.rows = rows or []
            H = ["<table cellspacing=0 cellpadding=4 border=1 style='border-collapse:collapse;border-color:#cfd8e3'>",
                 "<tr style='background:#e8eef5'><th></th><th>%s</th><th>%s</th></tr>" % (h_esc(T("Sprawdzenie", "Check")), h_esc(T("Wynik", "Result")))]
            nbad = 0
            for name, ok, det in self.rows:
                nbad += 0 if ok else 1
                H.append("<tr><td style='color:%s;font-weight:bold'>%s</td><td>%s</td><td>%s</td></tr>"
                         % (OK_COLOR if ok else ERR_COLOR, "\u2713" if ok else "\u2717", h_esc(name), h_esc(det)))
            H.append("</table>")
            H.append("<p style='color:%s'><b>%s</b></p>" % (OK_COLOR if not nbad else WARN_COLOR,
                                                            h_esc(T("Wszystkie sprawdzenia OK.", "All checks OK.") if not nbad else
                                                                  T("Sprawdze\u0144 z uwagami: %d (metryki bez warto\u015bci mog\u0105 nie dotyczy\u0107 typ\u00f3w element\u00f3w w modelu).",
                                                                    "Checks with remarks: %d (metrics without values may not apply to the element types in the model).", nbad))))
            self.tb.setHtml("".join(H))
            BUS.log(T("Diagnostyka: %d sprawdze\u0144, z uwagami %d", "Diagnostics: %d checks, %d with remarks", len(self.rows), nbad))
            for name, ok, det in self.rows:
                BUS.log("  %s %s: %s" % ("+" if ok else "-", name, det))
        self.studio.runner.run(work, done)


# =========================== GUI: KARTA "START" =========================
# Przeglad stanu pracy: model w sesji, pieciu krokow przebiegu (co jest
# zrobione, co dalej) z przyciskami "Przejdz" i szybkimi akcjami, narzedzia
# (przywroc siatke, diagnostyka, pomoc). Karta odswieza sie po kazdej akcji.
class StartTab(_QWidget):
    def __init__(self, studio):
        super(StartTab, self).__init__()
        self.studio = studio
        v = QtWidgets.QVBoxLayout(self)
        # --- model ---
        self.lab_model = QtWidgets.QLabel()
        self.lab_model.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
        self.lab_model.setWordWrap(True)
        b_ref = QtWidgets.QPushButton(T("Od\u015bwie\u017c", "Refresh"))
        b_ref.clicked.connect(lambda: self.studio.runner.run(self.refresh, busy=False))
        tip(b_ref, T("Odczytuje ponownie liczby element\u00f3w, w\u0119z\u0142\u00f3w i komponent\u00f3w z sesji HyperMesha.",
                     "Re-reads the element, node and component counts from the HyperMesh session."))
        v.addWidget(group(T("Model w sesji HyperMesha", "Model in the HyperMesh session"), hrow(self.lab_model, None, b_ref)))
        # --- kroki ---
        grid = QtWidgets.QGridLayout()
        grid.setSpacing(8)
        self.cards = {}
        steps = [("mq", 1, T("Metryki i grupy kolor\u00f3w", "Metrics & color groups"), T("\u25b6 Analizuj (F5)", "\u25b6 Analyze (F5)"), lambda: self.studio.tab_mq.analyze()),
                 ("delta", 2, T("Delta REF / INF", "Delta REF / INF"), T("Wykonaj analiz\u0119", "Run the analysis"), lambda: self.studio.tab_delta.run()),
                 ("report", 3, T("Raport jako\u015bci", "Quality report"), T("Generuj raport", "Generate report"), lambda: self.studio.tab_report.generate()),
                 ("views", 4, T("Widoki i zrzuty", "Views & screenshots"), T("Zapami\u0119taj widok (1)", "Remember view (1)"), lambda: self.studio.tab_views.remember()),
                 ("ppt", 5, T("Prezentacja PPTX", "PowerPoint"), T("\u25b6 Seria \u2192 PPTX (F7)", "\u25b6 Series \u2192 PPTX (F7)"), lambda: self.studio.tab_ppt.series()),
                 ("preview", 6, T("Podgl\u0105d slajdu", "Slide preview"), T("Podgl\u0105d na \u017cywo (F6)", "Live preview (F6)"), lambda: self.studio.live_toggle())]
        for i, (key, num, title, act, fn) in enumerate(steps):
            card = QtWidgets.QFrame()
            card.setObjectName("card")
            card.setStyleSheet("#card{background:white;border:1px solid #c8d0da;border-radius:4px}")
            cl = QtWidgets.QVBoxLayout(card)
            cl.setContentsMargins(10, 8, 10, 8)
            cl.addWidget(step_label(num if num <= 5 else "\u25a3", title))
            st = QtWidgets.QLabel()
            st.setWordWrap(True)
            st.setMinimumHeight(52)
            st.setAlignment(QtCore.Qt.AlignTop)
            cl.addWidget(st, 1)
            b_go = QtWidgets.QPushButton(T("Przejd\u017a \u2192", "Go \u2192"))
            b_go.clicked.connect(lambda _=False, k=key: self.studio.goto(k))
            b_act = styled_button(act, RUN_COLOR if key in ("mq", "delta") else (GO_COLOR if key in ("ppt", "report") else HDR_COLOR))
            b_act.clicked.connect(lambda _=False, f=fn: f())
            cl.addLayout(hrow(b_go, None, b_act))
            grid.addWidget(card, i // 3, i % 3)
            self.cards[key] = st
        v.addWidget(group(T("Przebieg pracy", "Workflow"), grid))
        # --- narzedzia ---
        b_rs = QtWidgets.QPushButton(T("Przywr\u00f3\u0107 siatk\u0119 (zdejmij kolory)", "Restore mesh (remove colors)"))
        b_rs.clicked.connect(self.studio.restore_any)
        tip(b_rs, T("Odk\u0142ada elementy do pierwotnych komponent\u00f3w i usuwa komponenty narz\u0119dzia (tylko puste) \u2013 nic nie ginie.",
                    "Puts elements back into their original components and removes the tool components (empty only) \u2013 nothing is lost."))
        b_diag = QtWidgets.QPushButton(T("Diagnostyka / weryfikacja odczytu metryk\u2026", "Diagnostics / metric read verification\u2026"))
        b_diag.clicked.connect(self.studio.show_diagnostics)
        b_help = QtWidgets.QPushButton(T("? Pomoc (F1)", "? Help (F1)"))
        b_help.clicked.connect(self.studio.show_help)
        v.addWidget(group(T("Narz\u0119dzia", "Tools"), hrow(b_rs, b_diag, b_help, None)))
        v.addWidget(note_label(T("Szybki start: 1) na karcie \u201eMetryki\u201d zaznacz metryki i kliknij \u201eAnalizuj i koloruj\u201d (F5); 2) ustaw model w oknie graficznym "
                                 "i zapami\u0119taj widoki (klawisz 1) \u2013 opcjonalnie; 3) na karcie \u201ePrezentacja\u201d wybierz plik i uk\u0142ad (podgl\u0105d na \u017cywo: F6), "
                                 "kliknij \u201eSeria \u2192 PPTX\u201d (F7) \u2013 najpierw zobaczysz wszystkie slajdy, potem powstanie plik. Delta REF/INF i raport to osobne kroki, "
                                 "te\u017c z eksportem na slajd. Wszystkie opcje zapisuj\u0105 si\u0119 same.",
                                 "Quick start: 1) on the \u201cMetrics\u201d tab tick metrics and click \u201cAnalyze and color\u201d (F5); 2) set the model in the graphics window "
                                 "and remember views (key 1) \u2013 optional; 3) on the \u201cPowerPoint\u201d tab choose the file and layout (live preview: F6), "
                                 "click \u201cSeries \u2192 PPTX\u201d (F7) \u2013 you first see all slides, then the file is written. The REF/INF delta and the report are separate steps, "
                                 "also exportable to a slide. All options are saved automatically.")))
        v.addStretch(1)

    def refresh(self):
        if HM.ok():
            f = HM.model_file()
            ne, nn, nc = HM.count(HM.ent.Element), HM.count(HM.ent.Node), HM.count(HM.ent.Component)
            s = "<b>%s</b><br>%s" % (h_esc(f or T("(model niezapisany / nieznany plik)", "(unsaved model / unknown file)")),
                                     h_esc(T("elementy: %d \u2022 w\u0119z\u0142y: %d \u2022 komponenty: %d", "elements: %d \u2022 nodes: %d \u2022 components: %d", ne, nn, nc)))
            if MQ.analyzed:
                s += "<br>" + h_esc(T("z ostatniej analizy: 2D %d, 3D %d, inne %d", "from the last analysis: 2D %d, 3D %d, other %d", MQ.n2d, MQ.n3d, MQ.nskip))
            if MESH.active():
                s += "<br><span style='color:%s'>%s</span>" % (WARN_COLOR, h_esc(T("siatka jest POKOLOROWANA przez narz\u0119dzie (%s) \u2013 \u201ePrzywr\u00f3\u0107 siatk\u0119\u201d zdejmuje kolory",
                                                                                      "the mesh is COLORED by the tool (%s) \u2013 \u201cRestore mesh\u201d removes the colors",
                                                                                      {"mq": T("metryki", "metrics"), "delta": T("delta", "delta")}.get(MESH.owner, MESH.owner))))
        else:
            s = "<span style='color:%s'>%s</span>" % (ERR_COLOR, h_esc(T("Brak API HyperMesha \u2013 uruchom plik w HyperMesh 2023+ (File > Run > Python Script).",
                                                                        "No HyperMesh API \u2013 run the file in HyperMesh 2023+ (File > Run > Python Script).")))
        self.lab_model.setText(s)
        c = self.cards
        if MQ.analyzed:
            c["mq"].setText(T("Analiza z %s: %d element\u00f3w, metryki: %s.\nNa siatce: %s. Poza norm\u0105 (\u2265 1 metryka): %d.",
                              "Analysis from %s: %d elements, metrics: %s.\nOn the mesh: %s. Out of limits (\u2265 1 metric): %d.",
                              MQ.when, len(MQ.elems), ", ".join(MQ.tag(m) for m in MQ.analyzed),
                              MQ.label(MQ.view) if MQ.view else T("(kolory zdj\u0119te)", "(colors removed)"),
                              MQ.stats.get("all", {}).get("nBad", 0)))
        else:
            c["mq"].setText(T("Brak analizy. Zaznacz metryki (AR, Jacobian, Jacobian Zero, Skewness) i kliknij \u201eAnalizuj i koloruj\u201d \u2013 siatka zostanie pokolorowana wg prog\u00f3w.",
                              "No analysis yet. Tick metrics (AR, Jacobian, Jacobian Zero, Skewness) and click \u201cAnalyze and color\u201d \u2013 the mesh is colored by thresholds."))
        if DELTA.done:
            cnt = DELTA.counts
            c["delta"].setText(T("Wynik na siatce: %s (%s), %s.\nW skali %d \u2022 bez zmian %d \u2022 bez odpowiednika %d \u2022 MAX %s (el. %s)",
                                 "Result on the mesh: %s (%s), %s.\nIn scale %d \u2022 no change %d \u2022 unmatched %d \u2022 MAX %s (el. %s)",
                                 DELTA.delta_label(), DELTA.dim_label(), DELTA.when, cnt["band"], cnt["gray"], cnt["unm"],
                                 fmt_num(DELTA.max_info[0], 4) if DELTA.max_info else "-", DELTA.max_info[1] if DELTA.max_info else "-"))
        else:
            c["delta"].setText(T("Brak wyniku. Por\u00f3wnuje dwa modele .hm (REF \u2192 INF) po ID element\u00f3w: pogorszenie AR / Jacobiana / Skewness albo przesuni\u0119cie w\u0119z\u0142\u00f3w [mm]; jeden model: pasma wg warto\u015bci albo podzia\u0142 wg progu.",
                                 "No result. Compares two .hm models (REF \u2192 INF) by element ID: AR / Jacobian / Skewness worsening or node displacement [mm]; one model: bands by value or threshold split."))
        if REPORT.last:
            c["report"].setText(T("Ostatni raport: %s (%d element\u00f3w).\nPliki: %s", "Last report: %s (%d elements).\nFiles: %s",
                                  REPORT.last["when"], REPORT.last["nsel"], ", ".join(os.path.basename(f) for f in REPORT.last_files[:4]) or "-"))
        else:
            c["report"].setText(T("Brak raportu. TXT / CSV / XLSX / HTML ze statystykami, histogramami, wska\u017anikiem 0\u2013100; por\u00f3wnanie REF vs INF ze zgodno\u015bci\u0105 element\u00f3w i zestawieniem w\u0119z\u0142\u00f3w.",
                                  "No report. TXT / CSV / XLSX / HTML with statistics, histograms, 0\u2013100 score; REF vs INF comparison with element identity and node comparison."))
        c["views"].setText(T("Zapami\u0119tanych widok\u00f3w: %d (zaznaczonych: %d, z klatk\u0105 WYSIWYG: %d).\nZrzuty: folder %s.",
                             "Remembered views: %d (ticked: %d, with a WYSIWYG frame: %d).\nScreenshots: folder %s.",
                             len(VIEWS.views), len(VIEWS.selected()), sum(1 for r in VIEWS.views if VIEWS.shot_ok(r)),
                             VIEWS.target_dir()[0] or T("(nie wybrano)", "(not chosen)")))
        p = PRESENT.target(False)
        le = PRESENT.last_export
        c["ppt"].setText(T("Plik: %s\nTryb: %s, uk\u0142ad: %s.%s", "File: %s\nMode: %s, layout: %s.%s",
                           os.path.basename(p) if p else T("(nie wybrano \u2013 zapytam przy zapisie)", "(not chosen \u2013 asked when saving)"),
                           T("dopisanie slajd\u00f3w", "append slides") if PRESENT.mode == "append" else T("nowa prezentacja", "new presentation"),
                           PRESENT.style.layout,
                           ("\n" + T("Ostatni eksport: %d slajd(\u00f3w), %s", "Last export: %d slide(s), %s", le["n"], le["when"])) if le else ""))
        c["preview"].setText(T("Karta \u201ePodgl\u0105d slajdu\u201d pokazuje przyk\u0142adowy slajd z bie\u017c\u0105cych opcji i \u201eco powstanie\u201d; okno na \u017cywo (F6) odswie\u017ca si\u0119 po ka\u017cdej zmianie opcji.",
                               "The \u201cSlide preview\u201d tab shows a sample slide from the current options and \u201cwhat will be made\u201d; the live window (F6) refreshes after every option change."))


# ==================== GUI: KARTA "METRYKI I GRUPY" =====================
# Kroki: 1 zakres i wymiar elementow, 2 tabela metryk (prog, tryb legendy,
# liczba pasm, "Legenda..."), 3 opcje grup i kolory specjalne, 4 "Analizuj
# i koloruj"; przelacznik widoku na siatce i tabela wynikow.
class MqTab(_QWidget):
    def __init__(self, studio):
        super(MqTab, self).__init__()
        self.studio = studio
        v = QtWidgets.QVBoxLayout(self)
        # --- zakres ---
        self.rb_all = QtWidgets.QRadioButton(T("ca\u0142a siatka", "whole mesh"))
        self.rb_disp = QtWidgets.QRadioButton(T("tylko wy\u015bwietlone", "displayed only"))
        radio_group(self, self.rb_all, self.rb_disp)
        self.cb_2d = QtWidgets.QCheckBox(T("2D (pow\u0142oki)", "2D (shells)"))
        self.cb_3d = QtWidgets.QCheckBox(T("3D (bry\u0142y)", "3D (solids)"))
        self.e_prefix = QtWidgets.QLineEdit()
        self.e_prefix.setMaximumWidth(90)
        tip(self.e_prefix, T("Prefiks nazw komponent\u00f3w i zestaw\u00f3w tworzonych na siatce (np. MQ_AR_s01_4_5).",
                             "Prefix of the component and set names created on the mesh (e.g. MQ_AR_s01_4_5)."))
        v.addWidget(group(T("1  Zakres", "1  Scope"), hrow(self.rb_all, self.rb_disp, 16, self.cb_2d, self.cb_3d, None,
                                                            QtWidgets.QLabel(T("Prefiks grup:", "Group prefix:")), self.e_prefix)))
        # --- metryki ---
        grid = QtWidgets.QGridLayout()
        heads = [T("Metryka", "Metric"), T("Poza norm\u0105", "Out of limits"), T("Pr\u00f3g", "Threshold"), T("Legenda", "Legend"),
                 T("Pasm (auto)", "Bands (auto)"), T("Przedzia\u0142y", "Bands"), ""]
        for c, h in enumerate(heads):
            lab = QtWidgets.QLabel(h)
            lab.setStyleSheet("font-weight:bold;color:#444")
            grid.addWidget(lab, 0, c)
        self.w = {}
        for r, m in enumerate(MQ_ORDER, start=1):
            d = MQ_DEF[m]
            use = QtWidgets.QCheckBox(d.label)
            use.setStyleSheet("font-weight:bold")
            if d.dir == "above":
                dl = QtWidgets.QLabel(T("> pr\u00f3g (wy\u017csze gorsze)", "> thr. (higher worse)"))
            else:
                dl = QtWidgets.QLabel(T("%s pr\u00f3g (ni\u017csze gorsze)", "%s thr. (lower worse)", "\u2264" if d.strict else "<"))
            dl.setStyleSheet("color:#555")
            thr_on = QtWidgets.QCheckBox()
            tip(thr_on, T("Podzia\u0142 wg progu: w normie = jedna zielona grupa, poza norm\u0105 = pasma kolor\u00f3w.",
                          "Threshold split: within limits = one green group, out of limits = color bands."))
            thr = QtWidgets.QLineEdit()
            thr.setMaximumWidth(70)
            ra = QtWidgets.QRadioButton("auto")
            rm = QtWidgets.QRadioButton(T("r\u0119czna", "manual"))
            radio_group(self, ra, rm)
            nb = QtWidgets.QSpinBox()
            nb.setRange(1, 40)
            info = QtWidgets.QLabel()
            info.setStyleSheet("color:%s" % HDR_COLOR)
            ed = QtWidgets.QPushButton(T("Legenda\u2026", "Legend\u2026"))
            ed.clicked.connect(lambda _=False, m=m: (self.store_to_engine(), self.studio.open_editor(m)))
            ra.toggled.connect(self.update_rows)
            grid.addWidget(use, r, 0)
            grid.addWidget(dl, r, 1)
            grid.addLayout(hrow(thr_on, thr, spacing=2), r, 2)
            grid.addLayout(hrow(ra, rm, spacing=2), r, 3)
            grid.addWidget(nb, r, 4)
            grid.addWidget(info, r, 5)
            grid.addWidget(ed, r, 6)
            self.w[m] = {"use": use, "thr_on": thr_on, "thr": thr, "auto": ra, "manual": rm, "nb": nb, "info": info}
        grid.setColumnStretch(5, 1)
        gv = QtWidgets.QVBoxLayout()
        gv.addLayout(grid)
        gv.addWidget(note_label(T("Jacobian Ratio = min/max det(J) z HyperMesha (0\u20261). Jacobian Zero = najmniejszy det(J) w naro\u017cach elementu; "
                                  "\u2264 0 = element zdegenerowany (liczony z w\u0119z\u0142\u00f3w). Metryki s\u0105 czytane jednocze\u015bnie, w jednym przebiegu po elementach. "
                                  "Nazwy danych HM s\u0105 dobierane automatycznie (\u201eDiagnostyka\u201d na karcie Start pokazuje, kt\u00f3re dzia\u0142aj\u0105).",
                                  "Jacobian Ratio = min/max det(J) from HyperMesh (0\u20261). Jacobian Zero = smallest det(J) at element corners; "
                                  "\u2264 0 = degenerate element (computed from nodes). The metrics are read simultaneously, in one pass over the elements. "
                                  "HM data names are chosen automatically (\u201cDiagnostics\u201d on the Start tab shows which ones work).")))
        v.addWidget(group(T("2  Metryki (sprawdzane jednocze\u015bnie)", "2  Metrics (checked simultaneously)"), gv))
        # --- opcje ---
        self.cb_over = QtWidgets.QCheckBox(T("poza skal\u0105 = osobna grupa", "beyond scale = own group"))
        self.cb_sets = QtWidgets.QCheckBox(T("zestawy (sets) dla wszystkich metryk", "sets for all metrics"))
        self.cb_worst = QtWidgets.QCheckBox(T("tag przy najgorszym", "tag the worst"))
        self.cb_hide = QtWidgets.QCheckBox(T("ukryj elementy w normie", "hide elements within limits"))
        self.cb_hide.toggled.connect(lambda on: self.studio.runner.run(lambda: MQ.set_hide_ok(on), busy=False))
        self.cb_nice = QtWidgets.QCheckBox(T("\u0142adne liczby (auto)", "nice numbers (auto)"))
        self.c_ok = ColorButton(MQ.col_ok, lambda c: setattr(MQ, "col_ok", c))
        self.c_over = ColorButton(MQ.col_over, lambda c: setattr(MQ, "col_over", c))
        self.c_na = ColorButton(MQ.col_na, lambda c: setattr(MQ, "col_na", c))
        og = QtWidgets.QGridLayout()
        og.addWidget(self.cb_over, 0, 0)
        og.addWidget(self.cb_sets, 0, 1)
        og.addWidget(self.cb_nice, 0, 2)
        og.addWidget(self.cb_worst, 1, 0)
        og.addWidget(self.cb_hide, 1, 1)
        og.addLayout(hrow(QtWidgets.QLabel(T("Kolor \u201ew normie\u201d:", "\u201cWithin\u201d color:")), self.c_ok, 8,
                          QtWidgets.QLabel(T("poza skal\u0105:", "beyond scale:")), self.c_over, 8,
                          QtWidgets.QLabel(T("n/d:", "n/a:")), self.c_na, None), 2, 0, 1, 3)
        og.setColumnStretch(3, 1)
        v.addWidget(group(T("3  Opcje grup", "3  Group options"), og))
        # --- akcje ---
        self.b_go = styled_button(T("\u25b6 Analizuj i koloruj (F5)", "\u25b6 Analyze and color (F5)"), RUN_COLOR, big=True)
        self.b_go.clicked.connect(self.analyze)
        tip(self.b_go, T("Czyta zaznaczone metryki w jednym przebiegu, dzieli elementy na grupy wg prog\u00f3w i koloruje siatk\u0119 (komponenty).",
                         "Reads the ticked metrics in one pass, splits elements into groups by thresholds and colors the mesh (components)."))
        b_leg = QtWidgets.QPushButton(T("Legenda", "Legend"))
        b_leg.clicked.connect(lambda: self.studio.show_legend("mq"))
        b_rs = QtWidgets.QPushButton(T("Przywr\u00f3\u0107 siatk\u0119", "Restore mesh"))
        b_rs.clicked.connect(self.restore)
        b_ppt = QtWidgets.QPushButton(T("Seria \u2192 PPTX (F7)", "Series \u2192 PPTX (F7)"))
        b_ppt.clicked.connect(lambda: self.studio.tab_ppt.series())
        v.addLayout(hrow(self.b_go, None, b_leg, b_rs, b_ppt))
        self.view_row = QtWidgets.QHBoxLayout()
        lab = QtWidgets.QLabel(T("Poka\u017c na siatce:", "Show on mesh:"))
        lab.setStyleSheet("font-weight:bold")
        self.view_row.addWidget(lab)
        self.view_btns = {}
        vg = QtWidgets.QButtonGroup(self)
        vg.setExclusive(True)
        for m in MQ_ORDER + ["all"]:
            b = QtWidgets.QPushButton(MQ.label(m))
            b.setCheckable(True)
            b.clicked.connect(lambda _=False, m=m: self.show_view(m))
            vg.addButton(b)
            self.view_row.addWidget(b)
            self.view_btns[m] = b
        self.view_row.addStretch(1)
        v.addLayout(self.view_row)
        self.res = QtWidgets.QTextBrowser()
        self.res.setMinimumHeight(150)
        v.addWidget(self.res, 1)
        self.load_from_engine()

    # ---------------------------------------------------------- dane <-> okno
    def load_from_engine(self):
        (self.rb_disp if MQ.scope == "displayed" else self.rb_all).setChecked(True)
        self.cb_2d.setChecked(MQ.dim2)
        self.cb_3d.setChecked(MQ.dim3)
        self.e_prefix.setText(MQ.prefix)
        for m, w in self.w.items():
            c = MQ.cfg[m]
            w["use"].setChecked(bool(c["use"]))
            w["thr_on"].setChecked(bool(c["thr_on"]))
            w["thr"].setText(fmt_num(c["thr"], 6))
            (w["auto"] if c["mode"] == "auto" else w["manual"]).setChecked(True)
            w["nb"].setValue(int(c["nb"]))
        for cb, key in ((self.cb_over, "over_sep"), (self.cb_sets, "make_sets"), (self.cb_worst, "mark_worst"),
                        (self.cb_nice, "nice")):
            cb.setChecked(bool(getattr(MQ, key)))
        self.cb_hide.blockSignals(True)
        self.cb_hide.setChecked(MQ.hide_ok)
        self.cb_hide.blockSignals(False)
        self.c_ok.set_rgb(MQ.col_ok)
        self.c_over.set_rgb(MQ.col_over)
        self.c_na.set_rgb(MQ.col_na)
        self.update_rows()

    def store_to_engine(self):
        MQ.scope = "displayed" if self.rb_disp.isChecked() else "all"
        MQ.dim2, MQ.dim3 = self.cb_2d.isChecked(), self.cb_3d.isChecked()
        MQ.prefix = self.e_prefix.text().strip() or "MQ"
        for m, w in self.w.items():
            c = MQ.cfg[m]
            c["use"] = w["use"].isChecked()
            c["thr_on"] = w["thr_on"].isChecked()
            t = to_float(w["thr"].text())
            if t is not None:
                c["thr"] = t
            c["mode"] = "auto" if w["auto"].isChecked() else "manual"
            c["nb"] = w["nb"].value()
        MQ.over_sep = self.cb_over.isChecked()
        MQ.make_sets = self.cb_sets.isChecked()
        MQ.mark_worst = self.cb_worst.isChecked()
        MQ.nice = self.cb_nice.isChecked()

    def update_rows(self, *a):
        for m, w in self.w.items():
            auto = w["auto"].isChecked()
            w["nb"].setEnabled(auto)
            c = MQ.cfg[m]
            E = manual_edges(c["edges"], c["thr_on"], to_float(c["thr"]), MQ_DEF[m].dir)
            if not auto and len(E) >= 2:
                w["info"].setText(T("%d kolor\u00f3w: %s \u2026 %s", "%d colors: %s \u2026 %s", len(E) - 1, fmt_num(E[0], 4), fmt_num(E[-1], 4)))
            else:
                w["info"].setText(T("z warto\u015bci skrajnych", "from extreme values"))

    # ---------------------------------------------------------- akcje
    def analyze(self):
        self.store_to_engine()
        save_settings()

        def done(_):
            PRESENT.invalidate_frame()
            self.studio.show_legend("mq")
        self.studio.runner.run(MQ.analyze, done)

    def show_view(self, v):
        def work():
            MQ.apply_view(v)
            PRESENT.invalidate_frame()
        self.studio.runner.run(work)

    def restore(self):
        self.studio.runner.run(lambda: (MQ.restore(), PRESENT.invalidate_frame()))

    def refresh(self):
        for m, b in self.view_btns.items():
            on = (m == "all" and MQ.combined_ok()) or (m in MQ.analyzed)
            b.setEnabled(on)
            b.setChecked(m == MQ.view)
        if not MQ.analyzed:
            self.res.setHtml("<i>%s</i>" % h_esc(T("(brak analizy \u2013 zaznacz metryki i kliknij \u201eAnalizuj i koloruj\u201d)",
                                                   "(no analysis \u2013 tick metrics and click \u201cAnalyze and color\u201d)")))
            return
        tot = len(MQ.elems)
        H = ["<div style='color:%s'>%s</div>" % (HDR_COLOR, h_esc(T("Element\u00f3w: %d (2D: %d, 3D: %d), pomini\u0119to innych: %d  \u2022  %s",
                                                                    "Elements: %d (2D: %d, 3D: %d), other skipped: %d  \u2022  %s",
                                                                    tot, MQ.n2d, MQ.n3d, MQ.nskip, MQ.when)))]
        H.append("<table cellspacing=0 cellpadding=3 border=1 style='border-collapse:collapse;border-color:#cfd8e3'>")
        heads = [T("Metryka", "Metric"), T("Pr\u00f3g", "Thr."), T("W normie", "Within"), T("Poza", "Out"), "%", "Min", "Max", T("\u0179r\u00f3d\u0142o", "Source")]
        H.append("<tr style='background:#e8eef5'>%s</tr>" % "".join("<th>%s</th>" % h_esc(h) for h in heads))
        for m in MQ.analyzed:
            s, sc = MQ.stats[m], MQ.scales[m]
            if sc.thr_on:
                thr = "%s %s" % (sc.bad_op(), fmt_num(sc.thr, 3))
                ok, bad, pc = "%d" % s["nOk"], "%d" % s["nBad"], "%.2f" % pct(s["nBad"], tot)
                col = ERR_COLOR if s["nBad"] else OK_COLOR
            else:
                thr = ok = bad = pc = "\u2013"
                col = "#333"
            H.append("<tr><td><b>%s</b></td><td>%s</td><td align=right>%s</td><td align=right style='color:%s'><b>%s</b></td>"
                     "<td align=right>%s</td><td align=right>%s</td><td align=right>%s</td><td>%s</td></tr>"
                     % (h_esc(MQ.label(m)), h_esc(thr), ok, col, bad, pc, fmt_num(s["min"], 4), fmt_num(s["max"], 4), h_esc(MQ.source_text(m))))
        H.append("</table>")
        if MQ.combined_ok():
            s = MQ.stats["all"]
            H.append("<div style='color:%s'>%s</div>" % (HDR_COLOR, h_esc(T("Zbiorczo: w normie we wszystkich metrykach %d, poza norm\u0105 w \u2265 1 metryce %d (%.2f%%)",
                                                                          "Combined: within limits in all metrics %d, out of limits in \u2265 1 metric %d (%.2f%%)",
                                                                          s["nOk"], s["nBad"], pct(s["nBad"], tot)))))
        if MQ.jz_note:
            H.append("<div>Jacobian Zero: %s</div>" % h_esc(MQ.jz_note))
        if MESH.sets:
            H.append("<div>%s</div>" % h_esc(T("Zestawy (sets): %d \u2013 wszystkie metryki naraz (prefiks %sS_)", "Sets: %d \u2013 all metrics at once (prefix %sS_)",
                                               len(MESH.sets), MQ.clean_prefix())))
        self.res.setHtml("".join(H))

    def set_busy(self, on):
        self.b_go.setEnabled(not on)


# =================== GUI: KARTA "DELTA REF / INF" ======================
# Kroki: 1 tryb (delta / jeden model: pasma / jeden model: prog), 2 pliki,
# 3 metryka (grupa radiowa) i typ elementow (OSOBNA grupa radiowa), skala,
# opcje, 4 analiza, widocznosc, "Sprawdz delte" po ID (z danych analizy).
class DeltaTab(_QWidget):
    def __init__(self, studio):
        super(DeltaTab, self).__init__()
        self.studio = studio
        v = QtWidgets.QVBoxLayout(self)
        # --- tryb ---
        self.rb_mode = {}
        mrow = []
        for key, pl, en, tp in (("delta", "Delta REF \u2192 INF (dwa modele)", "Delta REF \u2192 INF (two models)",
                                  T("Pogorszenie metryki INF wzgl\u0119dem REF dla element\u00f3w o tym samym ID (albo przesuni\u0119cie w\u0119z\u0142\u00f3w).",
                                    "Metric worsening of INF vs REF for elements with the same ID (or node displacement).")),
                                 ("single", "Jeden model: pasma wg warto\u015bci", "One model: bands by value",
                                  T("Elementy jednego modelu dzielone na pasma kolor\u00f3w wg warto\u015bci metryki.", "Elements of one model split into color bands by the metric value.")),
                                 ("fail", "Jeden model: podzia\u0142 wg progu", "One model: threshold split",
                                  T("Poza norm\u0105 (czerwony) / w normie (zielony) wg podanego progu.", "Out of limits (red) / within limits (green) by the given threshold."))):
            rb = QtWidgets.QRadioButton(T(pl, en))
            tip(rb, tp)
            rb.toggled.connect(lambda on, k=key: on and self.on_mode(k))
            self.rb_mode[key] = rb
            mrow.append(rb)
        radio_group(self, *mrow)
        v.addWidget(group(T("1  Tryb", "1  Mode"), hrow(*(mrow + [None]))))
        # --- pliki ---
        g = QtWidgets.QGridLayout()
        self.e_ref = QtWidgets.QLineEdit()
        self.e_inf = QtWidgets.QLineEdit()
        b_ref = QtWidgets.QPushButton(T("Wybierz\u2026", "Browse\u2026"))
        b_inf = QtWidgets.QPushButton(T("Wybierz\u2026", "Browse\u2026"))
        b_ref.clicked.connect(lambda: self.pick(self.e_ref, "REF"))
        b_inf.clicked.connect(lambda: self.pick(self.e_inf, "INF"))
        self.lab_ref = QtWidgets.QLabel(T("Plik REF:", "REF file:"))
        self.lab_inf = QtWidgets.QLabel(T("Plik INF:", "INF file:"))
        g.addWidget(self.lab_ref, 0, 0)
        g.addWidget(self.e_ref, 0, 1)
        g.addWidget(b_ref, 0, 2)
        g.addWidget(self.lab_inf, 1, 0)
        g.addWidget(self.e_inf, 1, 1)
        g.addWidget(b_inf, 1, 2)
        self.b_inf = b_inf
        self.cb_open = QtWidgets.QCheckBox(T("INF = aktualnie otwarty model (wtedy wybierasz tylko REF; model jest prze\u0142adowywany z dysku \u2013 zapisz go)",
                                             "INF = currently open model (then choose REF only; the model is reloaded from disk \u2013 save it)"))
        self.cb_refc = QtWidgets.QCheckBox(T("Do\u0142\u0105cz REF jako komponent (nak\u0142adka poka\u017c/ukryj)", "Add REF as a component (show/hide overlay)"))
        self.cb_infc = QtWidgets.QCheckBox(T("Zbierz pozosta\u0142e elementy INF w komponencie", "Collect the remaining INF elements in a component"))
        for i, cb in enumerate((self.cb_open, self.cb_refc, self.cb_infc)):
            g.addWidget(cb, 2 + i, 0, 1, 3)
        self.cb_open.toggled.connect(self.sync)
        v.addWidget(group(T("2  Pliki wej\u015bciowe (.hm) \u2013 wczytanie ZAST\u0118PUJE model w sesji", "2  Input files (.hm) \u2013 loading REPLACES the model in the session"), g))
        # --- parametry ---
        p = QtWidgets.QGridLayout()
        self.rb_metric = {}
        mrow = []
        for key, lab in (("ar", "Aspect Ratio"), ("jac", "Jacobian"), ("skew", "Skewness"), ("disp", T("Przesuni\u0119cie (mm)", "Displacement (mm)"))):
            rb = QtWidgets.QRadioButton(lab)
            rb.toggled.connect(lambda on, k=key: on and self.on_metric(k))
            self.rb_metric[key] = rb
            mrow.append(rb)
        radio_group(self, *mrow)                                   # metryka: wlasna grupa
        p.addWidget(QtWidgets.QLabel(T("Metryka jako\u015bci:", "Quality metric:")), 0, 0)
        p.addLayout(hrow(*(mrow + [None])), 0, 1)
        self.rb_2d = QtWidgets.QRadioButton(T("2D (pow\u0142oki)", "2D (shells)"))
        self.rb_3d = QtWidgets.QRadioButton(T("3D (bry\u0142y)", "3D (solids)"))
        radio_group(self, self.rb_2d, self.rb_3d)                  # typ elementow: OSOBNA grupa (nie odznacza metryki)
        p.addWidget(QtWidgets.QLabel(T("Typ element\u00f3w:", "Element type:")), 1, 0)
        p.addLayout(hrow(self.rb_2d, self.rb_3d, None), 1, 1)
        self.e_prefix = QtWidgets.QLineEdit()
        self.e_target = QtWidgets.QLineEdit()
        p.addWidget(QtWidgets.QLabel(T("Prefiks nazw komponent\u00f3w:", "Component name prefix:")), 2, 0)
        p.addWidget(self.e_prefix, 2, 1)
        p.addWidget(QtWidgets.QLabel(T("Komponent przy przywracaniu (opc.):", "Restore into component (opt.):")), 3, 0)
        p.addWidget(self.e_target, 3, 1)
        tip(self.e_target, T("Puste = elementy wracaj\u0105 do pierwotnych komponent\u00f3w; nazwa = wszystkie do jednego komponentu.",
                             "Empty = elements return to their original components; a name = all into one component."))
        self.lab_fail = QtWidgets.QLabel()
        self.e_fail = QtWidgets.QLineEdit()
        self.e_fail.setMaximumWidth(110)
        p.addWidget(self.lab_fail, 4, 0)
        p.addLayout(hrow(self.e_fail, None), 4, 1)
        p.setColumnStretch(1, 1)
        v.addWidget(group(T("3  Parametry", "3  Parameters"), p))
        # --- skala i opcje ---
        s = QtWidgets.QGridLayout()
        self.cb_auto = QtWidgets.QCheckBox(T("Skala automatyczna (zakres z danych)", "Automatic scale (range from data)"))
        b_scale = QtWidgets.QPushButton(T("R\u0119czna skala (progi pasm)\u2026", "Manual scale (band limits)\u2026"))
        b_scale.clicked.connect(self.edit_scale)
        s.addLayout(hrow(self.cb_auto, b_scale, None), 0, 0, 1, 2)
        self.lab_dead = QtWidgets.QLabel()
        self.e_dead = QtWidgets.QLineEdit()
        self.e_dead.setMaximumWidth(110)
        s.addWidget(self.lab_dead, 1, 0)
        s.addLayout(hrow(self.e_dead, None), 1, 1)
        self.sp_bands = QtWidgets.QSpinBox()
        self.sp_bands.setRange(2, 96)
        self.cb_nice = QtWidgets.QCheckBox(T("zaokr\u0105glaj g\u00f3rny zakres do \u201e\u0142adnej\u201d liczby", "round the scale top to a \u201cnice\u201d number"))
        s.addWidget(QtWidgets.QLabel(T("Liczba pasm gradientu (2\u201396):", "Number of gradient bands (2\u201396):")), 2, 0)
        s.addLayout(hrow(self.sp_bands, 12, self.cb_nice, None), 2, 1)
        self.cb_mark = QtWidgets.QCheckBox(T("Oznacz element MAX i MIN (etykieta + wsp\u00f3\u0142rz\u0119dne)", "Mark the MAX and MIN element (label + coordinates)"))
        self.cb_fade = QtWidgets.QCheckBox(T("Wycisz elementy \u201ebez zmian\u201d (przezroczysto\u015b\u0107 %):", "Fade \u201cno change\u201d elements (transparency %):"))
        self.sp_fade = QtWidgets.QSpinBox()
        self.sp_fade.setRange(0, 100)
        self.cb_imp = QtWidgets.QCheckBox(T("Poka\u017c poprawione elementy (magenta)", "Show improved elements (magenta)"))
        s.addWidget(self.cb_mark, 3, 0, 1, 2)
        s.addLayout(hrow(self.cb_fade, self.sp_fade, None), 4, 0, 1, 2)
        s.addWidget(self.cb_imp, 5, 0, 1, 2)
        s.setColumnStretch(1, 1)
        v.addWidget(group(T("Skala i opcje", "Scale and options"), s))
        # --- akcje ---
        self.b_run = styled_button(T("\u25b6 Wykonaj analiz\u0119", "\u25b6 Run the analysis"), RUN_COLOR, big=True)
        self.b_run.clicked.connect(self.run)
        b_leg = QtWidgets.QPushButton(T("Legenda", "Legend"))
        b_leg.clicked.connect(lambda: self.studio.show_legend("delta"))
        b_rs = QtWidgets.QPushButton(T("Przywr\u00f3\u0107 siatk\u0119", "Restore mesh"))
        b_rs.clicked.connect(lambda: self.studio.runner.run(lambda: (DELTA.restore(), PRESENT.invalidate_frame())))
        b_ppt = QtWidgets.QPushButton(T("Wynik \u2192 slajd PPTX", "Result \u2192 PPTX slide"))
        b_ppt.clicked.connect(lambda: self.studio.tab_ppt.delta_to_pptx())
        v.addLayout(hrow(self.b_run, None, b_leg, b_rs, b_ppt))
        b_tr = QtWidgets.QPushButton(T("REF: poka\u017c/ukryj", "REF: show/hide"))
        b_ti = QtWidgets.QPushButton(T("INF: poka\u017c/ukryj", "INF: show/hide"))
        b_tg = QtWidgets.QPushButton(T("\u201eBez zmian\u201d: poka\u017c/ukryj", "\u201cNo change\u201d: show/hide"))
        b_tr.clicked.connect(lambda: self.studio.runner.run(DELTA.toggle_ref, busy=False))
        b_ti.clicked.connect(lambda: self.studio.runner.run(DELTA.toggle_inf, busy=False))
        b_tg.clicked.connect(lambda: self.studio.runner.run(DELTA.toggle_gray, busy=False))
        self.e_insp = QtWidgets.QLineEdit()
        self.e_insp.setMaximumWidth(90)
        self.e_insp.setPlaceholderText("ID")
        self.e_insp.returnPressed.connect(self.inspect)
        b_insp = QtWidgets.QPushButton(T("Sprawd\u017a delt\u0119 (po ID)", "Check delta (by ID)"))
        b_insp.clicked.connect(self.inspect)
        tip(b_insp, T("Q(REF), Q(INF), \u0394, pogorszenie D i pasmo elementu o podanym ID \u2013 z danych ostatniej analizy (dzia\u0142a niezale\u017cnie od bie\u017c\u0105cych opcji).",
                      "Q(REF), Q(INF), \u0394, worsening D and the band of the element with the given ID \u2013 from the last analysis (independent of the current options)."))
        v.addLayout(hrow(QtWidgets.QLabel(T("Widoczno\u015b\u0107:", "Visibility:")), b_tr, b_ti, b_tg, None,
                         QtWidgets.QLabel(T("Element:", "Element:")), self.e_insp, b_insp))
        self.res = QtWidgets.QTextBrowser()
        self.res.setMinimumHeight(110)
        v.addWidget(self.res, 1)
        self.load_from_engine()

    def load_from_engine(self):
        D = DELTA
        self.rb_mode[D.mode()].setChecked(True)
        self.e_ref.setText(D.ref_file)
        self.e_inf.setText(D.inf_file)
        self.cb_open.setChecked(D.use_open_inf)
        self.cb_refc.setChecked(D.mk_ref_comp)
        self.cb_infc.setChecked(D.mk_inf_comp)
        self.rb_metric[D.metric].setChecked(True)
        (self.rb_3d if D.dim == "3d" else self.rb_2d).setChecked(True)
        self.e_prefix.setText(D.prefix)
        self.e_target.setText(D.restore_target)
        self.cb_auto.setChecked(D.auto_scale)
        self.e_dead.setText(fmt_num(D.deadband, 6))
        self.sp_bands.setValue(int(D.band_count))
        self.cb_nice.setChecked(D.nice_round)
        self.cb_mark.setChecked(D.mark_extremes)
        self.cb_fade.setChecked(D.fade_gray)
        self.sp_fade.setValue(int(D.fade_level))
        self.cb_imp.setChecked(D.show_improved)
        self.e_fail.setText(fmt_num(D.fail_limit, 6))
        self.labels()
        self.sync()

    def store_to_engine(self):
        D = DELTA
        D.ref_file, D.inf_file = self.e_ref.text().strip(), self.e_inf.text().strip()
        D.use_open_inf = self.cb_open.isChecked()
        for k, rb in self.rb_mode.items():
            if rb.isChecked():
                D.set_mode(k)
        D.mk_ref_comp, D.mk_inf_comp = self.cb_refc.isChecked(), self.cb_infc.isChecked()
        for k, rb in self.rb_metric.items():
            if rb.isChecked():
                D.metric = k
        D.dim = "3d" if self.rb_3d.isChecked() else "2d"
        D.prefix = self.e_prefix.text().strip()
        D.restore_target = self.e_target.text().strip()
        D.auto_scale = self.cb_auto.isChecked()
        D.deadband = to_float(self.e_dead.text(), D.deadband)
        D.band_count = self.sp_bands.value()
        D.nice_round, D.mark_extremes = self.cb_nice.isChecked(), self.cb_mark.isChecked()
        D.fade_gray, D.fade_level = self.cb_fade.isChecked(), self.sp_fade.value()
        D.show_improved = self.cb_imp.isChecked()
        D.fail_limit = to_float(self.e_fail.text(), D.fail_limit)

    def labels(self):
        m = DELTA.metric
        self.lab_dead.setText(T("Tolerancja zmiany (mm) (D <):", "Change tolerance (mm) (D <):") if m == "disp"
                              else T("Pr\u00f3g \u201ebez zmian\u201d (|D| <):", "\u201cNo change\u201d threshold (|D| <):"))
        self.lab_fail.setText(T("Poza norm\u0105 gdy warto\u015b\u0107 <:", "Out of limits when value <:") if m == "jac"
                              else T("Poza norm\u0105 gdy warto\u015b\u0107 >:", "Out of limits when value >:"))

    def on_mode(self, key):
        DELTA.set_mode(key)
        self.sync()

    def on_metric(self, key):
        DELTA.metric = key
        DELTA.prefix = self.e_prefix.text().strip()
        DELTA.on_metric_change()
        self.e_prefix.setText(DELTA.prefix)
        self.labels()

    def sync(self, *a):
        mode = DELTA.mode()
        single = mode != "delta"
        inf_off = self.cb_open.isChecked() or single
        for w in (self.e_inf, self.b_inf, self.lab_inf):
            w.setEnabled(not inf_off)
        self.lab_ref.setText(T("Plik modelu:", "Model file:") if single else T("Plik REF:", "REF file:"))
        self.cb_refc.setEnabled(not single)
        self.cb_infc.setEnabled(not single)
        self.rb_metric["disp"].setEnabled(not single)
        if single and self.rb_metric["disp"].isChecked():
            self.rb_metric["ar"].setChecked(True)
        fail = mode == "fail"
        self.lab_fail.setVisible(fail)
        self.e_fail.setVisible(fail)
        for w in (self.cb_auto, self.lab_dead, self.e_dead, self.sp_bands, self.cb_nice):
            w.setEnabled(not fail)
        self.cb_fade.setEnabled(mode == "delta")
        self.sp_fade.setEnabled(mode == "delta")
        self.cb_imp.setEnabled(mode == "delta")

    def pick(self, edit, which):
        start = os.path.dirname(edit.text()) if edit.text() else HM.model_dir()
        p = ask_open_file(self, T("Wybierz plik %s (.hm)", "Choose the %s file (.hm)", which), "HyperMesh (*.hm);;*.*", start)
        if p:
            edit.setText(os.path.normpath(p))

    def run(self):
        self.store_to_engine()
        D = DELTA
        if D.use_open_inf and not HM.model_file():
            p = ask_open_file(self, T("Plik otwartego modelu (INF)", "Open model file (INF)"), "HyperMesh (*.hm)")
            if not p:
                return
            D.use_open_inf = False
            D.inf_file = os.path.normpath(p)
            self.e_inf.setText(D.inf_file)
            self.cb_open.setChecked(False)
        reload_msg = T("Analiza wczyta pliki .hm i ZAST\u0104PI model w sesji HyperMesha (niezapisane zmiany przepadn\u0105).\n\nKontynuowa\u0107?",
                       "The analysis loads the .hm files and REPLACES the model in the HyperMesh session (unsaved changes are lost).\n\nContinue?")
        mw = hm_main_window()
        if D.use_open_inf and mw is not None and ".hm*" in mw.windowTitle():
            reload_msg = T("Otwarty model ma NIEZAPISANE zmiany (*). Analiza wczyta INF z pliku na dysku \u2013 "
                           "zmiany nie zostan\u0105 uwzgl\u0119dnione i przepadn\u0105.\n\nZapisz model (File > Save) i uruchom analiz\u0119 ponownie.\n\n"
                           "Kontynuowa\u0107 mimo to?",
                           "The open model has UNSAVED changes (*). The analysis loads INF from the file on disk \u2013 "
                           "the changes are not included and will be lost.\n\nSave the model (File > Save) and run the analysis again.\n\n"
                           "Continue anyway?")
        loads = bool(D.ref_file) or (bool(D.inf_file) and not D.use_open_inf)
        if loads and not yes_no(self, T("Uwaga", "Warning"), reload_msg):
            BUS.status(T("Anulowano.", "Cancelled."), "warn")
            return
        save_settings()
        self.studio.runner.run(D.run, lambda _: self.done())

    def done(self):
        PRESENT.invalidate_frame()
        lm = DELTA.legend_model()
        if lm:
            c = DELTA.counts
            if DELTA.single:
                msg = T("Gotowe (%s, %s): %d element\u00f3w w pasmach.", "Done (%s, %s): %d elements in bands.", DELTA.dim_label(), DELTA.metric_label(), c["band"])
            else:
                msg = T("Gotowe (%s, %s). W skali: %d | bez zmian: %d | bez odpow.: %d | poprawione: %d",
                        "Done (%s, %s). In scale: %d | no change: %d | unmatched: %d | improved: %d",
                        DELTA.dim_label(), DELTA.metric_label(), c["band"], c["gray"], c["unm"], c["imp"])
            if DELTA.clamped:
                msg += T(" | powy\u017cej zakresu: %d", " | above range: %d", DELTA.clamped)
            if getattr(DELTA, "paint_fail", 0):
                BUS.status(msg + T(" UWAGA: %d grup nie utworzono.", " WARNING: %d groups not created.", DELTA.paint_fail), "warn")
            else:
                BUS.status(msg)
            self.studio.show_legend("delta")

    def refresh(self):
        if not DELTA.done:
            self.res.setHtml("<i>%s</i>" % h_esc(T("(brak wyniku \u2013 wybierz tryb, pliki, metryk\u0119 i typ element\u00f3w, potem \u201eWykonaj analiz\u0119\u201d)",
                                                   "(no result \u2013 choose the mode, files, metric and element type, then \u201cRun the analysis\u201d)")))
            return
        c = DELTA.counts
        rows = [(T("Tryb", "Mode"), {"delta": T("delta REF \u2192 INF", "delta REF \u2192 INF"), "single": T("jeden model: pasma", "one model: bands"),
                                     "fail": T("jeden model: pr\u00f3g", "one model: threshold")}.get("fail" if DELTA.fail_active else ("single" if DELTA.single else "delta"))),
                (T("Metryka / typ", "Metric / type"), "%s / %s" % (DELTA.metric_label(), DELTA.dim_label())),
                (T("Model w sesji", "Model in the session"), HM.model_file() or "-"),
                (T("Analiza", "Analysis"), DELTA.when)]
        if DELTA.fail_active and DELTA.fail_info:
            rows.append((T("Poza norm\u0105 / w normie", "Out of / within limits"), "%d / %d  (%s %s)" % (DELTA.fail_info[0], DELTA.fail_info[1], DELTA.fail_info[3], DELTA.fail_info[2])))
        else:
            rows.append((T("W skali", "In scale"), "%d" % c["band"]))
            if not DELTA.single:
                rows.append((T("Bez zmian / bez odpowiednika / poprawione", "No change / unmatched / improved"), "%d / %d / %d" % (c["gray"], c["unm"], c["imp"])))
        for tag, info in (("MAX", DELTA.max_info), ("MIN", DELTA.min_info)):
            if info:
                xyz = info[2] if len(info) > 2 and info[2] else None
                rows.append(("%s %s" % (tag, DELTA.delta_label()), "%s  (el. %s%s)" % (fmt_num(info[0], 4), info[1],
                                                                                        (", %s (%g, %g, %g)" % ((T("\u015brodek", "center"),) + tuple(xyz))) if xyz else "")))
        if DELTA.skipped[0] or DELTA.skipped[1]:
            rows.append((T("Pomini\u0119te (inny wymiar) REF / INF", "Skipped (other dimension) REF / INF"), "%d / %d" % DELTA.skipped))
        H = ["<table cellspacing=0 cellpadding=3>"]
        for k, val in rows:
            H.append("<tr><td style='color:#555'>%s</td><td><b>%s</b></td></tr>" % (h_esc(k), h_esc(val)))
        H.append("</table>")
        self.res.setHtml("".join(H))

    def edit_scale(self):
        self.store_to_engine()
        if DeltaScaleEditor(self).exec_():
            self.cb_auto.setChecked(False)

    def inspect(self):
        eid = to_float(self.e_insp.text())
        if eid is None:
            BUS.status(T("Podaj ID elementu.", "Enter an element ID."), "err")
            return

        def work():
            txt = DELTA.inspect(int(eid))
            BUS.status(txt.replace("\n", "  |  "))
            msg_box(self, T("Element %d", "Element %d", int(eid)), txt)
        self.studio.runner.run(work, busy=False)

    def set_busy(self, on):
        self.b_run.setEnabled(not on)


# =================== GUI: KARTA "RAPORT JAKOSCI" =======================
# Kroki: 1 zrodlo (otwarty model / plik .hm / porownanie REF vs INF), 2 zakres
# i co sprawdzac (1D / 2D / 3D, zgodnosc ID, topologia, WEZLY, objetosc),
# 3 metryki, 4 plik wynikowy i formaty (TXT / CSV / XLSX / HTML), generuj.
class ReportTab(_QWidget):
    def __init__(self, studio):
        super(ReportTab, self).__init__()
        self.studio = studio
        v = QtWidgets.QVBoxLayout(self)
        # --- zrodlo ---
        g = QtWidgets.QGridLayout()
        self.rb_cur = QtWidgets.QRadioButton(T("Aktualny model (otwarty w HyperMesh)", "Current model (open in HyperMesh)"))
        self.rb_file = QtWidgets.QRadioButton(T("Wczytaj plik .hm", "Load a .hm file"))
        self.rb_cmp = QtWidgets.QRadioButton(T("Por\u00f3wnanie dw\u00f3ch siatek: A = REF vs B = INF (wczyta oba pliki)", "Compare two meshes: A = REF vs B = INF (loads both files)"))
        radio_group(self, self.rb_cur, self.rb_file, self.rb_cmp)
        self.e_file, self.e_ref, self.e_inf = QtWidgets.QLineEdit(), QtWidgets.QLineEdit(), QtWidgets.QLineEdit()
        g.addWidget(self.rb_cur, 0, 0, 1, 3)
        g.addWidget(self.rb_file, 1, 0)
        g.addWidget(self.e_file, 1, 1)
        g.addWidget(self._browse(self.e_file, self.rb_file), 1, 2)
        g.addWidget(self.rb_cmp, 2, 0, 1, 3)
        g.addWidget(QtWidgets.QLabel("REF (.hm):"), 3, 0, QtCore.Qt.AlignRight)
        g.addWidget(self.e_ref, 3, 1)
        g.addWidget(self._browse(self.e_ref, self.rb_cmp), 3, 2)
        g.addWidget(QtWidgets.QLabel("INF (.hm):"), 4, 0, QtCore.Qt.AlignRight)
        g.addWidget(self.e_inf, 4, 1)
        g.addWidget(self._browse(self.e_inf, self.rb_cmp), 4, 2)
        g.setColumnStretch(1, 1)
        v.addWidget(group(T("1  Dane wej\u015bciowe", "1  Input"), g))
        # --- zakres i kontrola ---
        self.rb_all = QtWidgets.QRadioButton(T("Ca\u0142a siatka", "Whole mesh"))
        self.rb_disp = QtWidgets.QRadioButton(T("Tylko wy\u015bwietlone", "Displayed only"))
        radio_group(self, self.rb_all, self.rb_disp)
        self.cb_dim = dict((d, QtWidgets.QCheckBox(lab)) for d, lab in (
            ("1d", T("1D (belki, pr\u0119ty, spr\u0119\u017cyny)", "1D (beams, rods, springs)")), ("2d", T("2D (powierzchniowe)", "2D (shells)")),
            ("3d", T("3D (bry\u0142owe)", "3D (solids)"))))
        self.cb_ids = QtWidgets.QCheckBox(T("Por\u00f3wnanie ID element\u00f3w mi\u0119dzy modelami A i B (typ, warto\u015bci metryk)", "Element ID comparison between models A and B (type, metric values)"))
        self.cb_topo = QtWidgets.QCheckBox(T("Por\u00f3wnanie w\u0119z\u0142\u00f3w (topologii) wsp\u00f3lnych ID element\u00f3w", "Node (topology) comparison of common element IDs"))
        self.cb_nodes = QtWidgets.QCheckBox(T("Zestawienie W\u0118Z\u0141\u00d3W A vs B: wsp\u00f3lne ID, tylko A / B, przesuni\u0119te; tolerancja [mm]:",
                                              "NODE comparison A vs B: common IDs, only A / B, moved; tolerance [mm]:"))
        self.e_ntol = QtWidgets.QLineEdit()
        self.e_ntol.setMaximumWidth(80)
        tip(self.cb_nodes, T("Wsp\u00f3\u0142rz\u0119dne wszystkich w\u0119z\u0142\u00f3w obu modeli s\u0105 por\u00f3wnywane po ID; w\u0119ze\u0142 jest \u201eprzesuni\u0119ty\u201d, gdy |d| > tolerancja. "
                             "Wynik: sekcja w raporcie TXT / HTML i arkusz \u201eW\u0119z\u0142y\u201d w XLSX (lista najbardziej przesuni\u0119tych).",
                             "The coordinates of all nodes of both models are compared by ID; a node is \u201cmoved\u201d when |d| > tolerance. "
                             "Result: a section in the TXT / HTML report and the \u201cNodes\u201d sheet in the XLSX (list of the most displaced)."))
        self.cb_vol = QtWidgets.QCheckBox(T("Obj\u0119to\u015b\u0107 modelu 3D (mm\u00b3) + pole 2D (mm\u00b2)", "3D model volume (mm\u00b3) + 2D area (mm\u00b2)"))
        self.e_tol = QtWidgets.QLineEdit()
        self.e_tol.setMaximumWidth(80)
        c = QtWidgets.QVBoxLayout()
        c.addLayout(hrow(self.rb_all, self.rb_disp, 20, self.cb_dim["1d"], self.cb_dim["2d"], self.cb_dim["3d"], None))
        c.addLayout(hrow(self.cb_ids, None, QtWidgets.QLabel(T("Tolerancja wzgl\u0119dna warto\u015bci:", "Relative value tolerance:")), self.e_tol))
        c.addLayout(hrow(20, self.cb_topo, None))
        c.addLayout(hrow(self.cb_nodes, self.e_ntol, None))
        c.addWidget(self.cb_vol)
        c.addWidget(note_label(T("Por\u00f3wnania ID, topologii i w\u0119z\u0142\u00f3w dzia\u0142aj\u0105 w trybie por\u00f3wnania dw\u00f3ch modeli (A = REF, B = INF).",
                                 "ID, topology and node comparisons work in the two-model comparison mode (A = REF, B = INF).")))
        v.addWidget(group(T("2  Zakres i co sprawdza\u0107", "2  Scope and what to check"), c))
        # --- metryki ---
        mg = QtWidgets.QGridLayout()
        self.cb_m = {}
        for i, k in enumerate(REP_ORDER):
            cb = QtWidgets.QCheckBox(REP[k].label)
            cb.setToolTip("%s \u2022 %s" % (REP[k].desc, data_name(k)))
            self.cb_m[k] = cb
            mg.addWidget(cb, i // 3, i % 3)
        v.addWidget(group(T("3  Metryki", "3  Metrics"), mg))
        # --- plik i formaty ---
        self.e_out = QtWidgets.QLineEdit()
        b_out = QtWidgets.QPushButton(T("Przegl\u0105daj\u2026", "Browse\u2026"))
        b_out.clicked.connect(self.browse_out)
        self.cb_fmt = dict((f, QtWidgets.QCheckBox(lab)) for f, lab in (
            ("txt", T("TXT (czytelny)", "TXT (readable)")), ("csv", "CSV"), ("xlsx", "XLSX (Excel)"), ("html", T("HTML (interaktywny)", "HTML (interactive)"))))
        self.cb_elem = QtWidgets.QCheckBox(T("Tabela per element (CSV/XLSX \u2013 mo\u017ce by\u0107 du\u017cy plik)", "Per-element table (CSV/XLSX \u2013 may be a big file)"))
        self.cb_comp = QtWidgets.QCheckBox(T("Podzia\u0142 per komponent", "Per-component breakdown"))
        self.rb_pl = QtWidgets.QRadioButton("PL (1,23 ;)")
        self.rb_en = QtWidgets.QRadioButton("EN (1.23 ,)")
        radio_group(self, self.rb_pl, self.rb_en)
        self.rb_full = QtWidgets.QRadioButton(T("pe\u0142ny", "full"))
        self.rb_basic = QtWidgets.QRadioButton(T("podstawowy", "basic"))
        radio_group(self, self.rb_full, self.rb_basic)
        f = QtWidgets.QVBoxLayout()
        f.addLayout(hrow(QtWidgets.QLabel(T("Plik (baza nazwy):", "File (base name):")), self.e_out, b_out))
        f.addLayout(hrow(*([self.cb_fmt[k] for k in ("txt", "csv", "xlsx", "html")] + [None])))
        f.addLayout(hrow(self.cb_elem, self.cb_comp, None))
        f.addLayout(hrow(QtWidgets.QLabel(T("Liczby w CSV:", "Numbers in CSV:")), self.rb_pl, self.rb_en, 20,
                         QtWidgets.QLabel(T("Zawarto\u015b\u0107 raportu:", "Report content:")), self.rb_full, self.rb_basic, None))
        v.addWidget(group(T("4  Plik wynikowy i formaty", "4  Output file and formats"), f))
        self.b_go = styled_button(T("\u25b6 Generuj raport", "\u25b6 Generate report"), GO_COLOR, big=True)
        self.b_go.clicked.connect(self.generate)
        b_open = QtWidgets.QPushButton(T("Otw\u00f3rz folder", "Open folder"))
        b_open.clicked.connect(lambda: open_path(os.path.dirname(REPORT.out_file) or "."))
        b_html = QtWidgets.QPushButton(T("Otw\u00f3rz HTML", "Open HTML"))
        b_html.clicked.connect(self.open_html)
        b_ppt = QtWidgets.QPushButton(T("Raport \u2192 slajd PPTX", "Report \u2192 PPTX slide"))
        b_ppt.clicked.connect(lambda: self.studio.tab_ppt.report_to_pptx())
        v.addLayout(hrow(self.b_go, None, b_open, b_html, b_ppt))
        self.res = QtWidgets.QTextBrowser()
        self.res.setMinimumHeight(100)
        v.addWidget(self.res, 1)
        self.load_from_engine()

    def _browse(self, edit, rb):
        b = QtWidgets.QPushButton(T("Przegl\u0105daj\u2026", "Browse\u2026"))

        def pick():
            p = ask_open_file(self, T("Wybierz plik .hm", "Choose a .hm file"), "HyperMesh (*.hm);;*.*",
                              os.path.dirname(edit.text()) if edit.text() else "")
            if p:
                edit.setText(os.path.normpath(p))
                rb.setChecked(True)
        b.clicked.connect(pick)
        return b

    def browse_out(self):
        start = self.e_out.text() or os.path.join(HM.model_dir() or "", "%s_jakosc.txt" % (HM.model_name() if HM.ok() else "raport"))
        p = ask_save_file(self, T("Zapisz raport jako", "Save report as"), T("Tekst (*.txt);;Wszystkie (*.*)", "Text (*.txt);;All (*.*)"), start)
        if p:
            self.e_out.setText(os.path.normpath(p))

    def open_html(self):
        htmls = [f for f in REPORT.last_files if f.lower().endswith(".html")]
        if htmls:
            open_path(htmls[-1])
        else:
            BUS.status(T("Brak raportu HTML z ostatniego przebiegu.", "No HTML report from the last run."), "warn")

    def load_from_engine(self):
        R = REPORT
        {"file": self.rb_file, "compare": self.rb_cmp}.get(R.src, self.rb_cur).setChecked(True)
        self.e_file.setText(R.hm_file)
        self.e_ref.setText(R.ref_file)
        self.e_inf.setText(R.inf_file)
        (self.rb_disp if R.scope == "displayed" else self.rb_all).setChecked(True)
        for d, cb in self.cb_dim.items():
            cb.setChecked(bool(R.use_dim.get(d)))
        self.cb_ids.setChecked(R.chk_ids)
        self.cb_topo.setChecked(R.chk_topo)
        self.cb_nodes.setChecked(R.chk_nodes)
        self.e_ntol.setText("%g" % R.ntol())
        self.cb_vol.setChecked(R.chk_vol)
        self.e_tol.setText("%g" % R.tol())
        for k, cb in self.cb_m.items():
            cb.setChecked(bool(R.use_metric.get(k)))
        self.e_out.setText(R.out_file)
        for k, cb in self.cb_fmt.items():
            cb.setChecked(bool(R.fmt.get(k)))
        self.cb_elem.setChecked(R.per_elem)
        self.cb_comp.setChecked(R.per_comp)
        (self.rb_en if R.num_style == "en" else self.rb_pl).setChecked(True)
        (self.rb_basic if R.level == "basic" else self.rb_full).setChecked(True)

    def store_to_engine(self):
        R = REPORT
        R.src = "file" if self.rb_file.isChecked() else "compare" if self.rb_cmp.isChecked() else "current"
        R.hm_file, R.ref_file, R.inf_file = self.e_file.text().strip(), self.e_ref.text().strip(), self.e_inf.text().strip()
        R.scope = "displayed" if self.rb_disp.isChecked() else "all"
        R.use_dim = dict((d, cb.isChecked()) for d, cb in self.cb_dim.items())
        R.chk_ids, R.chk_topo, R.chk_vol = self.cb_ids.isChecked(), self.cb_topo.isChecked(), self.cb_vol.isChecked()
        R.chk_nodes = self.cb_nodes.isChecked()
        t = to_float(self.e_tol.text())
        if t is None or t < 0:
            raise ValueError(T("Tolerancja musi by\u0107 liczb\u0105 dodatni\u0105 (np. 1e-6).", "The tolerance must be a positive number (e.g. 1e-6)."))
        R.cmp_tol = t
        nt = to_float(self.e_ntol.text())
        if nt is None or nt < 0:
            raise ValueError(T("Tolerancja w\u0119z\u0142\u00f3w musi by\u0107 liczb\u0105 dodatni\u0105 w mm (np. 0.0001).", "The node tolerance must be a positive number in mm (e.g. 0.0001)."))
        R.node_tol = nt
        R.use_metric = dict((k, cb.isChecked()) for k, cb in self.cb_m.items())
        R.out_file = self.e_out.text().strip()
        R.fmt = dict((k, cb.isChecked()) for k, cb in self.cb_fmt.items())
        R.per_elem, R.per_comp = self.cb_elem.isChecked(), self.cb_comp.isChecked()
        R.num_style = "en" if self.rb_en.isChecked() else "pl"
        R.level = "basic" if self.rb_basic.isChecked() else "full"

    def generate(self):
        try:
            self.store_to_engine()
        except ValueError as e:
            BUS.status("%s" % e, "err")
            return
        if not REPORT.out_file:
            self.browse_out()
            REPORT.out_file = self.e_out.text().strip()
            if not REPORT.out_file:
                return
        if REPORT.src in ("file", "compare"):
            msg = (T("Por\u00f3wnanie wczyta kolejno modele REF i INF, ZAST\u0118PUJ\u0104C model w sesji HyperMesha (na ko\u0144cu zostanie INF).\n\nKontynuowa\u0107?",
                     "The comparison loads REF and then INF, REPLACING the model in the HyperMesh session (INF stays loaded).\n\nContinue?")
                   if REPORT.src == "compare" else
                   T("Otwarcie pliku .hm ZAST\u0104PI bie\u017c\u0105cy model w sesji HyperMesha.\n\nKontynuowa\u0107?",
                     "Opening the .hm file REPLACES the current model in the HyperMesh session.\n\nContinue?"))
            if not yes_no(self, T("Uwaga", "Warning"), msg):
                return
            if MESH.active():
                MQ.restore() if MESH.owner == "mq" else DELTA.restore()
        save_settings()
        BUS.log("-" * 40)

        def done(msg):
            BUS.log(msg)
            BUS.status(msg.splitlines()[0])
            PRESENT.invalidate_frame()
            self.refresh()
            msg_box(self, T("Gotowe", "Done"), msg)
        self.studio.runner.run(REPORT.generate, done)

    def refresh(self):
        run = REPORT.last
        if not run:
            self.res.setHtml("<i>%s</i>" % h_esc(T("(brak raportu \u2013 wybierz \u017ar\u00f3d\u0142o, metryki i plik, potem \u201eGeneruj raport\u201d)",
                                                   "(no report \u2013 choose the source, metrics and file, then \u201cGenerate report\u201d)")))
            return
        H = ["<div style='color:%s'><b>%s</b> %s \u2022 %s</div>" % (HDR_COLOR, h_esc(T("Ostatni raport:", "Last report:")), h_esc(run["when"]),
                                                                     h_esc(T("element\u00f3w: %d (%s)", "elements: %d (%s)", run["nsel"], run["dims"])))]
        sc = run["score"]
        if sc["has"]:
            H.append("<div>%s</div>" % h_esc(T("Wska\u017anik jako\u015bci: %.1f / 100 (ocena %s)", "Quality score: %.1f / 100 (grade %s)", sc["score"], sc["grade"])))
        nc = REPORT.last_ncmp
        if nc and nc["on"]:
            H.append("<div>%s</div>" % h_esc(T("W\u0119z\u0142y A vs B: wsp\u00f3lne %d, tylko A %d, tylko B %d, przesuni\u0119te %d, maks. %s mm",
                                               "Nodes A vs B: common %d, only A %d, only B %d, moved %d, max %s mm",
                                               nc["common"], nc["onlyA"], nc["onlyB"], nc["moved"], fmt_num(nc["dmax"], 4))))
        cm = REPORT.last_cmp
        if cm and cm["on"]:
            H.append("<div>%s</div>" % h_esc(T("Elementy A vs B: wsp\u00f3lne %d, tylko A %d, tylko B %d, r\u00f3\u017cni\u0105ce si\u0119 %d",
                                               "Elements A vs B: common %d, only A %d, only B %d, differing %d", cm["common"], cm["onlyA"], cm["onlyB"], cm["difftot"])))
        if REPORT.last_files:
            H.append("<div>%s<br>%s</div>" % (h_esc(T("Pliki:", "Files:")), "<br>".join(h_esc(f) for f in REPORT.last_files)))
        self.res.setHtml("".join(H))

    def set_busy(self, on):
        self.b_go.setEnabled(not on)


# =================== GUI: KARTA "WIDOKI I ZRZUTY" ======================
# Lista kafelkow (checkbox, miniatura, nazwa + dane VS, przyciski Kopiuj /
# Testuj / Miniatura / Edytuj / Usun), zapamietywanie widoku (klawisz 1),
# nagrywanie ukladu, raport / sesja HTML, eksport zrzutow (klawisz K).
class ViewTile(_QWidget):
    def __init__(self, tab, rec):
        super(ViewTile, self).__init__()
        self.tab, self.rec = tab, rec
        self.setObjectName("tile")
        self.setStyleSheet("#tile{background:white;border:1px solid #c8d0da;border-radius:3px}")
        h = QtWidgets.QHBoxLayout(self)
        h.setContentsMargins(6, 4, 6, 4)
        self.cb = QtWidgets.QCheckBox()
        self.cb.setChecked(rec["sel"])
        self.cb.toggled.connect(self.toggled)
        h.addWidget(self.cb, 0, QtCore.Qt.AlignTop)
        img = QtWidgets.QLabel()
        img.setFixedSize(THUMB_W // 2 + 4, THUMB_H // 2 + 4)
        img.setAlignment(QtCore.Qt.AlignCenter)
        img.setStyleSheet("background:#e8ecf0;border:1px solid #bbb")
        if rec.get("thumb"):
            pm = QtGui.QPixmap()
            pm.loadFromData(rec["thumb"])
            img.setPixmap(pm.scaled(THUMB_W // 2, THUMB_H // 2, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation))
        else:
            img.setText(T("(brak podgl\u0105du)", "(no preview)"))
        h.addWidget(img)
        info = QtWidgets.QVBoxLayout()
        nm = rec["name"] + ("  \u2699" if rec.get("disp") else "") + ("  \u25c9" if ViewStore.shot_ok(rec) else "")
        lab = QtWidgets.QLabel(nm)
        lab.setStyleSheet("font-weight:bold")
        vs = QtWidgets.QLineEdit(view_text(rec["view"]))
        vs.setReadOnly(True)
        vs.setCursorPosition(0)                # pokaz poczatek "VS: ...", nie koniec
        vs.setStyleSheet("font-family:Consolas,monospace;font-size:8pt")
        info.addWidget(lab)
        info.addWidget(vs)
        info.addStretch(1)
        h.addLayout(info, 1)
        btns = QtWidgets.QVBoxLayout()
        btns.setSpacing(2)
        for text, fn in ((T("Kopiuj", "Copy"), tab.copy_view), (T("Testuj", "Test"), tab.test_view),
                         (T("Miniatura", "Thumbnail"), tab.refresh_thumb), (T("Edytuj", "Edit"), tab.edit_view),
                         (T("Usu\u0144", "Delete"), tab.delete_view)):
            b = QtWidgets.QPushButton(text)
            b.setMinimumWidth(84)
            b.clicked.connect(lambda _=False, fn=fn: fn(rec["id"]))
            btns.addWidget(b)
        h.addLayout(btns)

    def toggled(self, on):
        self.rec["sel"] = on
        self.tab.update_count()
        self.tab.studio.preview_changed()


class ViewEditDialog(_QDialog):
    def __init__(self, parent, rec):
        super(ViewEditDialog, self).__init__(parent)
        self.setWindowTitle(T("Edytuj widok", "Edit view"))
        v = QtWidgets.QVBoxLayout(self)
        self.e_name = QtWidgets.QLineEdit(rec["name"])
        self.t_view = QtWidgets.QPlainTextEdit(view_text(rec["view"]))
        self.t_disp = QtWidgets.QPlainTextEdit(rec.get("disp", ""))
        for w in (self.t_view, self.t_disp):
            w.setStyleSheet("font-family:Consolas,monospace;font-size:8pt")
        v.addWidget(QtWidgets.QLabel(T("Nazwa:", "Name:")))
        v.addWidget(self.e_name)
        v.addWidget(QtWidgets.QLabel(T("Dane widoku (VS: / MTX: albo \u2265 16 liczb):", "View data (VS: / MTX: or \u2265 16 numbers):")))
        v.addWidget(self.t_view)
        v.addWidget(note_label(T("Uk\u0142ad wy\u015bwietlania \u2013 komendy HM odtwarzane przed zrzutem (sekcje, style, maskowanie; mo\u017cna wklei\u0107 z command.tcl):",
                                 "Display layout \u2013 HM commands replayed before the shot (sections, styles, masking; may be pasted from command.tcl):")))
        v.addWidget(self.t_disp)
        ok = styled_button(T("Zapisz", "Save"), HDR_COLOR)
        ok.clicked.connect(self.accept)
        c = QtWidgets.QPushButton(T("Anuluj", "Cancel"))
        c.clicked.connect(self.reject)
        v.addLayout(hrow(None, ok, c))
        self.resize(560, 420)


class ViewsTab(_QWidget):
    def __init__(self, studio):
        super(ViewsTab, self).__init__()
        self.studio = studio
        v = QtWidgets.QVBoxLayout(self)
        self.b_rem = styled_button(T("\u25cf Zapami\u0119taj widok (1)", "\u25cf Remember view (1)"), HDR_COLOR)
        self.b_rem.clicked.connect(self.remember)
        tip(self.b_rem, T("Zapisuje kamer\u0119 (orientacja, zoom, pan) i PE\u0141N\u0104 klatk\u0119 okna graficznego (WYSIWYG) z miniatur\u0105.",
                          "Stores the camera (orientation, zoom, pan) and the FULL graphics-window frame (WYSIWYG) with a thumbnail."))
        b_all = QtWidgets.QPushButton(T("Zaznacz wszystkie", "Select all"))
        b_none = QtWidgets.QPushButton(T("Odznacz wszystkie", "Deselect all"))
        b_all.clicked.connect(lambda: self.set_all(True))
        b_none.clicked.connect(lambda: self.set_all(False))
        self.lab_cnt = QtWidgets.QLabel()
        v.addLayout(hrow(self.b_rem, b_all, b_none, None, self.lab_cnt))
        b_add = QtWidgets.QPushButton(T("Dodaj z tekstu\u2026", "Add from text\u2026"))
        b_add.clicked.connect(self.add_text)
        b_rep = QtWidgets.QPushButton(T("Zapisz raport\u2026", "Save report\u2026"))
        b_rep.clicked.connect(self.save_report)
        b_load = QtWidgets.QPushButton(T("Wczytaj sesj\u0119\u2026", "Load session\u2026"))
        b_load.clicked.connect(self.load_session)
        self.b_rec = QtWidgets.QPushButton()
        self.b_rec.clicked.connect(self.toggle_record)
        b_clr = QtWidgets.QPushButton(T("Wyczy\u015b\u0107", "Clear"))
        b_clr.clicked.connect(self.clear)
        v.addLayout(hrow(b_add, b_rep, b_load, 12, self.b_rec, 12, b_clr, None))
        self.list_w = QtWidgets.QWidget()
        self.list_lay = QtWidgets.QVBoxLayout(self.list_w)
        self.list_lay.setSpacing(4)
        sa = QtWidgets.QScrollArea()
        sa.setWidget(self.list_w)
        sa.setWidgetResizable(True)
        sa.setMinimumHeight(220)
        sa.setStyleSheet("QScrollArea{background:#f0f0f0}")
        self.scroll = sa
        v.addWidget(sa, 1)
        # --- zapis ---
        g = QtWidgets.QGridLayout()
        self.e_dir = QtWidgets.QLineEdit()
        b_dir = QtWidgets.QPushButton(T("Wybierz\u2026", "Browse\u2026"))
        b_dir.clicked.connect(self.pick_dir)
        self.cb_mdir = QtWidgets.QCheckBox(T("Zapisz do folderu otwartego modelu (.hm)", "Save to the open model's folder (.hm)"))
        self.e_pfx = QtWidgets.QLineEdit()
        self.e_pfx.setMaximumWidth(120)
        self.cb_sub = QtWidgets.QCheckBox(T("Utw\u00f3rz podfolder o nazwie prefiksu", "Create a subfolder named after the prefix"))
        self.cmb_fmt = QtWidgets.QComboBox()
        self.cmb_fmt.addItems(["png", "jpg", "bmp"])
        self.cmb_bg = QtWidgets.QComboBox()
        for key, pl, en, hx in BG_MODES:
            self.cmb_bg.addItem(T(pl, en), key)
        b_bg = QtWidgets.QPushButton(T("Testuj t\u0142o", "Test background"))
        b_bg.clicked.connect(self.test_bg)
        self.cb_ws = QtWidgets.QCheckBox(T("Eksportuj klatki zapisane przy \u201eZapami\u0119taj widok\u201d (WYSIWYG)", "Export the frames saved at \u201cRemember view\u201d (WYSIWYG)"))
        self.cb_rep = QtWidgets.QCheckBox(T("Raport HTML (samodzielny, z podgl\u0105dami; mo\u017cna go wczyta\u0107 jako sesj\u0119)", "HTML report (standalone, with previews; can be loaded as a session)"))
        self.lab_pfx = QtWidgets.QLabel(T("Prefiks plik\u00f3w:", "File prefix:"))
        g.addWidget(QtWidgets.QLabel(T("Folder zrzut\u00f3w:", "Screenshot folder:")), 0, 0)
        g.addWidget(self.e_dir, 0, 1, 1, 2)
        g.addWidget(b_dir, 0, 3)
        g.addWidget(self.cb_mdir, 1, 0, 1, 4)
        g.addWidget(self.lab_pfx, 2, 0)
        g.addWidget(self.e_pfx, 2, 1)
        g.addWidget(self.cb_sub, 2, 2, 1, 2)
        g.addWidget(QtWidgets.QLabel(T("Format:", "Format:")), 3, 0)
        g.addWidget(self.cmb_fmt, 3, 1)
        g.addLayout(hrow(QtWidgets.QLabel(T("T\u0142o zrzutu:", "Screenshot background:")), self.cmb_bg, b_bg, None), 3, 2, 1, 2)
        g.addWidget(self.cb_ws, 4, 0, 1, 4)
        g.addWidget(self.cb_rep, 5, 0, 1, 4)
        g.setColumnStretch(2, 1)
        v.addWidget(group(T("Zapis", "Output"), g))
        # --- nazwy automatyczne ---
        ng = QtWidgets.QGridLayout()
        self.cb_pre = {}
        for i, p in enumerate(VIEW_PRESETS):
            cb = QtWidgets.QCheckBox(p)
            cb.toggled.connect(self.update_names)
            self.cb_pre[p] = cb
            ng.addWidget(cb, i % 3, i // 3)
        ng.addWidget(note_label(T("Zaznaczone nazwy trafiaj\u0105 do kolejnych zaznaczonych widok\u00f3w (REF AR/Jac/Skew, potem INF). Pole prefiksu jest wtedy blokowane, a podfolder dostaje pierwsz\u0105 zaznaczon\u0105 nazw\u0119.",
                                  "Ticked names go to the selected views in order (REF AR/Jac/Skew, then INF). The prefix field is then locked and the subfolder takes the first ticked name.")), 3, 0, 1, 2)
        v.addWidget(group(T("Automatyczne nazwy plik\u00f3w", "Automatic file names"), ng))
        self.b_gen = styled_button(T("\u25b6 Generuj zrzuty (K)", "\u25b6 Generate screenshots (K)"), GO_COLOR, big=True)
        self.b_gen.clicked.connect(self.generate)
        b_ppt = QtWidgets.QPushButton(T("Widoki \u2192 slajdy PPTX", "Views \u2192 PPTX slides"))
        b_ppt.clicked.connect(lambda: self.studio.tab_ppt.deliver(PRESENT.shots_items))
        v.addLayout(hrow(self.b_gen, None, b_ppt))
        self.load_from_engine()
        self.redraw()

    # ---------------------------------------------------------- dane <-> okno
    def load_from_engine(self):
        V = VIEWS
        self.e_dir.setText(V.shot_dir)
        self.cb_mdir.setChecked(V.model_dir)
        self.e_pfx.setText(V.prefix)
        self.cb_sub.setChecked(V.subfolder)
        self.cmb_fmt.setCurrentText(V.fmt)
        keys = [k for k, _, _, _ in BG_MODES]
        self.cmb_bg.setCurrentIndex(keys.index(V.bg) if V.bg in keys else 0)
        self.cb_ws.setChecked(V.use_stored)
        self.cb_rep.setChecked(V.report)
        for p, cb in self.cb_pre.items():
            cb.setChecked(bool(V.presets.get(p)))
        self.update_names()
        self.sync_rec()

    def store_to_engine(self):
        V = VIEWS
        V.shot_dir = self.e_dir.text().strip()
        V.model_dir = self.cb_mdir.isChecked()
        V.prefix = self.e_pfx.text().strip() or "shot"
        V.subfolder = self.cb_sub.isChecked()
        V.fmt = self.cmb_fmt.currentText()
        V.bg = self.cmb_bg.currentData() or "none"
        V.use_stored = self.cb_ws.isChecked()
        V.report = self.cb_rep.isChecked()
        V.presets = dict((p, cb.isChecked()) for p, cb in self.cb_pre.items())

    def update_names(self, *a):
        on = any(cb.isChecked() for cb in self.cb_pre.values())
        self.e_pfx.setEnabled(not on)
        self.lab_pfx.setEnabled(not on)

    def sync_rec(self):
        if VIEWS.rec_on:
            self.b_rec.setText(T("\u25a0 Zako\u0144cz nagrywanie", "\u25a0 Stop recording"))
            self.b_rec.setStyleSheet("QPushButton{background:#b00020;color:white;font-weight:bold;padding:3px 10px}")
        else:
            self.b_rec.setText(T("\u25cf Nagrywaj uk\u0142ad", "\u25cf Record layout"))
            self.b_rec.setStyleSheet("")

    # ---------------------------------------------------------- lista
    def redraw(self):
        clear_layout(self.list_lay)
        if not VIEWS.views:
            lab = QtWidgets.QLabel(T("(brak widok\u00f3w \u2013 ustaw model w oknie graficznym i kliknij \u201eZapami\u0119taj widok\u201d albo naci\u015bnij 1)",
                                     "(no views \u2013 set up the model in the graphics window and click \u201cRemember view\u201d or press 1)"))
            lab.setStyleSheet("color:#777;padding:14px")
            self.list_lay.addWidget(lab)
        for r in VIEWS.views:
            self.list_lay.addWidget(ViewTile(self, r))
        self.list_lay.addStretch(1)
        self.update_count()

    def update_count(self):
        self.lab_cnt.setText(T("Zaznaczono: %d / %d", "Selected: %d / %d", len(VIEWS.selected()), len(VIEWS.views)))

    def scroll_end(self):
        QtCore.QTimer.singleShot(50, lambda: self.scroll.verticalScrollBar().setValue(self.scroll.verticalScrollBar().maximum()))

    # ---------------------------------------------------------- akcje
    def remember(self):
        self.store_to_engine()

        def done(r):
            self.redraw()
            self.scroll_end()
            extra = T(" + uk\u0142ad (%d kom.)", " + layout (%d cmds)", len(r["disp"].splitlines())) if r["disp"] else ""
            BUS.status(T("Zapami\u0119tano widok%s. Widok\u00f3w: %d.", "View remembered%s. Views: %d.", extra, len(VIEWS.views)))
        self.studio.runner.run(VIEWS.remember, done)

    def set_all(self, on):
        for r in VIEWS.views:
            r["sel"] = on
        self.redraw()
        self.studio.preview_changed()

    def copy_view(self, vid):
        r = VIEWS.rec(vid)
        if r:
            app_instance().clipboard().setText(view_text(r["view"]))
            BUS.status(T("Skopiowano dane widoku (VS) do schowka.", "View data (VS) copied to the clipboard."))

    def test_view(self, vid):
        self.studio.runner.run(lambda: BUS.status(VIEWS.test(vid)), busy=False)

    def refresh_thumb(self, vid):
        self.store_to_engine()

        def done(ok):
            self.redraw()
            BUS.status(T("Od\u015bwie\u017cono miniatur\u0119 i klatk\u0119 widoku.", "View thumbnail and frame refreshed.") if ok else
                       T("Nie uda\u0142o si\u0119 zrobi\u0107 miniatury.", "Could not create the thumbnail."), "ok" if ok else "err")
        self.studio.runner.run(lambda: VIEWS.refresh_thumb(vid), done)

    def edit_view(self, vid):
        r = VIEWS.rec(vid)
        if not r:
            return
        d = ViewEditDialog(self, r)
        if d.exec_():
            if VIEWS.edit(vid, d.e_name.text(), d.t_view.toPlainText(), d.t_disp.toPlainText()):
                BUS.status(T("Zaktualizowano \u201e%s\u201d.", "Updated \u201c%s\u201d.", r["name"]))
            else:
                BUS.status(T("Niepoprawne dane widoku (potrzeba VS: / MTX: albo \u2265 16 liczb).", "Invalid view data (need VS: / MTX: or \u2265 16 numbers)."), "err")
            self.redraw()

    def delete_view(self, vid):
        VIEWS.delete(vid)
        self.redraw()
        BUS.status(T("Usuni\u0119to widok. Pozosta\u0142o: %d.", "View deleted. Remaining: %d.", len(VIEWS.views)))

    def clear(self):
        if VIEWS.views and yes_no(self, T("Wyczy\u015b\u0107", "Clear"), T("Usun\u0105\u0107 wszystkie zapami\u0119tane widoki?", "Delete all remembered views?")):
            VIEWS.clear()
            self.redraw()

    def add_text(self):
        d = QtWidgets.QDialog(self)
        d.setWindowTitle(T("Dodaj widok z tekstu", "Add view from text"))
        lay = QtWidgets.QVBoxLayout(d)
        lay.addWidget(note_label(T("Wklej dane widoku/widok\u00f3w (1 linia = 1 widok): VS: / MTX: albo \u2265 16 liczb. Bez miniatury \u2013 mo\u017cna j\u0105 potem od\u015bwie\u017cy\u0107.",
                                   "Paste view data (1 line = 1 view): VS: / MTX: or \u2265 16 numbers. No thumbnail \u2013 refresh it later.")))
        t = QtWidgets.QPlainTextEdit()
        t.setStyleSheet("font-family:Consolas,monospace;font-size:8pt")
        lay.addWidget(t)
        ok = styled_button(T("Dodaj", "Add"), HDR_COLOR)
        ok.clicked.connect(d.accept)
        c = QtWidgets.QPushButton(T("Anuluj", "Cancel"))
        c.clicked.connect(d.reject)
        lay.addLayout(hrow(None, ok, c))
        d.resize(560, 260)
        if d.exec_():
            n = VIEWS.add_from_text(t.toPlainText())
            self.redraw()
            BUS.status(T("Dodano widok\u00f3w: %d.", "Views added: %d.", n) if n else T("Nie dodano nic \u2013 brak poprawnych linii.", "Nothing added \u2013 no valid lines."), "ok" if n else "err")

    def save_report(self):
        self.store_to_engine()
        d = VIEWS.shot_dir or ask_dir(self, T("Folder na raport widok\u00f3w", "Folder for the views report"))
        if not d:
            return
        self.studio.runner.run(lambda: BUS.status(T("Zapisano raport: %s", "Report saved: %s", VIEWS.save_report(d))))

    def load_session(self):
        p = ask_open_file(self, T("Wczytaj sesj\u0119 (HTML)", "Load session (HTML)"), "HTML (*.html *.htm)")
        if not p:
            return
        if VIEWS.views and not yes_no(self, T("Wczytaj sesj\u0119", "Load session"),
                                      T("Wczytanie zast\u0105pi obecn\u0105 list\u0119 (%d widok\u00f3w). Kontynuowa\u0107?", "Loading replaces the current list (%d views). Continue?", len(VIEWS.views))):
            return

        def work():
            n, nd = VIEWS.load_session(p)
            self.redraw()
            BUS.status(T("Wczytano sesj\u0119: %d widok\u00f3w z %s (z uk\u0142adem: %d).", "Session loaded: %d views from %s (with layout: %d).",
                         n, os.path.basename(p), nd))
        self.studio.runner.run(work)

    def toggle_record(self):
        if not VIEWS.rec_on and not VIEWS.toggle_record():
            p = ask_open_file(self, T("Wska\u017c plik komend HM (command.tcl)", "Locate the HM command file (command.tcl)"), "command (*.tcl *.cmf);;*.*", os.getcwd())
            if not p or not VIEWS.toggle_record(p):
                BUS.status(T("Nie znaleziono pliku komend \u2013 uk\u0142ad mo\u017cna wklei\u0107 r\u0119cznie w \u201eEdytuj\u201d.", "Command file not found \u2013 paste the layout manually in \u201cEdit\u201d."), "err")
                return
        elif VIEWS.rec_on:
            VIEWS.toggle_record()
        self.sync_rec()
        BUS.status(T("NAGRYWAM uk\u0142ad: wykonaj sekcje / style / maskowanie, potem \u201eZapami\u0119taj widok\u201d.", "RECORDING layout: set up sections / styles / masking, then \u201cRemember view\u201d.")
                   if VIEWS.rec_on else T("Nagrywanie uk\u0142adu zako\u0144czone.", "Layout recording stopped."), "warn" if VIEWS.rec_on else "ok")

    def pick_dir(self):
        d = ask_dir(self, T("Folder na zrzuty", "Screenshot folder"), self.e_dir.text())
        if d:
            self.e_dir.setText(os.path.normpath(d))

    def test_bg(self):
        self.store_to_engine()
        hx = bg_hex(VIEWS.bg)
        if not hx:
            BUS.status(T("Wybrano \u201ebez zmian\u201d \u2013 t\u0142o nie b\u0119dzie zmieniane.", "\u201cNo change\u201d selected \u2013 the background stays."), "warn")
            return
        ok = HM.set_scene_background(hx)
        HM.redraw()
        PRESENT.invalidate_frame()
        BUS.status(T("T\u0142o sceny: %s. Na zrzutach t\u0142o jest nak\u0142adane zawsze (HM zapisuje zrzut na bia\u0142ym tle).",
                     "Scene background: %s. On screenshots the background is always applied (HM saves captures on white).", bg_label(VIEWS.bg))
                   if ok else T("Nie uda\u0142o si\u0119 zmieni\u0107 t\u0142a sceny \u2013 zrzuty i tak dostan\u0105 wybrane t\u0142o.", "Could not change the scene background \u2013 captures still get it."))

    def generate(self):
        self.store_to_engine()
        save_settings()
        self.studio.runner.run(lambda: BUS.status(VIEWS.generate()), lambda _: self.redraw())

    def set_busy(self, on):
        self.b_gen.setEnabled(not on)
        self.b_rem.setEnabled(not on)


# ================== GUI: KARTA "PREZENTACJA PPTX" ======================
# Kroki: 1 plik i tryb (dopisanie / nowa), 2 zawartosc i uklad z INTERAKTYWNYM
# podgladem (wlasny kadr myszka, szybkie dopasowania, zapisane kadry), tytul
# (szablon), tabela statystyk, slajd zbiorczy, kamera, format obrazu,
# 3 eksport: seria, biezacy widok, podglad serii, zbiorczy, widoki, delta, raport.
class PptTab(_QWidget):
    def __init__(self, studio):
        super(PptTab, self).__init__()
        self.studio = studio
        v = QtWidgets.QVBoxLayout(self)
        f = QtWidgets.QGridLayout()
        self.e_file = QtWidgets.QLineEdit()
        b_f = QtWidgets.QPushButton(T("Wybierz\u2026", "Browse\u2026"))
        b_f.clicked.connect(self.browse)
        self.rb_app = QtWidgets.QRadioButton(T("dopisz slajdy, je\u015bli plik istnieje (np. firmowy szablon; kopia .bak)", "append slides if the file exists (e.g. company template; .bak copy)"))
        self.rb_new = QtWidgets.QRadioButton(T("zawsze nowa prezentacja 16:9", "always a new 16:9 presentation"))
        radio_group(self, self.rb_app, self.rb_new)
        f.addWidget(self.e_file, 0, 0)
        f.addWidget(b_f, 0, 1)
        f.addLayout(hrow(self.rb_app, self.rb_new, None), 1, 0, 1, 2)
        v.addWidget(group(T("1  Plik prezentacji", "1  Presentation file"), f))
        # --- zawartosc + interaktywny podglad ---
        c = QtWidgets.QGridLayout()
        self.cb_prev = QtWidgets.QCheckBox(T("PODGL\u0104D WSZYSTKICH SLAJD\u00d3W przed zapisem (Utw\u00f3rz / Anuluj)", "PREVIEW ALL SLIDES before saving (Create / Cancel)"))
        self.cb_prev.setStyleSheet("font-weight:bold")
        c.addWidget(self.cb_prev, 0, 0, 1, 4)
        self.bar = LayoutBar(self, self.on_layout_change)
        c.addWidget(self.bar, 1, 0, 1, 4)
        sw, sh = PRESENT.slide_size()
        self.canvas = SlideCanvas(self, sw, sh, min_w=520, interactive=True)
        self.canvas.setMaximumHeight(360)
        self.canvas.rebuild_cb = self._ops
        self.canvas.changed_cb = self.on_box_dragged
        tip(self.canvas, T("Interaktywny podgl\u0105d slajdu: w uk\u0142adzie \u201ew\u0142asny kadr\u201d przeci\u0105gaj obraz i zmieniaj rozmiar uchwytami w rogach. "
                           "Przeci\u0105gni\u0119cie obrazu w innym uk\u0142adzie prze\u0142\u0105cza na w\u0142asny kadr. Dwuklik = kadr domy\u015blny.",
                           "Interactive slide preview: in the \u201ccustom frame\u201d layout drag the image and resize it with the corner handles. "
                           "Dragging the image in another layout switches to the custom frame. Double-click = default frame."))
        c.addWidget(self.canvas, 2, 0, 1, 4)
        self.cb_title = QtWidgets.QCheckBox(T("Tytu\u0142 (szablon):", "Title (template):"))
        self.e_tpl = QtWidgets.QLineEdit()
        c.addWidget(self.cb_title, 3, 0)
        c.addWidget(self.e_tpl, 3, 1, 1, 3)
        c.addWidget(note_label(T("(puste = \u201e%s\u201d; pola: {metric} {thr} {view} {model} {date} {dim})", "(empty = \u201c%s\u201d; fields: {metric} {thr} {view} {model} {date} {dim})",
                                 Presenter.default_tpl())), 4, 1, 1, 3)
        self.cb_sum = QtWidgets.QCheckBox(T("slajd zbiorczy (tabela wszystkich metryk)", "summary slide (table of all metrics)"))
        self.cb_comb = QtWidgets.QCheckBox(T("slajd widoku zbiorczego", "combined-view slide"))
        self.cb_keep = QtWidgets.QCheckBox(T("zapisz te\u017c obrazy obok PPTX", "also save images next to the PPTX"))
        self.cb_leg = QtWidgets.QCheckBox(T("legenda aktywnej metryki tak\u017ce na slajdach widok\u00f3w", "active metric legend also on view slides"))
        c.addWidget(self.cb_sum, 5, 0, 1, 2)
        c.addWidget(self.cb_comb, 5, 2, 1, 2)
        c.addWidget(self.cb_keep, 6, 0, 1, 2)
        c.addWidget(self.cb_leg, 6, 2, 1, 2)
        self.rb_png = QtWidgets.QRadioButton("PNG")
        self.rb_jpg = QtWidgets.QRadioButton("JPG")
        radio_group(self, self.rb_png, self.rb_jpg)
        self.rb_cur = QtWidgets.QRadioButton(T("bie\u017c\u0105ca kamera", "current camera"))
        self.rb_shots = QtWidgets.QRadioButton(T("zaznaczone widoki z karty \u201eWidoki\u201d (ka\u017cda metryka \u00d7 ka\u017cdy widok)", "ticked views from the \u201cViews\u201d tab (each metric \u00d7 each view)"))
        radio_group(self, self.rb_cur, self.rb_shots)
        fmt_note = note_label(T("(t\u0142o zrzutu \u2013 jak na karcie \u201eWidoki\u201d)", "(capture background \u2013 as on the \u201cViews\u201d tab)"))
        fmt_note.setWordWrap(False)
        c.addLayout(hrow(QtWidgets.QLabel(T("Kamera serii:", "Series camera:")), self.rb_cur, self.rb_shots, None), 7, 0, 1, 4)
        c.addLayout(hrow(QtWidgets.QLabel(T("Format obrazu:", "Image format:")), self.rb_png, self.rb_jpg, 12, fmt_note, None), 8, 0, 1, 4)
        c.setColumnStretch(3, 1)
        v.addWidget(group(T("2  Zawarto\u015b\u0107 i uk\u0142ad slajd\u00f3w (podgl\u0105d interaktywny)", "2  Slide content and layout (interactive preview)"), c))
        # --- eksport ---
        self.b_series = styled_button(T("\u25b6 Seria: wszystkie metryki \u2192 PPTX (F7)", "\u25b6 Series: all metrics \u2192 PPTX (F7)"), GO_COLOR, big=True)
        self.b_series.clicked.connect(self.series)
        tip(self.b_series, T("Ka\u017cda przeanalizowana metryka (+ widok zbiorczy) \u00d7 kamera \u2192 slajd z obrazem, legend\u0105 i tabel\u0105 statystyk; najpierw podgl\u0105d wszystkich slajd\u00f3w.",
                             "Each analysed metric (+ combined view) \u00d7 camera \u2192 slide with image, legend and statistics table; all slides are previewed first."))
        b_cur = QtWidgets.QPushButton(T("Bie\u017c\u0105cy widok \u2192 slajd", "Current view \u2192 slide"))
        b_cur.clicked.connect(self.current_to_pptx)
        b_prev = QtWidgets.QPushButton(T("Podgl\u0105d serii\u2026", "Preview the series\u2026"))
        b_prev.clicked.connect(self.preview_series)
        tip(b_prev, T("Buduje wszystkie slajdy serii (ze zrzutami) i pokazuje je bez zapisu.", "Builds all series slides (with captures) and shows them without saving."))
        b_live = QtWidgets.QPushButton(T("Podgl\u0105d na \u017cywo (F6)", "Live preview (F6)"))
        b_live.clicked.connect(self.studio.live_toggle)
        b_sum = QtWidgets.QPushButton(T("Slajd zbiorczy", "Summary slide"))
        b_sum.clicked.connect(lambda: self.deliver(PRESENT.summary_items))
        v.addLayout(hrow(self.b_series, b_cur, b_prev, b_live, b_sum, None))
        b_v = QtWidgets.QPushButton(amp(T("Widoki z karty \u201eWidoki\u201d \u2192 PPTX", "Views from \u201cViews\u201d \u2192 PPTX")))
        b_v.clicked.connect(lambda: self.deliver(PRESENT.shots_items))
        b_d = QtWidgets.QPushButton(T("Wynik delty REF/INF \u2192 slajd", "REF/INF delta result \u2192 slide"))
        b_d.clicked.connect(self.delta_to_pptx)
        b_r = QtWidgets.QPushButton(T("Raport jako\u015bci \u2192 slajd z tabel\u0105", "Quality report \u2192 table slide"))
        b_r.clicked.connect(self.report_to_pptx)
        b_last = QtWidgets.QPushButton(T("Ostatni eksport\u2026", "Last export\u2026"))
        b_last.clicked.connect(self.show_last)
        v.addWidget(group(T("3  Eksport z innych kart", "3  Export from other tabs"), hrow(b_v, b_d, b_r, None, b_last)))
        v.addStretch(1)
        self.load_from_engine()

    def load_from_engine(self):
        P = PRESENT
        self.e_file.setText(P.ppt_file)
        (self.rb_new if P.mode == "new" else self.rb_app).setChecked(True)
        self.cb_prev.setChecked(P.preview_on)
        self.bar.load()
        self.cb_title.setChecked(P.style.title_on)
        self.e_tpl.setText(P.title_tpl)
        self.cb_sum.setChecked(P.summary_on)
        self.cb_comb.setChecked(P.combined_on)
        self.cb_keep.setChecked(P.keep_images)
        (self.rb_jpg if P.img_fmt == "jpg" else self.rb_png).setChecked(True)
        (self.rb_shots if P.cam_src == "shots" else self.rb_cur).setChecked(True)
        self.cb_leg.setChecked(P.leg_in_shots)

    def store_to_engine(self):
        P = PRESENT
        P.ppt_file = self.e_file.text().strip()
        P.mode = "new" if self.rb_new.isChecked() else "append"
        P.preview_on = self.cb_prev.isChecked()
        P.style.title_on = self.cb_title.isChecked()
        P.title_tpl = self.e_tpl.text()
        P.summary_on, P.combined_on = self.cb_sum.isChecked(), self.cb_comb.isChecked()
        P.keep_images = self.cb_keep.isChecked()
        P.img_fmt = "jpg" if self.rb_jpg.isChecked() else "png"
        P.cam_src = "shots" if self.rb_shots.isChecked() else "current"
        P.leg_in_shots = self.cb_leg.isChecked()

    # ---------------------------------------------------------- podglad
    def _ops(self):
        try:
            sw, sh = PRESENT.slide_size()
            self.canvas.set_slide_size(sw, sh)
            return slide_ops(PRESENT.sample_item(capture=False), sw, sh, PRESENT.style)
        except Exception as e:
            BUS.log("ppt preview: %s" % e)
            return []

    def refresh_preview(self):
        """Nowy obraz plotna z biezacego stanu (bez zrzutu z HM; zrzut robi
        karta "Podglad slajdu" / okno na zywo, gdy stan sie zmienil)."""
        self.store_to_engine()
        self.canvas.clear_cache()
        self.canvas.set_ops(self._ops())

    def on_layout_change(self):
        self.store_to_engine()
        self.refresh_preview()
        save_settings()
        self.studio.preview_changed()

    def on_box_dragged(self):
        self.bar.refresh_from_style()
        save_settings()
        self.studio.preview_changed()

    # ---------------------------------------------------------- akcje
    def browse(self):
        start = self.e_file.text() or os.path.join(HM.model_dir() or "", "%s_jakosc.pptx" % (HM.model_name() if HM.ok() else "prezentacja"))
        p = ask_save_file(self, T("Plik prezentacji (.pptx)", "Presentation file (.pptx)"), "PowerPoint (*.pptx)", start, confirm=False)
        if p:
            self.e_file.setText(os.path.normpath(p))
            PRESENT.ppt_file = self.e_file.text()
            save_settings()
            self.studio.preview_changed()

    def ask_path(self):
        self.browse()
        return self.e_file.text().strip()

    def deliver(self, make_items):
        """Buduje slajdy (zrzuty) i przekazuje je do podgladu / zapisu."""
        self.store_to_engine()
        save_settings()
        items = self.studio.runner.run(make_items)
        if items:
            self.studio.runner.run(lambda: PRESENT.deliver(items), busy=False)
            self.load_from_engine()
            PRESENT.invalidate_frame()

    def series(self):
        self.studio.tab_mq.store_to_engine()
        if not MQ.analyzed:
            BUS.status(T("Seria wymaga analizy metryk \u2013 karta \u201e1 Metryki\u201d, przycisk \u201eAnalizuj i koloruj\u201d (F5).",
                         "The series needs a metrics analysis \u2013 \u201c1 Metrics\u201d tab, \u201cAnalyze and color\u201d (F5)."), "warn")
            self.studio.goto("mq")
            return
        self.deliver(PRESENT.series_items)

    def current_to_pptx(self):
        self.deliver(lambda: [PRESENT.current_item()])

    def preview_series(self):
        self.store_to_engine()
        self.studio.tab_mq.store_to_engine()
        items = self.studio.runner.run(PRESENT.series_items if MQ.analyzed else (lambda: [PRESENT.current_item()]))
        if items:
            SlidePreview.run(self.studio, items, ask=False)
            Presenter.drop_items(items)
            self.load_from_engine()
            self.refresh_preview()

    def delta_to_pptx(self):
        self.deliver(PRESENT.delta_items)

    def report_to_pptx(self):
        self.deliver(PRESENT.report_items)

    def show_last(self):
        le = PRESENT.last_export
        if not le:
            BUS.status(T("W tej sesji nie by\u0142o jeszcze eksportu.", "No export yet in this session."), "warn")
            return
        ExportDoneDialog(self.studio, le).exec_()

    def set_busy(self, on):
        self.b_series.setEnabled(not on)


# =================== GUI: KARTA "PODGLAD SLAJDU" ========================
# Tylko do odczytu (jak karta 0 w ANSYS SHOTS): przykladowy slajd z biezacego
# stanu, "co powstanie", przyciski: nowa klatka z HM, okno na zywo, podglad
# serii, utworz prezentacje. Odswiezana przy wejsciu i po zmianie opcji.
class PreviewTab(_QWidget):
    def __init__(self, studio):
        super(PreviewTab, self).__init__()
        self.studio = studio
        v = QtWidgets.QVBoxLayout(self)
        v.addWidget(note_label(T("Ta karta tylko pokazuje, co powstanie z bie\u017c\u0105cych opcji. Opcje zmieniasz na kartach \u201eMetryki\u201d, \u201eWidoki\u201d i \u201ePrezentacja\u201d "
                                 "(tam te\u017c interaktywny kadr). Ten sam podgl\u0105d w osobnym oknie, widoczny przy pracy na ka\u017cdej karcie: F6 (\u201ePodgl\u0105d na \u017cywo\u201d).",
                                 "This tab only shows what the current options produce. You change them on the \u201cMetrics\u201d, \u201cViews\u201d and \u201cPowerPoint\u201d tabs "
                                 "(the interactive frame is there too). The same preview in a separate window, visible while you work on any tab: F6 (\u201cLive preview\u201d)."), "#556"))
        self.panel = PreviewPanel(self, studio, min_w=520)
        v.addWidget(self.panel, 1)
        b_ref = QtWidgets.QPushButton(T("Od\u015bwie\u017c (nowa klatka z okna HM)", "Refresh (new frame from the HM window)"))
        b_ref.clicked.connect(self.recapture)
        b_live = QtWidgets.QPushButton(T("Podgl\u0105d na \u017cywo (F6)", "Live preview (F6)"))
        b_live.clicked.connect(self.studio.live_toggle)
        b_ser = QtWidgets.QPushButton(T("Podgl\u0105d wszystkich slajd\u00f3w serii\u2026", "Preview all series slides\u2026"))
        b_ser.clicked.connect(self.studio.tab_ppt.preview_series)
        b_go = styled_button(T("\u25b6 Utw\u00f3rz prezentacj\u0119 (F7)", "\u25b6 Create the presentation (F7)"), GO_COLOR)
        b_go.clicked.connect(self.studio.tab_ppt.series)
        v.addLayout(hrow(b_ref, b_live, b_ser, None, b_go))

    def refresh(self, capture=True):
        self.panel.refresh(capture)

    def recapture(self):
        PRESENT.invalidate_frame()
        self.studio.runner.run(lambda: self.refresh(True), busy=False)

    def set_busy(self, on):
        pass


# =========================== GUI: OKNO GLOWNE ==========================
# StudioWindow: naglowek (tytul, skroty, PL / EN, Diagnostyka, Pomoc), karty
# Start / 1 Metryki / 2 Delta / 3 Raport / 4 Widoki / 5 Prezentacja / Podglad
# slajdu, pasek stanu z paskiem postepu i globalnym "Przerwij", dziennik.
# Okno jest dzieckiem glownego okna HyperMesha (trzyma sie nad nim), ale
# NIEMODALNE. Kazda zmiana opcji (checkbox, radio, pole, lista) odswieza po
# LIVE_DELAY_MS podglad slajdu (karta "Podglad slajdu", okno na zywo, plotno
# karty PPTX) - jedno miejsce (_hook_changes) zamiast wywolan w kazdej obsludze.
class StudioWindow(_QWidget):
    TAB_KEYS = ("start", "mq", "delta", "report", "views", "ppt", "preview")

    def __init__(self, parent=None):
        super(StudioWindow, self).__init__(parent, QtCore.Qt.Window)
        self.setWindowTitle("%s v%s" % (APP_TITLE, VERSION))
        self.runner = ActionRunner(self)
        self.legends = {}
        self.editor = None
        self.live = None
        self.diag = None
        self._closing = False
        self._timer = QtCore.QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self.refresh_previews)
        v = QtWidgets.QVBoxLayout(self)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)
        # --- naglowek ---
        hdr = QtWidgets.QFrame()
        hdr.setStyleSheet("QFrame{background:%s}QLabel{color:white}" % HDR_COLOR)
        hl = QtWidgets.QHBoxLayout(hdr)
        hl.setContentsMargins(12, 8, 10, 8)
        tv = QtWidgets.QVBoxLayout()
        t = QtWidgets.QLabel(APP_TITLE)
        tf = t.font()
        tf.setPointSizeF(tf.pointSizeF() + 6)
        tf.setBold(True)
        t.setFont(tf)
        s = QtWidgets.QLabel(T("Metryki \u2192 grupy kolor\u00f3w \u2022 delta REF/INF \u2022 raport \u2022 widoki i zrzuty \u2022 podgl\u0105d slajdu i eksport PPTX   "
                               "|   F5 analiza \u2022 F6 podgl\u0105d na \u017cywo \u2022 F7 seria PPTX \u2022 1 widok \u2022 K zrzuty \u2022 F1 pomoc",
                               "Metrics \u2192 color groups \u2022 REF/INF delta \u2022 report \u2022 views & screenshots \u2022 slide preview and PPTX export   "
                               "|   F5 analysis \u2022 F6 live preview \u2022 F7 PPTX series \u2022 1 view \u2022 K shots \u2022 F1 help"))
        s.setStyleSheet("color:#cfe0f5")
        s.setWordWrap(True)
        tv.addWidget(t)
        tv.addWidget(s)
        hl.addLayout(tv, 1)
        for code in ("pl", "en"):
            b = QtWidgets.QPushButton(code.upper())
            b.setCheckable(True)
            b.setChecked(lang() == code)
            b.setFixedWidth(38)
            b.setStyleSheet("QPushButton{background:#16487a;color:white;border:1px solid #cfe0f5;font-weight:bold}"
                            "QPushButton:checked{background:white;color:%s}" % HDR_COLOR)
            b.clicked.connect(lambda _=False, c=code: self.switch_lang(c))
            hl.addWidget(b, 0, QtCore.Qt.AlignTop)
        bd = QtWidgets.QPushButton(T("Diagnostyka", "Diagnostics"))
        bd.setStyleSheet("QPushButton{background:#16487a;color:white;border:1px solid #cfe0f5;padding:2px 8px}")
        bd.clicked.connect(self.show_diagnostics)
        hl.addWidget(bd, 0, QtCore.Qt.AlignTop)
        bh = QtWidgets.QPushButton(T("? Pomoc (F1)", "? Help (F1)"))
        bh.setStyleSheet("QPushButton{background:#16487a;color:white;border:1px solid #cfe0f5;padding:2px 8px}")
        bh.clicked.connect(self.show_help)
        hl.addWidget(bh, 0, QtCore.Qt.AlignTop)
        v.addWidget(hdr)
        # --- karty ---
        self.tabs = QtWidgets.QTabWidget()
        self.tab_start = StartTab(self)
        self.tab_mq = MqTab(self)
        self.tab_delta = DeltaTab(self)
        self.tab_report = ReportTab(self)
        self.tab_views = ViewsTab(self)
        self.tab_ppt = PptTab(self)
        self.tab_preview = PreviewTab(self)
        self.tab_list = [self.tab_start, self.tab_mq, self.tab_delta, self.tab_report, self.tab_views, self.tab_ppt, self.tab_preview]
        for w, pl, en in ((self.tab_start, "Start", "Start"),
                          (self.tab_mq, "1  Metryki i grupy", "1  Metrics & groups"),
                          (self.tab_delta, "2  Delta REF / INF", "2  Delta REF / INF"),
                          (self.tab_report, "3  Raport jako\u015bci", "3  Quality report"),
                          (self.tab_views, "4  Widoki i zrzuty", "4  Views & screenshots"),
                          (self.tab_ppt, "5  Prezentacja PPTX", "5  PowerPoint"),
                          (self.tab_preview, "\u25a3 Podgl\u0105d slajdu", "\u25a3 Slide preview")):
            sa = QtWidgets.QScrollArea()
            sa.setWidget(w)
            sa.setWidgetResizable(True)
            sa.setFrameShape(QtWidgets.QFrame.NoFrame)
            self.tabs.addTab(sa, amp(T(pl, en)))
        self.tabs.currentChanged.connect(self.on_tab)
        cont = QtWidgets.QWidget()
        cl = QtWidgets.QVBoxLayout(cont)
        cl.setContentsMargins(8, 6, 8, 4)
        cl.addWidget(self.tabs, 1)
        # --- pasek stanu: status, postep, przerwij, dziennik ---
        self.status = QtWidgets.QLabel()
        self.status.setWordWrap(True)
        self.status.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
        self.status.setMinimumHeight(34)
        self.status.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        self.pbar = QtWidgets.QProgressBar()
        self.pbar.setRange(0, 0)
        self.pbar.setFixedWidth(160)
        self.pbar.setTextVisible(True)
        self.pbar.setVisible(False)
        self.b_stop = QtWidgets.QPushButton(T("\u25a0 Przerwij", "\u25a0 Stop"))
        self.b_stop.setEnabled(False)
        self.b_stop.clicked.connect(self.stop)
        tip(self.b_stop, T("Przerywa d\u0142ug\u0105 operacj\u0119 (analiza, raport, seria) w najbli\u017cszym bezpiecznym miejscu.",
                           "Stops a long operation (analysis, report, series) at the next safe point."))
        self.b_log = QtWidgets.QPushButton(T("Dziennik", "Log"))
        self.b_log.setCheckable(True)
        self.b_log.toggled.connect(self._toggle_log)
        for b in (self.b_stop, self.b_log):
            b.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        row = hrow(self.status, self.pbar, self.b_stop, self.b_log)
        row.setStretch(0, 1)
        for w in (self.pbar, self.b_stop, self.b_log):
            row.setAlignment(w, QtCore.Qt.AlignTop)
        cl.addLayout(row)
        self.log = QtWidgets.QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(140)
        self.log.setStyleSheet("font-family:Consolas,monospace;font-size:8pt")
        self.log.setVisible(False)
        cl.addWidget(self.log)
        v.addWidget(cont, 1)
        # --- polaczenia z silnikami ---
        BUS.status_cb = self.set_status
        BUS.log_cb = self.add_log
        BUS.confirm_cb = lambda title, msg: yes_no(self, title, msg)
        BUS.info_cb = lambda title, msg, level: msg_box(self, title, msg, level)
        PRESENT.preview_cb = lambda items, ask: SlidePreview.run(self, items, ask)
        PRESENT.ask_path_cb = self.tab_ppt.ask_path
        PRESENT.open_cb = self.export_done
        PRESENT.render_cb = render_items_png
        for key, fn in (("1", self.tab_views.remember), ("K", self.tab_views.generate)):
            sc = QtWidgets.QShortcut(QtGui.QKeySequence(key), self)
            sc.setContext(QtCore.Qt.WindowShortcut)
            sc.activated.connect(lambda fn=fn: self.hotkey(fn))
        for key, fn in ((QtCore.Qt.Key_F1, self.show_help), (QtCore.Qt.Key_F5, self.tab_mq.analyze),
                        (QtCore.Qt.Key_F6, self.live_toggle), (QtCore.Qt.Key_F7, self.tab_ppt.series)):
            sc = QtWidgets.QShortcut(QtGui.QKeySequence(key), self)
            sc.setContext(QtCore.Qt.WindowShortcut)
            sc.activated.connect(fn)
        # szerokosc wg zawartosci kart (bez poziomego przewijania przy skali
        # ekranu 125-150 %), ale nie wiecej niz 95 % ekranu
        scr = screen_rect(self)
        want = max([1000] + [t.sizeHint().width() + 60 for t in self.tab_list])
        self.resize(min(want, int(scr.width() * 0.95)), min(920, int(scr.height() * 0.88)))
        geo = GUI_PREFS.get("win_geo") or []
        if len(geo) == 4:
            r = QtCore.QRect(*[int(x) for x in geo])
            if scr.intersects(r) and r.width() > 400 and r.height() > 300:
                self.setGeometry(r)
        if GUI_PREFS.get("log_open"):
            self.b_log.setChecked(True)
        self._hook_changes()
        self.refresh_all()
        self.tabs.setCurrentIndex(max(0, min(len(self.TAB_KEYS) - 1, int(GUI_PREFS.get("tab", 0)))))
        self.set_status(T("Gotowy. Zaznacz metryki na karcie \u201e1 Metryki\u201d i kliknij \u201eAnalizuj i koloruj\u201d (F5), potem \u201eSeria \u2192 PPTX\u201d (F7). Karta Start pokazuje stan pracy.",
                          "Ready. Tick metrics on the \u201c1 Metrics\u201d tab and click \u201cAnalyze and color\u201d (F5), then \u201cSeries \u2192 PPTX\u201d (F7). The Start tab shows the work status."))

    # ---------------------------------------------------------- stan / dziennik
    def set_status(self, msg, level="ok"):
        self.status.setText(msg)
        self.status.setStyleSheet("color:%s" % Bus.LEVELS.get(level, OK_COLOR))
        if level in ("warn", "err"):
            self.add_log(msg)
        if self.pbar.isVisible():
            m = re.search(r"(\d+)\s*/\s*(\d+)", msg)
            if m and int(m.group(2)) > 0:
                self.pbar.setRange(0, int(m.group(2)))
                self.pbar.setValue(min(int(m.group(1)), int(m.group(2))))
            else:
                self.pbar.setRange(0, 0)

    def add_log(self, msg):
        self.log.appendPlainText("%s  %s" % (now_text("%H:%M:%S"), msg))

    def _toggle_log(self, on):
        self.log.setVisible(on)
        GUI_PREFS["log_open"] = bool(on)

    def set_busy(self, on):
        for tab in self.tab_list:
            try:
                tab.set_busy(on)
            except Exception:
                pass
        self.b_stop.setEnabled(on)
        self.pbar.setVisible(on)
        self.pbar.setRange(0, 0)
        if on:
            app_instance().setOverrideCursor(QtCore.Qt.WaitCursor)
        else:
            app_instance().restoreOverrideCursor()

    def stop(self):
        BUS.cancel = True
        self.set_status(T("Przerywanie\u2026 (operacja zatrzyma si\u0119 w najbli\u017cszym bezpiecznym miejscu)", "Stopping\u2026 (the operation stops at the next safe point)"), "warn")

    def refresh_all(self):
        for tab in (self.tab_mq, self.tab_delta, self.tab_report, self.tab_start):
            try:
                tab.refresh()
            except Exception as e:
                BUS.log("refresh: %s" % e)
        for mode, w in list(self.legends.items()):
            if w is not None and w.isVisible():
                w.refresh()
        self.preview_changed()

    def goto(self, key):
        if key in self.TAB_KEYS:
            self.tabs.setCurrentIndex(self.TAB_KEYS.index(key))

    def on_tab(self, idx):
        GUI_PREFS["tab"] = int(idx)
        if idx == self.TAB_KEYS.index("preview"):
            self.tab_preview.refresh(not self.runner.busy)
        elif idx == self.TAB_KEYS.index("ppt"):
            self.tab_ppt.refresh_preview()
        elif idx == 0:
            self.tab_start.refresh()

    def hotkey(self, fn):
        f = app_instance().focusWidget()
        if isinstance(f, (QtWidgets.QLineEdit, QtWidgets.QPlainTextEdit, QtWidgets.QTextEdit, QtWidgets.QAbstractSpinBox)):
            return
        fn()

    # ---------------------------------------------------------- podglady
    def _hook_changes(self):
        """Kazda opcja okna odswieza podglad slajdu po zmianie (z opoznieniem)."""
        for tab in self.tab_list[:-1]:
            for w in tab.findChildren(QtWidgets.QWidget):
                try:
                    if isinstance(w, (QtWidgets.QCheckBox, QtWidgets.QRadioButton)):
                        w.toggled.connect(lambda *a: self.preview_changed())
                    elif isinstance(w, QtWidgets.QLineEdit) and not w.isReadOnly():
                        w.textChanged.connect(lambda *a: self.preview_changed())
                    elif isinstance(w, QtWidgets.QAbstractSpinBox):
                        w.valueChanged.connect(lambda *a: self.preview_changed())
                    elif isinstance(w, QtWidgets.QComboBox):
                        w.currentIndexChanged.connect(lambda *a: self.preview_changed())
                except Exception:
                    pass

    def preview_changed(self):
        if not self._closing:
            self._timer.start(LIVE_DELAY_MS)

    def refresh_previews(self):
        """Jeden obraz biezacych opcji dla plotna karty PPTX, karty "Podglad
        slajdu" i okna na zywo. W trakcie innej operacji - bez zrzutu z HM."""
        if self._closing:
            return
        capture = not self.runner.busy
        self.store_all()
        try:
            self.tab_ppt.refresh_preview()
        except Exception as e:
            BUS.log("preview ppt: %s" % e)
        if self.tabs.currentIndex() == self.TAB_KEYS.index("preview"):
            self.tab_preview.refresh(capture)
        if self.live is not None:
            self.live.refresh(capture)

    def live_toggle(self):
        if self.live is not None:
            self.live.close()
            self.live = None
            return
        self.live = LivePreviewWindow(self)
        self.live.show()
        self.live.refresh(not self.runner.busy)
        GUI_PREFS["live_open"] = True
        save_settings()

    def export_done(self, msg, path):
        le = PRESENT.last_export or {"path": path, "n": 0, "msg": msg, "titles": [], "thumbs": []}
        ExportDoneDialog(self, le).exec_()

    def ask_open(self, msg, path):
        if yes_no(self, "PPTX", T("%s\n\nOtworzy\u0107 prezentacj\u0119?", "%s\n\nOpen the presentation?", msg)):
            open_path(path)

    # ---------------------------------------------------------- okna pomocnicze
    def show_legend(self, mode):
        if mode == "mq" and not MQ.view:
            BUS.status(T("Najpierw wykonaj analiz\u0119.", "Run the analysis first."), "err")
            return
        if mode == "delta" and not DELTA.done:
            BUS.status(T("Brak wyniku delty na siatce.", "No delta result on the mesh."), "err")
            return
        w = self.legends.get(mode)
        if w is None:
            w = LegendWindow(self, mode)
            self.legends[mode] = w
            g = self.frameGeometry()
            w.move(g.right() + 8, g.top())
        w.refresh()
        w.show()
        w.raise_()

    def open_editor(self, m=""):
        m = m if m in MQ_ORDER else (MQ.view if MQ.view in MQ_ORDER else "ar")
        if self.editor is not None:
            self.editor.close()
        self.editor = LegendEditor(self, m)
        self.editor.show()

    def show_diagnostics(self):
        if self.diag is None:
            self.diag = DiagnosticsDialog(self)
        self.diag.show()
        self.diag.raise_()
        if not self.diag.rows:
            self.diag.run()

    def restore_any(self):
        def work():
            if not MESH.active():
                BUS.status(T("Siatka nie jest pokolorowana przez narz\u0119dzie \u2013 nie ma czego przywraca\u0107.", "The mesh is not colored by the tool \u2013 nothing to restore."), "warn")
                return
            if MESH.owner == "delta":
                DELTA.restore()
            else:
                MQ.restore()
            PRESENT.invalidate_frame()
        self.runner.run(work)

    def show_help(self):
        d = QtWidgets.QDialog(self)
        d.setWindowTitle(T("Pomoc", "Help"))
        lay = QtWidgets.QVBoxLayout(d)
        tb = QtWidgets.QTextBrowser()
        tb.setHtml(help_html())
        lay.addWidget(tb)
        b = QtWidgets.QPushButton(T("Zamknij", "Close"))
        b.clicked.connect(d.close)
        lay.addLayout(hrow(None, b))
        d.resize(800, 680)
        d.show()

    def switch_lang(self, code):
        if code == lang():
            return
        self.store_all()
        set_lang(code)
        save_settings()
        geo = self.geometry()
        new = StudioWindow(self.parent())
        new.setGeometry(geo)
        builtins._HMQS_STUDIO = new
        new.show()
        self.close()

    def store_all(self):
        for tab in (self.tab_mq, self.tab_delta, self.tab_views, self.tab_ppt, self.tab_report):
            try:
                tab.store_to_engine()
            except Exception:
                pass

    def closeEvent(self, ev):
        self._closing = True
        self._timer.stop()
        self.store_all()
        g = self.geometry()
        GUI_PREFS["win_geo"] = [g.x(), g.y(), g.width(), g.height()]
        GUI_PREFS["live_open"] = self.live is not None
        save_settings()
        for w in list(self.legends.values()) + [self.editor, self.live, self.diag]:
            try:
                if w is not None:
                    w.close()
            except Exception:
                pass
        if BUS.status_cb == self.set_status:
            BUS.status_cb = BUS.log_cb = BUS.confirm_cb = BUS.info_cb = None
            PRESENT.preview_cb = PRESENT.ask_path_cb = PRESENT.open_cb = PRESENT.render_cb = None
        super(StudioWindow, self).closeEvent(ev)


# ================================ POMOC ================================
# Tresc okna "? Pomoc" (HTML). Jedna funkcja na jezyk - latwo dopisac akapit.
def help_html():
    return _HELP_EN if lang() == "en" else _HELP_PL


_HELP_PL = """
<h2>HM Quality Studio 3.0 \u2013 jedno narz\u0119dzie zamiast trzech makr</h2>
<p>Okno jest niemodalne: w trakcie pracy mo\u017cna obraca\u0107 model, zmienia\u0107 wy\u015bwietlanie i zapami\u0119tywa\u0107 widoki.
Ustawienia (progi, legendy, opcje, kadry) zapisuj\u0105 si\u0119 same w pliku <code>.hm_quality_studio.json</code> w katalogu u\u017cytkownika.
Skr\u00f3ty: <b>F5</b> analiza metryk, <b>F6</b> podgl\u0105d na \u017cywo, <b>F7</b> seria \u2192 PPTX, <b>1</b> zapami\u0119taj widok, <b>K</b> zrzuty, <b>F1</b> pomoc.
Karta <b>Start</b> pokazuje stan ka\u017cdego kroku i ma przyciski \u201ePrzejd\u017a\u201d.</p>

<h3>1. Metryki i grupy kolor\u00f3w</h3>
<ul>
<li>Zaznacz metryki: <b>Aspect Ratio, Jacobian Ratio, Jacobian Zero, Skewness</b>. S\u0105 czytane <b>jednocze\u015bnie</b>,
w jednym przebiegu po elementach (2D i/lub 3D, ca\u0142a siatka albo tylko wy\u015bwietlone).</li>
<li><b>Podzia\u0142 wg progu</b>: elementy w normie trafiaj\u0105 do jednej grupy bazowej (zielonej), elementy poza norm\u0105 \u2013
do kolejnych pasm kolor\u00f3w (np. AR: 4\u20135 niebieski \u2026 12\u201315 czerwony), warto\u015bci poza ostatni\u0105 granic\u0105 \u2013 do osobnej grupy
\u201epoza skal\u0105\u201d (magenta).</li>
<li><b>Legenda r\u0119czna</b> (\u201eLegenda\u2026\u201d): liczba kolor\u00f3w, granice i kolor ka\u017cdego pasma. <b>Legenda automatyczna</b>: przedzia\u0142y od progu
do warto\u015bci skrajnej podzielone na wskazan\u0105 liczb\u0119 pasm. Kolory pochodz\u0105 z <b>palety 64 kolor\u00f3w HyperMesha</b>, wi\u0119c legenda,
okno i slajd pokazuj\u0105 dok\u0142adnie kolory z siatki.</li>
<li>\u201ePoka\u017c na siatce\u201d prze\u0142\u0105cza metryk\u0119 <b>bez ponownego czytania</b>. \u201eZbiorczo\u201d = ile metryk element przekracza.
<b>Zestawy (sets)</b> powstaj\u0105 dla wszystkich metryk naraz. <b>Przywr\u00f3\u0107 siatk\u0119</b> odk\u0142ada elementy do pierwotnych komponent\u00f3w \u2013 nic nie ginie.</li>
<li>Nazwy danych HM (aspect / aspectratio, skew / skewness \u2026) s\u0105 dobierane automatycznie; odczyt hurtowy paruje warto\u015bci z ID element\u00f3w.
<b>Diagnostyka</b> (nag\u0142\u00f3wek okna) pokazuje nazwy danych i warto\u015bci na pr\u00f3bce element\u00f3w oraz sprawdza zgodno\u015b\u0107 odczytu hurtowego z pojedynczym
i wzorce geometrii (Jacobian Zero, tet collapse).</li>
</ul>

<h3>2. Delta REF / INF</h3>
<ul>
<li><b>Tryb</b>: delta REF \u2192 INF (dwa modele), jeden model: pasma wg warto\u015bci, jeden model: podzia\u0142 wg progu. Pliki s\u0105 wczytywane kolejno
(<b>zast\u0119puj\u0105 model w sesji</b> \u2013 okno pyta wcze\u015bniej). Elementy dopasowywane s\u0105 po ID.</li>
<li>Pogorszenie D: AR i Skewness: INF \u2212 REF; Jacobian: REF \u2212 INF; Przesuni\u0119cie: maks. przesuni\u0119cie w\u0119z\u0142\u00f3w elementu [mm].
D poni\u017cej progu \u201ebez zmian\u201d (w tym poprawa) = szary, brak odpowiednika = magenta.</li>
<li>Metryka i typ element\u00f3w to <b>osobne</b> wybory (wyb\u00f3r 2D/3D nie odznacza metryki). \u201eSprawd\u017a delt\u0119 (po ID)\u201d pokazuje Q(REF), Q(INF), \u0394, D
i pasmo elementu <b>z danych ostatniej analizy</b> \u2013 dzia\u0142a niezale\u017cnie od bie\u017c\u0105cych opcji okna.</li>
<li>Opcje: skala automatyczna / r\u0119czna, wyciszanie \u201ebez zmian\u201d (przezroczysto\u015b\u0107), znaczniki MIN / MAX, poprawione, nak\u0142adka REF
(plik REF do\u0142\u0105czony do sesji jako ukryty komponent; \u201ePrzywr\u00f3\u0107 siatk\u0119\u201d usuwa <b>wszystkie</b> do\u0142\u0105czone encje).</li>
</ul>

<h3>3. Raport jako\u015bci</h3>
<ul>
<li>\u0179r\u00f3d\u0142o: otwarty model, plik .hm albo por\u00f3wnanie A = REF vs B = INF. Metryki: AR, Jacobian, Skew, Warpage, Taper, k\u0105ty, Tet collapse,
min. d\u0142ugo\u015b\u0107 kraw\u0119dzi. Wyniki: TXT, CSV, XLSX (wykresy), interaktywny HTML (wska\u017anik 0\u2013100, histogramy, trend, <code>*createmark</code>).</li>
<li>Por\u00f3wnanie: osobne raporty _REF / _INF + raport por\u00f3wnawczy: zgodno\u015b\u0107 element\u00f3w (ID, typ, topologia, warto\u015bci, obj\u0119to\u015b\u0107 / pole) i
<b>zestawienie w\u0119z\u0142\u00f3w</b> (w\u0119z\u0142y razem, wsp\u00f3lne ID, tylko A / tylko B, bez zmiany po\u0142o\u017cenia, PRZESUNI\u0118TE ponad tolerancj\u0119, maks. i \u015brednie
przesuni\u0119cie, najbardziej przesuni\u0119te w\u0119z\u0142y; arkusz \u201eW\u0119z\u0142y\u201d w XLSX) z werdyktem.</li>
</ul>

<h3>4. Widoki i zrzuty</h3>
<ul>
<li>\u201eZapami\u0119taj widok\u201d (klawisz <b>1</b>) zapisuje kamer\u0119 i <b>pe\u0142n\u0105 klatk\u0119</b> (WYSIWYG) z miniatur\u0105. \u201eGeneruj zrzuty\u201d (klawisz <b>K</b>)
eksportuje zaznaczone widoki. HyperMesh zapisuje zrzut zawsze na bia\u0142ym tle \u2013 wybrany kolor t\u0142a jest nak\u0142adany na obraz.</li>
<li>\u201eNagrywaj uk\u0142ad\u201d do\u0142\u0105cza komendy z command.tcl (sekcje, style, maskowanie) do widoku. Raport HTML jest te\u017c <b>sesj\u0105</b> (\u201eWczytaj sesj\u0119\u2026\u201d).</li>
</ul>

<h3>5. Prezentacja PPTX i podgl\u0105d slajdu (jak w ANSYS SHOTS)</h3>
<ul>
<li><b>Interaktywny podgl\u0105d</b> na karcie PPTX: uk\u0142ad \u201ew\u0142asny kadr\u201d \u2013 przeci\u0105gasz obraz myszk\u0105, skalujesz uchwytami w rogach (Shift = proporcje),
szybkie dopasowania (1/2 lewa, pe\u0142ny\u2026), zapisane kadry pod w\u0142asn\u0105 nazw\u0105. Przeci\u0105gni\u0119cie obrazu w innym uk\u0142adzie prze\u0142\u0105cza na w\u0142asny kadr.</li>
<li>Karta <b>\u201ePodgl\u0105d slajdu\u201d</b> (tylko do odczytu): przyk\u0142adowy slajd z bie\u017c\u0105cego stanu + \u201eco powstanie\u201d (plik, tryb, seria, delta, raport, widoki).
Okno <b>\u201ePodgl\u0105d na \u017cywo\u201d (F6)</b>: to samo w osobnym oknie, odswie\u017cane samo po ka\u017cdej zmianie opcji.</li>
<li><b>Seria (F7)</b>: ka\u017cda metryka (+ widok zbiorczy) \u00d7 kamera \u2192 slajd z obrazem, legend\u0105 i tabel\u0105 statystyk (natywne kszta\u0142ty PowerPointa).
Przed zapisem <b>podgl\u0105d wszystkich slajd\u00f3w</b> (tytu\u0142y, uk\u0142ad, usuwanie, \u201eUtw\u00f3rz\u201d / \u201eAnuluj\u201d), po zapisie <b>galeria wyeksportowanych slajd\u00f3w</b>
z przyciskami \u201eOtw\u00f3rz prezentacj\u0119\u201d / \u201eOtw\u00f3rz folder\u201d.</li>
<li>Plik: nowa prezentacja 16:9 albo dopisanie slajd\u00f3w do istniej\u0105cej (np. firmowy szablon; kopia <code>.bak.pptx</code>). Z innych kart: widoki,
wynik delty, raport jako\u015bci (tabela + wska\u017anik).</li>
</ul>
"""

_HELP_EN = """
<h2>HM Quality Studio 3.0 \u2013 one tool instead of three macros</h2>
<p>The window is modeless: you can rotate the model, change the display and remember views while it is open.
Settings (thresholds, legends, options, frames) are saved automatically in <code>.hm_quality_studio.json</code> in the user folder.
Shortcuts: <b>F5</b> metrics analysis, <b>F6</b> live preview, <b>F7</b> series \u2192 PPTX, <b>1</b> remember view, <b>K</b> screenshots, <b>F1</b> help.
The <b>Start</b> tab shows the status of every step with \u201cGo\u201d buttons.</p>

<h3>1. Metrics &amp; color groups</h3>
<ul>
<li>Tick metrics: <b>Aspect Ratio, Jacobian Ratio, Jacobian Zero, Skewness</b>. They are read <b>simultaneously</b> in one pass
(2D and/or 3D, whole mesh or displayed only). <b>Threshold split</b>: within limits = one green group, out of limits = color bands,
beyond the last edge = a separate \u201cbeyond scale\u201d group.</li>
<li>Manual / automatic legend; colors from the <b>64-color HyperMesh palette</b>, so the legend, the window and the slide show exactly the mesh colors.
\u201cShow on mesh\u201d switches the metric without re-reading; sets for all metrics; \u201cRestore mesh\u201d puts elements back \u2013 nothing is lost.</li>
<li>HM data names (aspect / aspectratio, skew / skewness \u2026) are chosen automatically; bulk reads pair values with element IDs.
<b>Diagnostics</b> (window header) shows the data names with sample values and checks bulk-vs-single read consistency and geometry references.</li>
</ul>

<h3>2. Delta REF / INF</h3>
<ul>
<li><b>Mode</b>: delta REF \u2192 INF (two models), one model: bands by value, one model: threshold split. Files are loaded in sequence
(<b>they replace the model in the session</b>), elements matched by ID. Worsening D: AR / Skewness INF \u2212 REF; Jacobian REF \u2212 INF;
displacement: max node displacement [mm]. Below the \u201cno change\u201d threshold = gray, unmatched = magenta.</li>
<li>Metric and element type are <b>separate</b> choices. \u201cCheck delta (by ID)\u201d shows Q(REF), Q(INF), \u0394, D and the band of the element
<b>from the last analysis</b> \u2013 independent of the current window options.</li>
</ul>

<h3>3. Quality report</h3>
<ul>
<li>TXT, CSV, XLSX (charts) and interactive HTML (0\u2013100 score, histograms, trend, <code>*createmark</code>). REF vs INF comparison with an element
identity check (IDs, types, topology, values, volume / area) and a <b>node comparison</b> (nodes in total, common IDs, only A / only B,
unchanged, MOVED beyond the tolerance, max / mean displacement, most displaced nodes; \u201cNodes\u201d sheet in the XLSX) with a verdict.</li>
</ul>

<h3>4. Views &amp; screenshots</h3>
<ul>
<li>\u201cRemember view\u201d (key <b>1</b>) stores the camera and the <b>full frame</b> (WYSIWYG); \u201cGenerate screenshots\u201d (key <b>K</b>) exports the ticked views.
\u201cRecord layout\u201d attaches command.tcl commands to the view. The HTML report is also a <b>session</b>.</li>
</ul>

<h3>5. PowerPoint and slide preview (as in ANSYS SHOTS)</h3>
<ul>
<li><b>Interactive preview</b> on the PowerPoint tab: the \u201ccustom frame\u201d layout \u2013 drag the image, scale it with the corner handles (Shift = proportions),
quick fits (1/2 left, full\u2026), frames saved under your own names. Dragging the image in another layout switches to the custom frame.</li>
<li>The <b>\u201cSlide preview\u201d</b> tab (read only): a sample slide from the current state + \u201cwhat will be made\u201d. The <b>\u201cLive preview\u201d (F6)</b> window:
the same in a separate window, refreshed by itself after every option change.</li>
<li><b>Series (F7)</b>: each metric (+ combined view) \u00d7 camera \u2192 slide with image, legend and statistics table (native PowerPoint shapes).
<b>All slides are previewed</b> before saving (titles, layout, removal, \u201cCreate\u201d / \u201cCancel\u201d); after saving a <b>gallery of the exported slides</b>
with \u201cOpen the presentation\u201d / \u201cOpen folder\u201d. New 16:9 deck or append to an existing one (<code>.bak.pptx</code> backup).</li>
</ul>
"""


# ================================ START ================================
# Uruchomienie w HyperMeshu otwiera okno. Ponowne uruchomienie skryptu przy
# otwartym oknie zamyka stare (ustawienia sa zapisywane) i otwiera nowe -
# bez mnozenia okien i przebiegow na jednym modelu. Referencja do okna
# siedzi w builtins, bo HyperMesh wykonuje kazde uruchomienie w nowej
# przestrzeni nazw.
def main():
    if not HM.ok():
        print("[%s] %s" % (APP_TITLE, T("Brak API HyperMesha (modu\u0142 hm) \u2013 uruchom plik w HyperMesh 2023+: File > Run > Python Script.",
                                         "No HyperMesh API (hm module) \u2013 run the file in HyperMesh 2023+: File > Run > Python Script.")))
        return None
    if qt() is None:
        print("[%s] %s" % (APP_TITLE, T("Brak PyQt5 \u2013 okno niedost\u0119pne (funkcje MQ / DELTA / REPORT dzia\u0142aj\u0105 z konsoli).",
                                         "No PyQt5 \u2013 window unavailable (MQ / DELTA / REPORT work from the console).")))
        return None
    load_settings()
    cleanup_work()                          # stare klatki z poprzednich sesji
    app = app_instance()
    own_app = app is None
    if own_app:                             # poza HyperMeshem (testy)
        app = QtWidgets.QApplication(sys.argv)
    old = getattr(builtins, "_HMQS_STUDIO", None)
    geo = None
    if old is not None:
        try:
            geo = old.geometry()
            old.close()
        except Exception:
            pass
    win = StudioWindow(hm_main_window())
    if geo is not None:
        win.setGeometry(geo)
    builtins._HMQS_STUDIO = win
    win.show()
    win.raise_()
    if GUI_PREFS.get("live_open"):
        try:
            win.live_toggle()
        except Exception:
            pass
    if own_app:
        app.exec_()
    return win


if not os.environ.get("HMQS_NO_GUI"):
    main()
