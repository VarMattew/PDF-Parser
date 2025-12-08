![Python](https://img.shields.io/badge/Python-3.8%2B-3776AB?style=flat&logo=python&logoColor=white) ![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey)

![GUI](https://img.shields.io/badge/GUI-Tkinter-red?style=flat)
![Status](https://img.shields.io/badge/Status-Development-orange?style=flat) 
# PDF Parser - PDF Adatkinyerő

Modern, minimalista asztali alkalmazás, amely automatizálja az adatok kinyerését PDF fájlokból. A program egy kiválasztott mappában végigolvassa az összes PDF-et, megkeresi a specifikus adatokat, és egy formázott Excel vagy CSV fájlba menti azokat.

## Funkciók

* **Modern Felület:** Letisztult, "Dark Mode" stílusú GUI (CustomTkinter).
* **Intelligens Keresés:** Mappa szintű beolvasás és automatikus PDF szűrés.
* **Valós idejű visszajelzés:** Progress bar és állapotjelző ablak a feldolgozás alatt.
* **Könnyűsúlyú:** Pandas nélkül készült, optimalizált Excel (`openpyxl`) és CSV export.
* **Biztonságos:** Validáció a hiányzó adatok és mappák ellen.

## Technológiai háttér

A projekt Python nyelven készült, az alábbi könyvtárak felhasználásával:
* `customtkinter` (Modern felhasználói felület)
* `openpyxl` (Excel generálás formázott fejléccel)
* `os`, `csv` (Fájlkezelés)

## Telepítés és Futtatás

1.  **Klónozd a repót vagy töltsd le a fájlokat.**
2.  **Telepítsd a szükséges csomagokat:**
    ```bash
    pip install customtkinter openpyxl
    ```
    *(Megjegyzés: Ha PDF olvasót is használsz a logikában, pl. PyMuPDF, azt is telepítsd: `pip install pymupdf`)*

3.  **Futtasd a programot:**
    ```bash
    python main.py
    ```