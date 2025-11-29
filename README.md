# Dell Laptop Vergleichs-Tool

Ein Python-Tool zum Extrahieren und Vergleichen von Dell Laptop-Daten aus PDF-Dokumenten.

## Funktionen

- Extrahiert Bilder aus PDF-Dateien
- Analysiert Laptop-Spezifikationen
- Berechnet Preis-Rabatte und Bewertungen
- Generiert Vergleichstabellen in mehreren Formaten (CSV, Excel, Markdown)

## Verwendung

```bash
# Virtual Environment aktivieren
source venv/bin/activate

# Script ausführen
python extract_laptop_data.py
```

## Ergebnisse

Das Skript generiert:
- `output/laptops_vergleich.csv` - CSV-Tabelle
- `output/laptops_vergleich.xlsx` - Excel-Datei mit Formatierung
- `output/laptops_vergleich.md` - Markdown-Report mit Bildern
- `images/` - Extrahierte Bilder aus den PDFs

## Aktueller Stand

- **Laptops analysiert:** 11
- **Bilder extrahiert:** 896
- **Durchschnittspreis:** €889.08
- **Beste Angebote:** 5 Laptops mit >60% Rabatt

## Abhängigkeiten

- PyMuPDF (fitz)
- pandas
- Pillow
- openpyxl

## Lizenz

Privates Projekt
