# Dell Laptop Vergleich

Automatisches Tool zur Extraktion und Analyse von Dell Laptop-Daten aus PDFs.

## Features

- 📄 **PDF-Import**: Automatische Extraktion von Laptop-Daten aus PDF-Dateien
- 🖼️ **Bild-Extraktion**: Speichert alle Produktbilder in organisierten Verzeichnissen
- 📊 **Vergleichstabellen**: Generiert CSV, Excel und Markdown-Berichte
- 💰 **Preis-Bewertung**: Automatische Bewertung der Preise basierend auf Rabatten
- ⭐ **Bewertungen**: Zeigt Produktbewertungen an

## Generierte Dateien

Das Skript erzeugt automatisch:

- `output/laptops_vergleich.csv` - CSV-Format für Datenanalyse
- `output/laptops_vergleich.xlsx` - Excel-Format mit Formatierung
- `output/laptops_vergleich.md` - Markdown-Bericht mit Bildern
- `images/` - Verzeichnis mit extrahierten Produktbildern

## Installation

```bash
# Virtual Environment erstellen
python3 -m venv venv

# Virtual Environment aktivieren
source venv/bin/activate  # Linux/Mac
# oder
venv\Scripts\activate  # Windows

# Abhängigkeiten installieren
pip install pymupdf pandas openpyxl pillow
```

## Verwendung

```bash
# Skript ausführen
python extract_laptop_data.py
```

Das Skript verarbeitet automatisch alle PDF-Dateien im Hauptverzeichnis.

## Preis-Bewertung

Das System bewertet Preise basierend auf dem Rabatt:

- 🟢 **Sehr gut**: >60% Rabatt
- 🟢 **Gut**: 50-60% Rabatt
- 🟡 **Fair**: 40-50% Rabatt
- 🟡 **Akzeptabel**: 30-40% Rabatt
- 🔴 **Zu teuer**: <30% Rabatt

## Ergebnisse

Aktuell analysierte Laptops: **11**
- Durchschnittspreis: **€889.08**
- Günstigster: **€679.00**
- Teuerster: **€980.00**
- Extrahierte Bilder: **286**

## Lizenz

Open Source - Frei verwendbar
