#!/usr/bin/env python3
"""
Script to extract laptop data from PDFs and create a comprehensive table
"""

import os
import sys
from pathlib import Path
import fitz  # PyMuPDF
import pandas as pd
from PIL import Image
import io

# Define base directory
BASE_DIR = Path(__file__).parent
IMAGES_DIR = BASE_DIR / "images"
OUTPUT_DIR = BASE_DIR / "output"

# Create directories if they don't exist
IMAGES_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# Laptop data extracted from PDFs
laptops_data = [
    {
        "Modell": "Dell Latitude 3550",
        "Prozessor": "Intel Core i5-1335U",
        "Prozessor_Kerne": 10,
        "Prozessor_GHz": "1.30/0.90",
        "RAM_GB": 16,
        "RAM_Technologie": "DDR5",
        "SSD_GB": 512,
        "Bildschirmgröße": '15.6"',
        "Auflösung": "1920x1080 (FullHD)",
        "Preis_EUR": 851.99,
        "Neu_Preis_EUR": 1399.00,
        "Gewicht_kg": None,
        "Grafikkarte": None,
        "Bewertung": 4.5,
        "Quelle": "refurbed.de",
        "Quelle_URL": "https://www.refurbed.de/dell-latitude-3550"
    },
    {
        "Modell": "Dell Precision 5550",
        "Prozessor": "Intel Core i7-10850H",
        "Prozessor_Kerne": 6,
        "Prozessor_GHz": "2.70",
        "RAM_GB": 16,
        "RAM_Technologie": "DDR4",
        "SSD_GB": 512,
        "Bildschirmgröße": '15.6"',
        "Auflösung": "1920x1200 (WUXGA)",
        "Preis_EUR": 855.99,
        "Neu_Preis_EUR": 1999.00,
        "Gewicht_kg": None,
        "Grafikkarte": "Nvidia Quadro T1000",
        "Bewertung": 4.5,
        "Quelle": "refurbed.de",
        "Quelle_URL": "https://www.refurbed.de/dell-precision-5550"
    },
    {
        "Modell": "Dell Precision 3561",
        "Prozessor": "Intel Core i9-11950H",
        "Prozessor_Kerne": None,
        "Prozessor_GHz": None,
        "RAM_GB": None,
        "RAM_Technologie": None,
        "SSD_GB": None,
        "Bildschirmgröße": '15"',
        "Auflösung": None,
        "Preis_EUR": 980.00,
        "Neu_Preis_EUR": 3000.00,
        "Gewicht_kg": None,
        "Grafikkarte": None,
        "Bewertung": None,
        "Quelle": "kleinanzeigen.de",
        "Quelle_URL": "https://www.kleinanzeigen.de"
    },
    {
        "Modell": "Dell Precision 7560",
        "Prozessor": "Intel Core i7-11850H",
        "Prozessor_Kerne": 8,
        "Prozessor_GHz": "2.50",
        "RAM_GB": 32,
        "RAM_Technologie": "DDR4",
        "SSD_GB": 1000,
        "Bildschirmgröße": '15.6"',
        "Auflösung": "1920x1080 (FullHD)",
        "Preis_EUR": 888.99,
        "Neu_Preis_EUR": 1719.00,
        "Gewicht_kg": None,
        "Grafikkarte": "RTX A3000",
        "Bewertung": 4.5,
        "Quelle": "refurbed.de",
    "Quelle_URL": "https://www.refurbed.de"
    },
    {
        "Modell": "Dell Precision 7550",
        "Prozessor": "Intel Core i7-10750H",
        "Prozessor_Kerne": 6,
        "Prozessor_GHz": "2.60",
        "RAM_GB": 32,
        "RAM_Technologie": "DDR4",
        "SSD_GB": 512,
        "Bildschirmgröße": '15.6"',
        "Auflösung": "1920x1080 (FullHD)",
        "Preis_EUR": 876.99,
        "Neu_Preis_EUR": 2439.00,
        "Gewicht_kg": None,
        "Grafikkarte": "Quadro T1000",
        "Bewertung": 4.5,
        "Quelle": "refurbed.de",
    "Quelle_URL": "https://www.refurbed.de"
    },
    {
        "Modell": "Dell Latitude 5501",
        "Prozessor": "Intel Core i7-9850H",
        "Prozessor_Kerne": 6,
        "Prozessor_GHz": "2.60",
        "RAM_GB": 32,
        "RAM_Technologie": "DDR4",
        "SSD_GB": 1000,
        "Bildschirmgröße": '15.6"',
        "Auflösung": "1920x1080 (FullHD)",
        "Preis_EUR": 901.99,
        "Neu_Preis_EUR": 1999.00,
        "Gewicht_kg": None,
        "Grafikkarte": "MX150",
        "Bewertung": 4.5,
        "Quelle": "refurbed.de",
    "Quelle_URL": "https://www.refurbed.de"
    },
    {
        "Modell": "Dell Precision 7540",
        "Prozessor": "Intel Core i9-9880H",
        "Prozessor_Kerne": 8,
        "Prozessor_GHz": "2.30",
        "RAM_GB": 32,
        "RAM_Technologie": "DDR4",
        "SSD_GB": 512,
        "Bildschirmgröße": '15.6"',
        "Auflösung": "1920x1080 (FullHD)",
        "Preis_EUR": 923.99,
        "Neu_Preis_EUR": 2299.00,
        "Gewicht_kg": None,
        "Grafikkarte": "Quadro T2000 Mobile",
        "Bewertung": 4.5,
        "Quelle": "refurbed.de",
    "Quelle_URL": "https://www.refurbed.de"
    },
    {
        "Modell": "Dell Precision 5560",
        "Prozessor": "Intel Core i7-10850H",
        "Prozessor_Kerne": 6,
        "Prozessor_GHz": "2.70",
        "RAM_GB": 16,
        "RAM_Technologie": "DDR4",
        "SSD_GB": 512,
        "Bildschirmgröße": '15.6"',
        "Auflösung": "1920x1200 (WUXGA)",
        "Preis_EUR": 925.99,
        "Neu_Preis_EUR": 2999.00,
        "Gewicht_kg": None,
        "Grafikkarte": "T1200 Mobile",
        "Bewertung": 4.5,
        "Quelle": "refurbed.de",
    "Quelle_URL": "https://www.refurbed.de"
    },
    {
        "Modell": "Dell Precision 5560",
        "Prozessor": "Intel Core i7-11850H",
        "Prozessor_Kerne": 8,
        "Prozessor_GHz": "2.50",
        "RAM_GB": 16,
        "RAM_Technologie": "DDR4",
        "SSD_GB": 512,
        "Bildschirmgröße": '15.6"',
        "Auflösung": "1920x1200 (WUXGA)",
        "Preis_EUR": 944.99,
        "Neu_Preis_EUR": 2669.00,
        "Gewicht_kg": None,
        "Grafikkarte": "T1200",
        "Bewertung": 4.5,
        "Quelle": "refurbed.de",
    "Quelle_URL": "https://www.refurbed.de"
    },
    {
        "Modell": "Dell Precision 5560",
        "Prozessor": "Intel Core i5-11500H",
        "Prozessor_Kerne": 6,
        "Prozessor_GHz": "2.40",
        "RAM_GB": 32,
        "RAM_Technologie": "DDR4",
        "SSD_GB": 500,
        "Bildschirmgröße": '15.6"',
        "Auflösung": "1920x1080 (FullHD)",
        "Preis_EUR": 950.00,
        "Neu_Preis_EUR": 2839.00,
        "Gewicht_kg": None,
        "Grafikkarte": "T1200",
        "Bewertung": 4.5,
        "Quelle": "refurbed.de",
    "Quelle_URL": "https://www.refurbed.de"
    },
    {
        "Modell": "Dell Precision 5550",
        "Prozessor": "Intel Core i9-10885H",
        "Prozessor_Kerne": None,
        "Prozessor_GHz": "2.4",
        "RAM_GB": 32,
        "RAM_Technologie": "DDR4",
        "SSD_GB": 1000,
        "Bildschirmgröße": '15.6"',
        "Auflösung": "3840x2400 (WQUXGA)",
        "Preis_EUR": 679.00,
        "Neu_Preis_EUR": 699.00,
        "Gewicht_kg": None,
        "Grafikkarte": "NVIDIA Quadro T2000",
        "Bewertung": None,
        "Quelle": "orbit365.de",
    "Quelle_URL": "https://www.orbit365.de"
    }
]


def extract_images_from_pdf(pdf_path, output_dir):
    """Extract all images from a PDF file"""
    pdf_name = Path(pdf_path).stem
    image_folder = output_dir / pdf_name
    image_folder.mkdir(exist_ok=True)

    extracted_images = []

    try:
        pdf_document = fitz.open(pdf_path)

        for page_number in range(len(pdf_document)):
            page = pdf_document[page_number]
            image_list = page.get_images(full=True)

            for img_index, img in enumerate(image_list):
                xref = img[0]
                base_image = pdf_document.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]

                # Save image
                image_filename = f"{pdf_name}_page{page_number + 1}_img{img_index + 1}.{image_ext}"
                image_path = image_folder / image_filename

                with open(image_path, "wb") as image_file:
                    image_file.write(image_bytes)

                extracted_images.append(str(image_path.relative_to(BASE_DIR)))
                print(f"Extracted: {image_filename}")

        pdf_document.close()
        return extracted_images

    except Exception as e:
        print(f"Error extracting images from {pdf_path}: {e}")
        return []


def evaluate_price(price, new_price):
    """Evaluate if the price is good, fair, or too high"""
    if not price or not new_price:
        return "Unbekannt"

    discount_percent = ((new_price - price) / new_price) * 100

    if discount_percent >= 60:
        return "Sehr gut (>60% Rabatt)"
    elif discount_percent >= 50:
        return "Gut (50-60% Rabatt)"
    elif discount_percent >= 40:
        return "Fair (40-50% Rabatt)"
    elif discount_percent >= 30:
        return "Akzeptabel (30-40% Rabatt)"
    else:
        return "Zu teuer (<30% Rabatt)"


def create_laptop_table(all_extracted_images):
    """Create a comprehensive table with all laptop data"""

    # Add price evaluation to each laptop
    for laptop in laptops_data:
        laptop["Preis_Bewertung"] = evaluate_price(
            laptop["Preis_EUR"],
            laptop["Neu_Preis_EUR"]
        )
        laptop["Rabatt_%"] = round(
            ((laptop["Neu_Preis_EUR"] - laptop["Preis_EUR"]) / laptop["Neu_Preis_EUR"]) * 100, 1
        ) if laptop["Preis_EUR"] and laptop["Neu_Preis_EUR"] else None

        # Find actual image for this laptop
        model_name = laptop['Modell']
        image_path = None
        best_match_score = 0

        for pdf_name, images in all_extracted_images.items():
            # Skip cover PDFs
            if 'cover' in pdf_name.lower() or 'skin' in pdf_name.lower() or 'etsy' in pdf_name.lower():
                continue

            # Calculate match score based on how many model parts are in PDF name
            # and how specific the match is
            score = 0

            # Extract model number (e.g., "3550", "5550", "7560")
            model_parts = model_name.split()
            for part in model_parts:
                if part in pdf_name:
                    # Model numbers get higher weight
                    if part.isdigit() or (len(part) == 4 and part[:2].isdigit()):
                        score += 10
                    else:
                        score += 1

            # Only use this match if it's better than previous matches
            if score > best_match_score and images:
                best_match_score = score
                image_path = images[0]

        laptop["Bild_Link"] = image_path if image_path else "Kein Bild"

    # Create DataFrame
    df = pd.DataFrame(laptops_data)

    # Reorder columns
    columns_order = [
        "Modell", "Bild_Link", "Preis_EUR", "Neu_Preis_EUR", "Rabatt_%", "Preis_Bewertung",
        "Prozessor", "Prozessor_Kerne", "Prozessor_GHz",
        "RAM_GB", "RAM_Technologie", "SSD_GB",
        "Bildschirmgröße", "Auflösung", "Grafikkarte",
        "Gewicht_kg", "Bewertung", "Quelle", "Quelle_URL"
    ]

    df = df[columns_order]

    # Save to CSV
    csv_path = OUTPUT_DIR / "laptops_vergleich.csv"
    df.to_csv(csv_path, index=False, encoding='utf-8-sig')
    print(f"\nCSV gespeichert: {csv_path}")

    # Save to Excel with formatting
    excel_path = OUTPUT_DIR / "laptops_vergleich.xlsx"
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Laptops')

        # Get the workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Laptops']

        # Auto-adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width

    print(f"Excel gespeichert: {excel_path}")

    return df


def create_markdown_report(df, all_extracted_images):
    """Create a comprehensive Markdown report"""
    md_path = OUTPUT_DIR / "laptops_vergleich.md"

    with open(md_path, 'w', encoding='utf-8') as f:
        # Header
        f.write("# 💻 Dell Laptop Vergleich\n\n")
        f.write(f"*📅 Generiert am: {pd.Timestamp.now().strftime('%d.%m.%Y %H:%M:%S')}*\n\n")

        # Summary Statistics
        f.write("## 📊 Zusammenfassung\n\n")
        f.write(f"- 💻 **Laptops analysiert:** {len(df)}\n")
        f.write(f"- 💰 **Durchschnittspreis:** €{df['Preis_EUR'].mean():.2f}\n")
        f.write(f"- 🏆 **Günstigster:** €{df['Preis_EUR'].min():.2f} ({df.loc[df['Preis_EUR'].idxmin(), 'Modell']})\n")
        f.write(f"- 💎 **Teuerster:** €{df['Preis_EUR'].max():.2f} ({df.loc[df['Preis_EUR'].idxmax(), 'Modell']})\n")
        f.write(f"- 📸 **Bilder extrahiert:** {sum(len(imgs) for imgs in all_extracted_images.values())}\n\n")

        # Price Distribution
        f.write("### Preis-Bewertung Verteilung\n\n")
        price_counts = df['Preis_Bewertung'].value_counts()
        for rating, count in price_counts.items():
            f.write(f"- **{rating}:** {count} Laptop(s)\n")
        f.write("\n")

        # Detailed Laptop List
        f.write("## 📋 Detaillierte Laptop-Liste\n\n")
        f.write("---\n\n")

        # Sort by price
        df_sorted = df.sort_values('Preis_EUR')

        for idx, laptop in df_sorted.iterrows():
            # Create anchor ID from model name
            anchor_id = laptop['Modell'].lower().replace(' ', '-')

            # Laptop header with model name and anchor
            f.write(f"### <a id=\"{anchor_id}\"></a>{laptop['Modell']}\n\n")

            # Find images for this laptop - improved matching
            model_images = []
            model_name = laptop['Modell']
            best_match_score = 0
            best_match_images = []

            # Extract key identifiers from model name
            model_parts = model_name.split()

            for pdf_name, images in all_extracted_images.items():
                # Skip cover PDFs
                if 'cover' in pdf_name.lower() or 'skin' in pdf_name.lower() or 'etsy' in pdf_name.lower():
                    continue

                # Calculate match score
                score = 0
                for part in model_parts:
                    if part in pdf_name:
                        # Model numbers get higher weight
                        if part.isdigit() or (len(part) == 4 and part[:2].isdigit()):
                            score += 10
                        else:
                            score += 1

                # Keep track of best match
                if score > best_match_score and images:
                    best_match_score = score
                    best_match_images = images[:5]  # Take first 5 images

            model_images = best_match_images

            # Display first image if available
            if model_images:
                # Fix path for markdown file in output/ directory
                img_path = f"../{model_images[0]}"
                f.write(f"![{laptop['Modell']}]({img_path})\n\n")
            else:
                # If no images found, show placeholder
                f.write(f"*📷 Keine spezifischen Produktbilder verfügbar*\n\n")

            # Price information with badge
            preis_emoji = {
                "Sehr gut (>60% Rabatt)": "🟢",
                "Gut (50-60% Rabatt)": "🟢",
                "Fair (40-50% Rabatt)": "🟡",
                "Akzeptabel (30-40% Rabatt)": "🟡",
                "Zu teuer (<30% Rabatt)": "🔴"
            }
            emoji = preis_emoji.get(laptop['Preis_Bewertung'], "⚪")

            f.write(f"**Preis:** €{laptop['Preis_EUR']:.2f} {emoji} *{laptop['Preis_Bewertung']}*\n\n")
            f.write(f"**Originalpreis:** €{laptop['Neu_Preis_EUR']:.2f} | **Rabatt:** {laptop['Rabatt_%']}%\n\n")

            # Technical specifications table
            f.write("#### Technische Daten\n\n")
            f.write("| Spezifikation | Details |\n")
            f.write("|--------------|----------|\n")
            f.write(f"| **Prozessor** | {laptop['Prozessor']} |\n")
            if laptop['Prozessor_Kerne']:
                f.write(f"| **Kerne** | {laptop['Prozessor_Kerne']} |\n")
            if laptop['Prozessor_GHz']:
                f.write(f"| **Taktfrequenz** | {laptop['Prozessor_GHz']} GHz |\n")
            if laptop['RAM_GB']:
                ram_text = f"{laptop['RAM_GB']} GB"
                if laptop['RAM_Technologie']:
                    ram_text += f" {laptop['RAM_Technologie']}"
                f.write(f"| **RAM** | {ram_text} |\n")
            if laptop['SSD_GB']:
                f.write(f"| **SSD** | {laptop['SSD_GB']} GB |\n")
            if laptop['Bildschirmgröße']:
                f.write(f"| **Bildschirm** | {laptop['Bildschirmgröße']} |\n")
            if laptop['Auflösung']:
                f.write(f"| **Auflösung** | {laptop['Auflösung']} |\n")
            if laptop['Grafikkarte']:
                f.write(f"| **Grafikkarte** | {laptop['Grafikkarte']} |\n")
            if laptop['Gewicht_kg']:
                f.write(f"| **Gewicht** | {laptop['Gewicht_kg']} kg |\n")
            if pd.notna(laptop['Bewertung']):
                f.write(f"| **Bewertung** | {'⭐' * int(laptop['Bewertung'])} ({laptop['Bewertung']}/5) |\n")
            # Make source clickable
            if laptop.get('Quelle_URL'):
                f.write(f"| **Quelle** | [🔗 {laptop['Quelle']}]({laptop['Quelle_URL']}) |\n")
            else:
                f.write(f"| **Quelle** | {laptop['Quelle']} |\n")
            f.write("\n")

            # Additional images if available
            if len(model_images) > 1:
                f.write("<details>\n")
                f.write("<summary>📸 Weitere Bilder</summary>\n\n")
                for img in model_images[1:]:
                    img_path = f"../{img}"
                    f.write(f"![{laptop['Modell']}]({img_path})\n\n")
                f.write("</details>\n\n")

            f.write("---\n\n")

        # Comparison Table
        f.write("## 📊 Schnellvergleich Tabelle\n\n")
        f.write("| Modell | Preis | Rabatt | CPU | RAM | SSD | GPU | Bewertung |\n")
        f.write("|--------|-------|--------|-----|-----|-----|-----|----------|\n")

        for idx, laptop in df_sorted.iterrows():
            ram = f"{laptop['RAM_GB']}GB" if laptop['RAM_GB'] else "N/A"
            ssd = f"{laptop['SSD_GB']}GB" if laptop['SSD_GB'] else "N/A"
            gpu = laptop['Grafikkarte'] if laptop['Grafikkarte'] else "N/A"
            rating_emoji = preis_emoji.get(laptop['Preis_Bewertung'], "⚪")

            # Shorten processor name
            cpu_short = laptop['Prozessor'].replace('Intel Core ', '').replace('Intel ', '')

            # Create anchor link
            anchor_id = laptop['Modell'].lower().replace(' ', '-')
            model_link = f"[{laptop['Modell']}](#{anchor_id})"

            f.write(f"| {model_link} | €{laptop['Preis_EUR']:.2f} {rating_emoji} | {laptop['Rabatt_%']}% | {cpu_short} | {ram} | {ssd} | {gpu} | {laptop['Preis_Bewertung']} |\n")

        f.write("\n---\n\n")
        f.write("*Legende: 🟢 = Sehr gut/Gut | 🟡 = Fair/Akzeptabel | 🔴 = Zu teuer*\n")

    print(f"Markdown Report gespeichert: {md_path}")
    return md_path


def main():
    """Main function"""
    print("=" * 80)
    print("DELL LAPTOP DATA EXTRACTOR")
    print("=" * 80)

    # Extract images from all PDFs (main directory and cover subdirectory)
    pdf_files = list(BASE_DIR.glob("*.pdf"))
    cover_pdf_files = list((BASE_DIR / "cover").glob("*.pdf")) if (BASE_DIR / "cover").exists() else []
    all_pdf_files = pdf_files + cover_pdf_files

    print(f"\n{len(pdf_files)} PDF-Dateien im Hauptverzeichnis gefunden")
    print(f"{len(cover_pdf_files)} PDF-Dateien im cover/ Verzeichnis gefunden")
    print(f"Gesamt: {len(all_pdf_files)} PDF-Dateien\n")

    print("Schritt 1: Bilder aus PDFs extrahieren...")
    print("-" * 80)

    all_extracted_images = {}
    for pdf_file in all_pdf_files:
        print(f"\nVerarbeite: {pdf_file.name}")
        images = extract_images_from_pdf(pdf_file, IMAGES_DIR)
        all_extracted_images[pdf_file.name] = images
        print(f"  → {len(images)} Bilder extrahiert")

    print("\n" + "-" * 80)
    print(f"Gesamt: {sum(len(imgs) for imgs in all_extracted_images.values())} Bilder extrahiert")

    # Create laptop comparison table
    print("\n" + "=" * 80)
    print("Schritt 2: Laptop-Vergleichstabelle erstellen...")
    print("-" * 80)

    df = create_laptop_table(all_extracted_images)

    # Create Markdown report
    print("\n" + "=" * 80)
    print("Schritt 3: Markdown-Report erstellen...")
    print("-" * 80)

    md_path = create_markdown_report(df, all_extracted_images)

    print("\n" + "=" * 80)
    print("ZUSAMMENFASSUNG")
    print("=" * 80)
    print(f"Laptops analysiert: {len(df)}")
    print(f"Durchschnittspreis: €{df['Preis_EUR'].mean():.2f}")
    print(f"Günstigster: €{df['Preis_EUR'].min():.2f} ({df.loc[df['Preis_EUR'].idxmin(), 'Modell']})")
    print(f"Teuerster: €{df['Preis_EUR'].max():.2f} ({df.loc[df['Preis_EUR'].idxmax(), 'Modell']})")
    print(f"\nPreis-Bewertung:")
    print(df['Preis_Bewertung'].value_counts())
    print(f"\nMarkdown Report: {md_path}")

    print("\n" + "=" * 80)
    print("FERTIG!")
    print("=" * 80)


if __name__ == "__main__":
    main()
