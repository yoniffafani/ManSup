from flask import Flask, render_template, request, send_file
import os
import easyocr
import pandas as pd
from deep_translator import GoogleTranslator
from io import BytesIO
from openpyxl import load_workbook, Workbook
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image, ImageDraw, ImageFont
import tempfile
import warnings

# Suppress harmless MPS warnings
warnings.filterwarnings("ignore", message=".*pin_memory.*")

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'results'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

reader = easyocr.Reader(['ch_sim', 'en'])  # Chinese simplified + English

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['file']
    version = request.form.get("version", "desktop")

    if not file or not file.filename.endswith('.xlsx'):
        return "Please upload a valid .xlsx file"

    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)

    wb_in = load_workbook(filepath)
    ws_in = wb_in.active

    results = []
    temp_dir = tempfile.mkdtemp()

    # Create output workbook
    wb_out = Workbook()
    ws_id = wb_out.active
    ws_id.title = "Indonesian"
    ws_en = wb_out.create_sheet("English")
    ws_table = wb_out.create_sheet("Translation Table")

    # Starting row per sheet
    id_row, en_row = 2, 2

    # Vertical spacing based on version
    row_spacing = 60 if version == "mobile" else 60

    # Process each image in input Excel
    for idx, image_obj in enumerate(ws_in._images, start=1):
        img = Image.open(BytesIO(image_obj._data()))
        img_path = os.path.join(temp_dir, f"{image_obj.anchor._from.row}_{image_obj.anchor._from.col}.png")
        img.save(img_path)

        # OCR read
        ocr_results = reader.readtext(img_path, detail=1)
        combined_text = " ".join([res[1] for res in ocr_results])

        eng_trans = GoogleTranslator(source='auto', target='en').translate(combined_text) or ""
        indo_trans = GoogleTranslator(source='auto', target='id').translate(combined_text) or ""

        results.append({
            "image_file": os.path.basename(img_path),
            "mandarin_text": combined_text,
            "english": eng_trans,
            "indonesian": indo_trans
        })

        # Generate overlay images for each language
        for lang, translation in [('en', eng_trans), ('id', indo_trans)]:
            img_copy = img.copy()
            draw = ImageDraw.Draw(img_copy)
            font = ImageFont.load_default()

            for (bbox, mandarin_text) in [(res[0], res[1]) for res in ocr_results]:
                top_left = bbox[0]
                bottom_right = bbox[2]
                x, y = top_left

                # Translate individual segment safely
                try:
                    translated_segment = GoogleTranslator(source='auto', target=lang).translate(mandarin_text)
                except Exception:
                    translated_segment = ""
                translated_segment = str(translated_segment or "")

                # Get text size
                bbox_text = draw.textbbox((0, 0), translated_segment, font=font)
                text_w = bbox_text[2] - bbox_text[0]
                text_h = bbox_text[3] - bbox_text[1]

                # Draw white background and dark magenta text
                draw.rectangle([x, y - text_h, x + text_w + 4, y + 2], fill="white")
                draw.text((x + 2, y - text_h), translated_segment, fill=(139, 0, 139), font=font)

            overlay_path = os.path.join(temp_dir, f"overlay_{lang}_{os.path.basename(img_path)}")
            img_copy.save(overlay_path)

            excel_img = ExcelImage(overlay_path)

            # Keep original resolution (no resize)
            # Place image vertically below previous one
            if lang == 'id':
                ws_id.add_image(excel_img, f"A{id_row}")
                id_row += row_spacing
            else:
                ws_en.add_image(excel_img, f"A{en_row}")
                en_row += row_spacing

    # Write translation summary table
    ws_table.append(["Image File", "Mandarin Text", "English", "Indonesian"])
    for item in results:
        ws_table.append([
            item["image_file"],
            item["mandarin_text"],
            item["english"],
            item["indonesian"]
        ])

    output_path = os.path.join(RESULT_FOLDER, f'translation_results_{version}.xlsx')
    wb_out.save(output_path)

    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
