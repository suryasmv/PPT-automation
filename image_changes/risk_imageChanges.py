import os
import json
import pandas as pd
from pptx import Presentation
from pptx.util import Cm

# Paths
excel_path = r"M:\\Kavya Project\\Regular&ScoringCharts\\ScoringCharts\\BATCH1\\KHAIGHGPPGX27_Scoring_chart.xlsx"
json_path = r"assets/json_mappings/risk_map.json"

# Severity-to-image mapping
image_mapping = {
    "Low": "assets/scales_thermo/Low-1.png",
    "Mild": "assets/scales_thermo/Mild-1.png",
    "Moderate": "assets/scales_thermo/Moderate-1.png",
    "Moderate to High": "images/ModerateHigh-1.png",
}

def extract_conditions_from_ppt(ppt_path):
    """Extracts conditions from slides 7 and 8 (including tables)."""
    ppt = Presentation(ppt_path)
    conditions_in_ppt = set()

    # Check slides 7 and 8 (0-based index: 6, 7)
    for slide in [ppt.slides[6], ppt.slides[7]]:
        for shape in slide.shapes:
            # Search in text frames
            if shape.has_text_frame:
                conditions_in_ppt.add(shape.text.strip())

            # Search in all table cells
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        conditions_in_ppt.add(cell.text.strip())

    return conditions_in_ppt

def replace_images_in_ppt(ppt_path, output_ppt_path, log_file="log.txt"):
    """
    - Finds conditions in slides 7 and 8.
    - Matches them with Excel severity levels.
    - Replaces horizontally aligned images accordingly.
    """

    # Load JSON data
    with open(json_path, 'r', encoding='utf-8') as f:
        json_data = json.load(f)

    # Read Excel file
    df = pd.read_excel(excel_path)

    log_entries = ["Excel Data:\n", str(df.head()), "\nPowerPoint Modifications:\n"]

    # Load PowerPoint
    ppt = Presentation(ppt_path)
    target_slide_numbers = [6, 7]  # Slides 7 and 8 (0-based index)

    # Extract conditions from PPT
    ppt_conditions = extract_conditions_from_ppt(ppt_path)

    for slide_index in target_slide_numbers:
        slide = ppt.slides[slide_index]

        for shape in slide.shapes:
            if shape.has_text_frame:
                for key, condition_text in json_data.items():
                    if condition_text in ppt_conditions:
                        matching_row = df[df["Medical Condition "] == condition_text]

                        if not matching_row.empty:
                            severity_level = None
                            for level in ["Low", "Mild", "Moderate", "Moderate to High"]:
                                if level in matching_row.columns and str(
                                        matching_row[level].values[0]).strip().upper() == "y":
                                    severity_level = level
                                    break

                            if severity_level and severity_level in image_mapping:
                                new_image_path = image_mapping[severity_level]

                                # Find horizontally aligned image
                                for img in slide.shapes:
                                    if img.shape_type == 13 and abs(img.top - shape.top) < Cm(1):
                                        sp = img._element
                                        sp.getparent().remove(sp)

                                        # Insert new image at same position
                                        slide.shapes.add_picture(new_image_path, img.left, img.top, width=img.width, height=img.height)
                                        log_entries.append(f"✅ Replaced image for '{condition_text}' ({severity_level}) on slide {slide_index + 1}")
                                        break
                            else:
                                log_entries.append(f"⚠️ No severity match for '{condition_text}'. Skipping.")

    ppt.save(output_ppt_path)
    log_entries.append(f"\n✅ Updated PowerPoint saved to: {output_ppt_path}")

    # Save log
    with open(log_file, 'w', encoding='utf-8') as log:
        log.write("\n".join(log_entries))
    print(f"Log saved to {log_file}")

