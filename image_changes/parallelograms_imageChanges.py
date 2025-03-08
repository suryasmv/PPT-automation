import os
import pandas as pd
from pptx import Presentation
from pptx.util import Cm
from config import scoring_charts, input_parallelograms, generated_outputs as GO

# Image sizes
IMAGE_SIZES = {
    "Moderate to High": (Cm(9.35), Cm(19.43)),
    "Moderate": (Cm(7), Cm(19.43)),
    "Mild": (Cm(7), Cm(19.43)),
    "Low": (Cm(5), Cm(19.43))
}

# Slide parameters
START_X = Cm(0.7)
START_Y = Cm(4.5)
MAX_Y = Cm(27)
SPACING = Cm(0.5)
START_SLIDE_INDEX = 8
END_SLIDE_INDEX = 29
SEVERITY_ORDER = ["Moderate to High", "Moderate", "Mild", "Low"]  # Priority order


def find_scoring_chart(patient_id):
    """Finds the scoring chart Excel file for a given patient ID."""
    for root, _, files in os.walk(scoring_charts):
        for file in files:
            if file.startswith(f"{patient_id}_Scoring_chart") and file.endswith(".xlsx"):
                return os.path.join(root, file)
    return None


def extract_severity_conditions(excel_path):
    """Extracts conditions and severity levels from the Excel file."""
    df = pd.read_excel(excel_path)
    results = []

    for severity in SEVERITY_ORDER:  # Maintain priority order
        for _, row in df.iterrows():
            if severity in df.columns and str(row.get(severity, '')).strip().lower() == 'y':
                condition_name = row["Medical Condition "].replace(" ", "_")  # Match filename
                results.append((severity, condition_name))

    return results


def find_condition_image(severity, condition):
    """Finds the image path for a given severity and condition."""
    severity_path = os.path.join(input_parallelograms, severity)
    if os.path.exists(severity_path):
        for file in os.listdir(severity_path):
            if file.lower().startswith(condition.lower()) and file.endswith(".png"):
                return os.path.join(severity_path, file)
    return None


def has_content(slide):
    """Checks if a slide contains any shapes (images or textboxes) in the valid area."""
    for shape in slide.shapes:
        if hasattr(shape, "left") and START_Y <= shape.top <= MAX_Y:
            return True
    return False


def delete_empty_slides(output_ppt_path):
    """Deletes empty slides in the range 9-30 if they have no images or textboxes within the valid area."""
    prs = Presentation(output_ppt_path)
    empty_slide_indexes = [i for i in range(START_SLIDE_INDEX, END_SLIDE_INDEX + 1) if not has_content(prs.slides[i])]

    if empty_slide_indexes:
        for i in sorted(empty_slide_indexes, reverse=True):
            xml_slides = prs.slides._sldIdLst
            xml_slides.remove(xml_slides[i])
        prs.save(output_ppt_path)
        print(f"✅ Empty Parallelogram slides deleted in the range 9 to 30 => {empty_slide_indexes}")
    else:
        print(f"✅ No empty Parallelogram slides found in the range 9 to 30.")


def insert_parallelogram_images(patient_id):
    """Inserts images into PowerPoint slides based on the patient's conditions."""
    output_ppt_path = os.path.join(GO, f"{patient_id}_report.pptx")

    if not os.path.exists(output_ppt_path):
        print(f"❌ Report not found for patient {patient_id}")
        return

    prs = Presentation(output_ppt_path)
    scoring_chart = find_scoring_chart(patient_id)
    if not scoring_chart:
        print(f"❌ No scoring chart found for patient {patient_id}")
        return

    conditions = extract_severity_conditions(scoring_chart)
    if not conditions:
        print(f"⚠️ No conditions found for patient {patient_id}")
        return

    current_slide_index = START_SLIDE_INDEX
    current_y = START_Y
    slide = prs.slides[current_slide_index]

    for severity in SEVERITY_ORDER:  # Ensure priority order
        for condition_severity, condition in conditions:
            if condition_severity == severity:  # Insert images in the correct order
                image_path = find_condition_image(severity, condition)
                if not image_path:
                    print(f"❌ Image not found for {condition} ({severity})")
                    continue

                img_height, img_width = IMAGE_SIZES[severity]

                if current_y + img_height > MAX_Y:
                    if current_slide_index < END_SLIDE_INDEX:
                        current_slide_index += 1
                        slide = prs.slides[current_slide_index]
                        current_y = START_Y
                    else:
                        print("⚠️ Slide limit reached, stopping image insertion.")
                        break

                slide.shapes.add_picture(image_path, START_X, current_y, width=img_width, height=img_height)
                current_y += img_height + SPACING

    prs.save(output_ppt_path)
    print(f"✅ Parallelogram Images inserted for patient {patient_id} into {output_ppt_path}")

    # Call delete_empty_slides after inserting images
    delete_empty_slides(output_ppt_path)
