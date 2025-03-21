import os
import shutil
from pptx import Presentation
from config import patients_folder as PF, generated_outputs as GO, lifeStyle_template as LT
from text_changes.change_SequencingDetails import replace_text_in_ppt
from image_changes.parallelograms_imageChanges import insert_parallelogram_images
from image_changes.diet_imageChanges import add_diet_images  
from text_changes.change_VitaminDetails import update_vitamin_details
from config import VITAMIN_SHEET_FILE, GENERATED_OUTPUTS

# Define patient names to process
selected_patients = ["KHGLBS782"]

def copy_template_ppt(target_path):
    """
    Copies the actual PPT template to the target location for modifications.
    """
    shutil.copy(LT, target_path)  # Copy the actual template PPT
    print("\n")

# Define Y-coordinate bounds (in EMUs, 1 cm = 360000 EMUs)
START_Y = 9 * 360000  
MAX_Y = 26 * 360000  

def has_content(slide):
    """Checks if a slide contains any shapes (images or textboxes) in the valid area."""
    for shape in slide.shapes:
        if hasattr(shape, "top") and START_Y <= shape.top <= MAX_Y:
            return True
    return False

def delete_empty_slides(output_ppt_path):
    """Deletes all empty slides in the PowerPoint if they have no images or textboxes within the valid area."""
    prs = Presentation(output_ppt_path)
    
    # Identify empty slides (across entire PPT)
    empty_slide_indexes = [i for i in range(len(prs.slides)) if not has_content(prs.slides[i])]

    if empty_slide_indexes:
        for i in sorted(empty_slide_indexes, reverse=True):
            xml_slides = prs.slides._sldIdLst
            xml_slides.remove(xml_slides[i])

        prs.save(output_ppt_path)
        print(f"✅ Empty slides deleted across the entire PPT => {empty_slide_indexes}")
    else:
        print(f"✅ No empty slides found in the PPT.")

    return empty_slide_indexes  # Return the list of deleted slides


def generate_reports():
    """
    Generates reports for specific patients by copying the actual PPT template,
    then processing replacements before final saving.
    """
    for file_name in os.listdir(PF):
        if file_name.endswith(".json"):
            patient_code = file_name.split(".")[0]  # Extract patient code

            if patient_code in selected_patients:
                json_path = os.path.join(PF, file_name)
                output_ppt_path = os.path.join(GO, f"{patient_code}_report.pptx")

                # Step 1: Copy the actual template PPT to the output folder
                copy_template_ppt(output_ppt_path)

                # Step 2: Process text replacement using change_SequencingDetails.py
                replace_text_in_ppt(json_path, output_ppt_path, output_ppt_path)

                # Step 3: Insert parallelogram images
                insert_parallelogram_images(patient_code)

                # Step 4: Insert diet images
                add_diet_images(output_ppt_path, patient_code)

                # Step 5: Update vitamin details
                update_vitamin_details(patient_code)

                delete_empty_slides(output_ppt_path)

                print(f"\n✅ Report generated: {output_ppt_path}\n")


# Execute the process
generate_reports()
