import os
import shutil
from config import patients_folder as PF, generated_outputs as GO, lifeStyle_template as LT
from text_changes.change_SequencingDetails import replace_text_in_ppt
from image_changes.parallelograms_imageChanges import insert_parallelogram_images

# Define patient names to process
selected_patients = ["KHGLBS782"]


def copy_template_ppt(target_path):
    """
    Copies the actual PPT template to the target location for modifications.
    """
    shutil.copy(LT, target_path)  # Copy the actual template PPT
    print("\n")


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

                insert_parallelogram_images(patient_code)

                print(f"\n✅ Report generated: {output_ppt_path}")


# Execute the process
generate_reports()
