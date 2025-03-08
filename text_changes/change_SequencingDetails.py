import json
import re
from pptx import Presentation
from pptx_replace import replace_text  # Using pptx-replace

def normalize_keys(json_data):
    """
    Converts JSON keys to valid placeholder names (removing spaces and special characters).
    """
    return {re.sub(r'\s+', '_', k): v for k, v in json_data.items()}


def replace_text_in_ppt(json_path, ppt_template_path, output_path):
    """
    Reads JSON, replaces placeholders in the PPT, and saves the modified file.
    """
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    prs = Presentation(ppt_template_path)

    # Normalize JSON keys
    formatted_data = normalize_keys(data)

    # Replace text in slides
    for key, value in formatted_data.items():
        replace_text(prs, f"{{{key}}}", str(value))

    # Save modified presentation
    prs.save(output_path)

    print("✅ Added Patient Sequencing Details")
