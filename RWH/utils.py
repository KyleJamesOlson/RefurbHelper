# utils.py
import logging
from docx.shared import Pt

def replace_in_runs(runs, placeholder, replacement):
    full_text = "".join(run.text for run in runs)
    if placeholder not in full_text:
        return False

    logging.debug(f"Found placeholder {placeholder} in runs")
    new_text = full_text.replace(placeholder, replacement)

    for run in runs:
        run.text = ""

    runs[0].text = new_text
    return True

def process_element(element, replacements):
    for placeholder, value in replacements.items():
        replaced = False
        if placeholder in element.text:
            element.text = element.text.replace(placeholder, value)
            for run in element.runs:
                run.font.name = 'Calibri'
                run.font.size = Pt(11)
            replaced = True
            logging.debug(f"Replaced {placeholder} in paragraph text")
        else:
            replaced = replace_in_runs(element.runs, placeholder, value)

        if replaced:
            logging.debug(f"Successfully replaced {placeholder} with {value}")