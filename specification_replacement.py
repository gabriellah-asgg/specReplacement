import os
from docx import Document
from docx.shared import RGBColor
from docx.enum.text import WD_COLOR_INDEX
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askdirectory

#file_directory = r"T:\JRD MTA\MTA Sample Sections - Copy"
file_directory = askdirectory(title='Select Folder with documents')
files = os.listdir(file_directory)
paths = [os.path.join(file_directory, file) for file in files]
docs = [f for f in paths if ".docx" in f]
excel = [ex for ex in paths if ".xlsx" in ex]

#replacement_mapping = r"T:\JRD MTA\MTA Sample Sections - Copy\specification_mapping.xlsx"
replacement_mapping = askdirectory(title='Select replacement file')
mapping = pd.read_excel(replacement_mapping)

current = [s.lower() for s in list(mapping['Current Word'])]
rep = [s.lower() for s in list(mapping['Replacement'])]

replacement_dict = dict(zip(current, rep))


def replace_and_highlight(doc_path, output_path, replacements):
    doc = Document(doc_path)

    # Iterate through each paragraph
    for para in doc.paragraphs:
        for run in para.runs:
            original_text = run.text
            lower_text = original_text.lower()  # Convert run text to lowercase for comparison

            for word, replacement in replacements.items():
                if word in lower_text:
                    # Preserve original case where possible
                    replaced_text = lower_text.replace(word, replacement)

                    # Apply changes to the run text
                    run.text = replaced_text

                    # Highlight the modified text
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW

                    # Save the modified document
    doc.save(output_path)


for file in docs:
    replace_and_highlight(file, "test.docx", replacement_dict)
    break
