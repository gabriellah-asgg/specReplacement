import os
import re
from docx import Document
from docx.shared import RGBColor
from docx.enum.text import WD_COLOR_INDEX
import pandas as pd
from tkinter.filedialog import askdirectory, askopenfilename

# Select the directory and files
file_directory = askdirectory(title='Select Folder with documents')
files = os.listdir(file_directory)
paths = [os.path.join(file_directory, file) for file in files]
docs = [f for f in paths if ".docx" in f]

# make new folder
new_folder_path = os.path.join(file_directory, 'Edited Documents')
os.makedirs(new_folder_path, exist_ok=True)

# Select the replacement mapping file
replacement_mapping = askopenfilename(title='Select replacement file')
word_mapping = pd.read_excel(replacement_mapping, 'Word Replacements')
file_mapping = pd.read_excel(replacement_mapping, 'File Replacements')

# Prepare the mapping dictionary
current = [s.lower() for s in list(word_mapping['Original'])]
rep = [s.lower() for s in list(word_mapping['Replacement'])]
replacement_dict = dict(zip(current, rep))

# Prepare the file mapping dictionary
ori_file = [s for s in list(file_mapping['Original'])]
rep_file = [s for s in list(file_mapping['Replacement'])]
file_replacement_dict = dict(zip(ori_file, rep_file))


def replace_and_highlight(doc_path, output_path, replacements):
    doc = Document(doc_path)

    # Iterate through each paragraph
    for para in doc.paragraphs:
        # split the text into separate runs for replacements
        full_text = para.text
        for word, replacement in replacements.items():
            pattern = r'\b' + re.escape(word) + r'\b'  # Use word boundary to match whole words
            if re.search(pattern, full_text, re.IGNORECASE):
                # split the paragraph into runs
                runs = []
                last_index = 0
                for match in re.finditer(pattern, full_text, re.IGNORECASE):
                    start, end = match.span()
                    # Add text before the match as a new run
                    if start > last_index:
                        runs.append((full_text[last_index:start], None))  # No replacement text here
                    # Add the replacement as a new run
                    runs.append((replacement, WD_COLOR_INDEX.YELLOW))  # Add the highlighted replacement
                    last_index = end
                # Add any remaining text after the last match as a new run
                if last_index < len(full_text):
                    runs.append((full_text[last_index:], None))  # No replacement text here

                # Clear the original runs and add the new runs
                para.clear()  # Clear the existing paragraph runs
                for text, highlight in runs:
                    new_run = para.add_run(text)
                    if highlight:
                        new_run.font.highlight_color = highlight

    # Save the modified document
    doc.save(output_path)


# Iterate over the documents and apply the replacement
for file in docs:
    output_path = file_replacement_dict.get(file.split('\\')[1])
    new_file_path = os.path.join(new_folder_path, output_path)
    replace_and_highlight(file, new_file_path, replacement_dict)
    break