import os
import re
from docx import Document
from docx.shared import RGBColor
from docx.enum.text import WD_COLOR_INDEX
import pandas as pd
from tkinter.filedialog import askdirectory, askopenfilename
import win32com.client as win32

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


def replace_and_track_changes(doc_path, output_path, replacements):
    doc_path = os.path.abspath(doc_path)
    doc_path = doc_path.replace('/', '\\')
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False  # Change to True if you want to watch it
    word.DisplayAlerts = False

    doc = word.Documents.Open(doc_path)
    doc.TrackRevisions = True
    doc.ShowRevisions = True

    for word_text, replacement in replacements.items():
        find = doc.Content.Find
        find.Text = word_text
        find.Replacement.Text = replacement
        find.Forward = True
        find.Wrap = 1  # wdFindContinue
        find.Format = False
        find.MatchCase = False
        find.MatchWholeWord = True
        find.MatchWildcards = False
        find.MatchSoundsLike = False
        find.MatchAllWordForms = False

        # wdReplaceAll = 2
        find.Execute(Replace=2)

    doc.SaveAs(output_path)
    doc.Close()
    word.Quit()


# Iterate over the documents and apply the replacement
for file in docs:
    output_path = file_replacement_dict.get(file.split('\\')[1])
    new_file_path = os.path.join(new_folder_path, output_path)
    replace_and_track_changes(file, new_file_path, replacement_dict)
    break
