import os
import glob
import re
import pandas as pd

def find_word_documents(folder_path):
    pattern = os.path.join(folder_path, '**', '*.docx')
    word_documents = glob.glob(pattern, recursive=True)
    data = {}
    seen_files = set()
    for file in word_documents:
        filename = os.path.basename(file)
        clean_filename = re.sub(r'[^\w\s]', '', filename.replace('.docx', ''))
        # Ignore files with special characters
        if clean_filename != filename.replace('.docx', ''):
            continue
        # Ignore repeated files
        if clean_filename in seen_files:
            continue
        seen_files.add(clean_filename)
        folder_name = os.path.dirname(file)
        # Use the folder containing the .docx file as the heading
        heading = os.path.basename(folder_name)
        if heading not in data:
            data[heading] = []
        data[heading].append(clean_filename)
    return data

folder_path = ''
data = find_word_documents(folder_path)

# Convert the data to a DataFrame and write it to an Excel file
df = pd.DataFrame(dict([ (k,pd.Series(v)) for k,v in data.items() ]))
df.to_excel('NBK.xlsx', index=False)