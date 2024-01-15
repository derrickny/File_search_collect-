import os
import glob
import re

def find_word_documents(folder_path):
    pattern = os.path.join(folder_path, '*.docx')
    word_documents = glob.glob(pattern)
    filenames = [os.path.basename(file) for file in word_documents]
    clean_filenames = [re.sub(r'[^\w\s]', '', name.replace('.docx', '')) for name in filenames]
    return clean_filenames

folder_path = '/Users/nyagaderrick/Downloads/KCB/AUGUST/24TH AUGUST'
filenames = find_word_documents(folder_path)

# Print the list of filenames
print("Word documents in the folder:")
for filename in filenames:
    print(filename)