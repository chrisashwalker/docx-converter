import os
import sys
from tqdm import tqdm
from win32com import client


class Converter:
    def convert(self, filepath, new_ext):
        try:
            new_filepath = ''
            return new_filepath
        finally:
            raise RuntimeError('"Converter.convert" method has not been implemented.')


# Uses Microsoft Word to change file formats
class WordConverter(Converter):
    def __init__(self):
        # The Word API has an enumeration of file formats. We'll define those that we may need.
        self.file_types = {
            '.doc': 0,
            '.docx':  16
        }
    
    def __enter__(self):
        # Create instance of COM object
        self.client = client.Dispatch('Word.Application')
        self.client.Visible = False
        return self

    def __exit__(self, type, value, traceback):
        self.client.Quit()

    # Overriding superclass method
    def convert(self, filepath, new_ext):
        doc = None
        try:
            doc = self.client.Documents.Open(filepath)
            new_filepath = os.path.splitext(filepath)[0] + new_ext
            doc.SaveAs2(new_filepath, self.file_types[new_ext])
            return new_filepath
        finally:
            if doc:
                doc.Close()


def deleteFiles(files):
    successes = []
    failures = []
    for file in tqdm(iterable=files, desc='Tidying up...', unit='files'):
        if os.path.exists(file):
            try:
                os.remove(file)
                if not os.path.exists(file):
                    successes.append(file)
                else:
                    raise Exception('Failed to delete: ' + file)
            except:
                failures.append(file)
    return successes, failures

def getFiles(folder, file_ext):
    matching_files = []
    for dirpath, dirnames, files in os.walk(folder):
        for filename in files:
            if os.path.splitext(filename)[1] == file_ext:
                filepath = os.path.join(dirpath, filename)
                matching_files.append(filepath)
    return matching_files

def askForFolder():
    while True:    
        folder = input('Enter the top level folder path containing the files to convert, or enter quit: \n')
        if folder.lower() == 'quit':
            sys.exit()
        if folder != '':
            break
    return folder

def askIfDeleting():
    while True:    
        deleting = input('Delete the converted files? Y or N: \n')
        if deleting.lower() in ['y', 'n']:
            return deleting.lower() == 'y'

def main():
    # Word file specifics
    old_ext = '.doc'
    new_ext = '.docx'
    converter = WordConverter()

    processed_files = []
    rejected_files = []
    unwanted_files = []

    folder = askForFolder()
    deleting_after_conversion = askIfDeleting(); 
    docs_to_convert = getFiles(folder, old_ext)

    with converter:
        for file in tqdm(iterable=docs_to_convert, desc='Converting...', unit='files'):
            try:
                converted_filepath = converter.convert(file, new_ext)
                if os.path.exists(converted_filepath):
                    processed_files.append(file)
                else:
                    raise Exception('Failed to convert: ' + file)
            except:
                rejected_files.append(file)

    if deleting_after_conversion:
        deleted_files, unwanted_files = deleteFiles(processed_files)

    if len(rejected_files) > 0:
        print('Warning - These files failed to convert: \n')
        [print(failure + '\n') for failure in rejected_files]

    if len(unwanted_files) > 0:
        print('Warning - These files converted but failed to be deleted: \n')
        [print(unwanted + '\n') for unwanted in unwanted_files]

    input('Done. Press any key to exit.')


# Program entry point
if __name__ == '__main__':
    main()