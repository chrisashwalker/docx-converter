import os
import sys
from win32com import client

def convert():
    dir = ''
    while dir == '':
        dir = input('Enter the top level folder path containing the .doc files to convert, or enter quit: \n')
    if dir.lower() == 'quit':
        sys.exit()
    word = client.Dispatch('Word.Application')
    word.Visible = False

    success_docs = []
    failed_docs = []
    unwanted_docs = []

    for dirpath, dirs, files in os.walk(dir):
        for filename in files:
            if filename[-4:] == '.doc':
                fullpath = os.path.join(dirpath, filename)
                doc = None
                try:
                    doc = word.Documents.Open(fullpath)
                    doc.SaveAs2(fullpath + 'x', 16)
                    success_docs.append(fullpath)
                except:
                    failed_docs.append(fullpath)
                finally:
                    if doc:
                        doc.Close()

    word.Quit()

    delete_docs = ''
    while not delete_docs.lower() in ['y', 'n']:
        delete_docs = input('Delete the converted .doc files? Y or N: \n')

    if delete_docs == 'y':
        for doc_to_delete in success_docs:
            if os.path.exists(doc_to_delete):
                try:
                    os.remove(doc_to_delete)
                except:
                    unwanted_docs.append(doc_to_delete)

    if len(failed_docs) > 0:
        print('Warning - These docs failed to convert: \n')
        [print(failure + '\n') for failure in failed_docs]

    if len(unwanted_docs) > 0:
        print('Warning - These docs converted but failed to be deleted: \n')
        [print(unwanted + '\n') for unwanted in unwanted_docs]

    input('Done. Press any key to exit.')

if __name__ == '__main__':
    convert()