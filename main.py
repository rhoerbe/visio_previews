import os
import re
import sys

import win32com.client
import yaml
from datetime import datetime
from pathlib import Path
import pandas as pd


configfile = os.getenv('CONFIGFILE', default = Path('config_test.yaml'))
with (open(configfile)) as fd:
    config = yaml.load(fd, Loader=yaml.BaseLoader)
    ROOT_FOLDERS = config['ROOT_FOLDERS']
    OUTPUT_FOLDER = Path(config.get('OUTPUT_FOLDER', 'data'))
    EXCLUDE_FILES = config.get('EXCLUDE_FILES', None)
    EXCLUDE_FOLDERS = config.get('EXCLUDE_FOLDERS', None)
    QUALIFYER_NAMES = config.get('QUALIFYER_NAMES', None)

def main():
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    dir_mapping = {'Dir': [], 'Preview': [], 'Visio': []}
    visio = win32com.client.Dispatch("Visio.Application")
    close_open_documents(visio)
    for root_folder in ROOT_FOLDERS:
        process_dir_tree(Path(root_folder), visio, dir_mapping)
    visio.Quit()
    df = pd.DataFrame(dir_mapping)
    df.to_excel(OUTPUT_FOLDER / 'preview2visio_mapping.xlsx')
    print('done.')

def close_open_documents(visio):
    doc_count = visio.Documents.Count
    if doc_count > 0:
        print(f"There are {doc_count} open documents in Visio, close all before starting this program.")
        sys.exit(1)

def process_dir_tree(root_folder, visio, dir_mapping):
    for root, dirs, files in os.walk(root_folder):
        if root in EXCLUDE_FOLDERS:
            continue
        for file in files:
            if file in EXCLUDE_FILES:
                pass
            elif file.endswith(".vsdx"):
                visio_file = Path(root) / file
                generate_preview(visio, visio_file, OUTPUT_FOLDER, dir_mapping)


def generate_preview(visio, visio_file, output_folder, dir_mapping: dict):
    def compute_image_filename() -> Path:
        last_modified = datetime.fromtimestamp(os.path.getmtime(visio_file)).strftime('%Y%m%d')
        prefix = ''
        for pathpart in visio_file.parent.parts:
            if pathpart in QUALIFYER_NAMES:
                prefix = prefix + pathpart + '_'
        filename = prefix + visio_file.stem + "_" + sanitize_filename(pagename) + last_modified + ".png"
        return Path(OUTPUT_FOLDER, filename).resolve()

    try:
        doc = visio.Documents.Open(visio_file.resolve())  # win32com requires absolute path
    except Exception as e:
        print("visio.Documents.Open failed with " + str(visio_file.resolve()), file= sys.stderr)
        print('skipping', file=sys.stderr)
        doc.Close()
        return
    print("processing " + str(visio_file))
    for i in range(1, doc.Pages.Count + 1):
        page = doc.Pages(i)
        page.ResizeToFitContents()
        if doc.Pages.Count > 1:
            print("    page " + str(doc.Pages(i).Name))
            pagename = page.Name + "_"
        else:
            pagename = ''
        image_fn = compute_image_filename()
        try:
            page.Export(image_fn)
        except Exception as e:
            print("visio.Documents.Pages.Export(" + str(image_fn) + ')', file=sys.stderr)
            raise e
    doc.Saved = True
    doc.Close()
    dir_mapping['Dir'].append(OUTPUT_FOLDER)
    dir_mapping['Preview'].append(image_fn.name)
    dir_mapping['Visio'].append(visio_file)

def sanitize_filename(filename: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", filename)


if __name__ == "__main__":
    main()
