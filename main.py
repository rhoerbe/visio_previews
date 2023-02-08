import os
import sys

import win32com.client
import yaml
from datetime import datetime
from pathlib import Path
import pandas as pd


configfile = os.getenv('CONFIGFILE', default = Path('config.yaml'))
with (open(configfile)) as fd:
    config = yaml.load(fd, Loader=yaml.BaseLoader)
    ROOT_FOLDER = Path(config.get('ROOT_FOLDER', None))
    OUTPUT_FOLDER = Path(config.get('OUTPUT_FOLDER', 'data'))
    EXCLUDE_FILES = config.get('EXCLUDE_FILES', None)
    EXCLUDE_FOLDERS = config.get('EXCLUDE_FOLDERS', None)
    QUALIFYER_NAMES = config.get('QUALIFYER_NAMES', None)

def main():
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    dir_mapping = {}
    for root, dirs, files in os.walk(ROOT_FOLDER):
        if root in EXCLUDE_FOLDERS:
            continue
        for file in files:
            if file in EXCLUDE_FILES:
                pass
            elif file.endswith(".vsdx"):
                visio_file = Path(root) / file
                generate_preview(visio_file, OUTPUT_FOLDER, dir_mapping)
    df = pd.DataFrame(dir_mapping, index=None)
    df.to_excel(OUTPUT_FOLDER / 'file2dir_mapping')


def generate_preview(visio_file, output_folder, dir_mapping: dict):
    def compute_image_filename() -> str:
        prefix = ''
        last_modified = datetime.fromtimestamp(os.path.getmtime(visio_file)).strftime('%Y%m%d')
        for name in QUALIFYER_NAMES:
            prefix = name + '_'
        return Path(OUTPUT_FOLDER, prefix + visio_file.stem + "_" + page.Name + "_" + last_modified + ".png").resolve()

    visio = win32com.client.Dispatch("Visio.Application")
    try:
        doc = visio.Documents.Open(visio_file.resolve())  # win32com requires absolute path
    except Exception:
        print("visio.Documents.Open failed with " + visio_file.resolve(), file= sys.stderr)
    print("processing " + str(visio_file))
    for i in range(1, doc.Pages.Count + 1):
        page = doc.Pages(i)
        image_fn = compute_image_filename()
        page.Export(image_fn)
        #page.Export(Path('test.png').resolve())
        if doc.Pages.Count > 1:
            print("exporting page " + str(doc.Pages(i).Name))
    doc.Close()
    visio.Quit()
    dir_mapping[image_fn] = visio_file

if __name__ == "__main__":
    main()
