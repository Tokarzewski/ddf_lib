from zipfile import ZipFile
from pathlib import Path

folderpath = Path(r"./samples/")

ddf_samples = list(folderpath.glob("**/*.ddf"))

for file in ddf_samples:
    print(file)
    with ZipFile(file, "r") as zip_object:
        new_folder = file.parent / file.stem
        zip_object.extractall(path=new_folder)
