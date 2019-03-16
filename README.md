# doc2img
Extract Images from Documents (*doc*,*docx*) from Partition

Create a virtuaenv and install the package using the requirements.txt:
```
pip install -r requirements.txt
```

Make sure that there's no Folders with the findall.py script before starting.
```
python findall.py
```

## Usage
from forensics perspective: Mount Image with FTK Imager Lite and set 'Mount Method' as 'FileSystem / Read Only'

It starts with asking which Partition you would like

then wait till propmt come back asking for copy selected to directory "results"

Use Windows Explorer Search "*.*" to copy images you like to get from "images"

You'll get all files with selected photos renamed to it's original in directory "input"

## How It Works
Copying all *doc* and *docx* files in a partition to "input" directory with alias names as ID -Using ID was for copying all files even it have duplicate names- then execlude if it have same md5 hash, and generate "input.csv" which contains id, filename, path, md5(file)

Convert All doc files to docx files so we can extarct images from

Extract Images and rename file back to its original


## Thanks
@github/11x256
