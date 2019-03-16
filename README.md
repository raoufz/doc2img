# doc2img
Extract Images from Documents (doc,docx) from Partition

Create a virtuaenv and install the package using the requirements.txt:
```
pip install -r requirements.txt
```

Make sure that there's no Folders with the findall.py script before starting.
```
python findall.py
```

# How It works
from forensics perspective: Mount Image with FTK Imager Lite and set 'Mount Method' as 'FileSystem / Read Only'
It starts with asking which Partition you would like
then wait till propmt come back asking for copy selected to directory "results"
You'll get all files with selected photos renamed to it's original in directory "input"
There is "input.csv" which contains file, path, md5(file)

# Thanks
