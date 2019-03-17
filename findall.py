import shutil
from glob import glob
import re
import os
import win32com.client as win32
from win32com.client import constants
from func_timeout import func_timeout, FunctionTimedOut
import pythoncom
import hashlib
from pathlib import Path
import docx2txt
import time

def startInitalize():
	os.system("taskkill /f /IM WINWORD.exe")
	dirs=['images','input','results']
	if os.path.isfile("input.csv"):
		os.remove("input.csv") 
	for i in dirs:
		os.makedirs(i)

def copy2dest(fileTypes,destFolder):
	print('Copying all document file to "input" Folder\n')
	id=0
	hashlist = []
	f=open("input.csv",'w',encoding="utf-8")
	for type in fileTypes:
		for file in type:
			fileHash = hashlib.md5(open(file,'rb').read()).hexdigest()
			if fileHash in hashlist:
				continue
			hashlist.append(fileHash)
			if "docx" in file:
				shutil.copy(file, destFolder+str(id)+".docx")
			else:
				shutil.copy(file, destFolder+str(id)+".doc")
			a,b=os.path.split(file)
			f.write((str(id)+","+b+","+a+","+fileHash+"\n"))
			id+=1
	f.close()

def save_as_docx(absPath):
	pythoncom.CoInitialize()
	# Opening MS Word
	word = win32.gencache.EnsureDispatch('Word.Application')
	doc = word.Documents.Open(absPath)
	doc.Activate ()

	# Rename path with .docx
	new_file_abs = os.path.abspath(absPath)
	new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)
	# Save and Close
	word.ActiveDocument.SaveAs(
		new_file_abs, FileFormat=constants.wdFormatXMLDocument
	)
	doc.Close()
	#doc.Close(False)
	
def doc2docx(destFolder):
	print("Converting doc files to docx Files\n")
	paths = glob(destFolder+'*.doc')
	c=0
	for absPath in paths:
		if c>250:
			os.system("taskkill /f /IM WINWORD.exe")
			time.sleep(5)
			c=0
		try:
			func_timeout(15, save_as_docx, args=(absPath,))
		except FunctionTimedOut:
			print("save_as_docx('"+absPath+") could not complete within 15 seconds and was terminated\n")
		except Exception as e:
			print(absPath,": ",str(e))
			pass
		c+=1


def extract_image(ABS_PATH):
	print("Extracting Images from docx Files\n")
	Flag=True
	docxFiles=glob(os.path.join(ABS_PATH, "input/")+"*.docx")
	for item in docxFiles:
		directory = os.path.join(ABS_PATH, "images/%s" % (os.path.splitext(os.path.basename(item))[0]))
		if not os.path.exists(directory):
			os.makedirs(directory)
		try:
			docx2txt.process(item, directory)
		except Exception as e:
			print(item,": " ,str(e))
			pass
	for path in Path('images').glob('**/*.*'):
		path.parts
		os.rename(path,path.parts[0]+"\\"+path.parts[1]+"\\"+path.parts[1]+"-"+path.parts[2])
	while Flag:
		confirm=input('Select Images and Copy it to Folder "Results" and print CONFIRM:\n')
		if confirm=="CONFIRM":
			Flag=False
		
def img2doc():
	print('Rename files in "input" folder to its original name\n')
	f=open("input.csv",'r',encoding='utf-8')
	l=[]
	for path in Path('results').glob('**/*.*'):
		if (path.parts[1].split("-")[0]) not in l:
			l.append(path.parts[1].split("-")[0])

	for line in f.readlines():
		for id in l:
			if str(id)==line.split(",")[0]:
				newName=line.split(",")[1]

				if not newName.startswith('~$'):
					try:
						if os.path.isfile("input\\"+id+".doc"):
							os.rename("input\\"+id+".doc","input\\"+newName)
						else:
							os.rename("input\\"+id+".docx","input\\"+newName)
					except:
						pass
	f.close()
			
			
def main():
	partition=input('Write drive letter you want to search: ')

	ABS_PATH = os.path.dirname(os.path.realpath(__file__))
	# Search for all doc extensions you like
	fileTypes = [glob(e, recursive=True) for e in [partition+":\\**\\*.doc",partition+":\\**\\*.docx"]]
	destFolder = os.path.join(ABS_PATH+"\\input\\")
	
	#Remove Old folders from last Iteration, and Creating empty directories
	startInitalize()
	# Copy selected files to destination folder
	copy2dest(fileTypes,destFolder)
	
	# Convert all doc to docx files
	doc2docx(destFolder)

	# Extract Images from docx files
	extract_image(ABS_PATH)

	# Rename files in "input" folder to its original name
	img2doc()
	
if __name__ == '__main__':
	main()