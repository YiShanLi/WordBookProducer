# -*- coding: utf-8 -*- 
from selenium import webdriver
from time import sleep
import sys, os, time
import win32com.client
import glob
from PyPDF2 import PdfFileMerger

numOfWordPerFile = 24
saveDir = os.path.join(os.getcwd(),"WordFile")

'''Delete all pdf files'''
filelist = glob.glob("WordFile/*.pdf")
for f in filelist:
    os.remove(f)
	
'''Text production'''

#Open Firefox browser
fp = webdriver.FirefoxProfile()
fp.set_preference("browser.download.folderList",2)
fp.set_preference("browser.download.manager.showWhenStarting",False)
fp.set_preference("browser.download.dir", saveDir)
fp.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/msword")
fp.set_preference("browser.download.animateNotifications", False)

browser = webdriver.Firefox(firefox_profile=fp)
browser.get("http://huayutools.mtc.ntnu.edu.tw/ebook/ebook1.aspx")

#Read words want to make the words_book from txt file
f = open('text.txt', 'r')
str = f.read()
#Delete newlines, tabs, spaces
str = str.replace('\n', '')
str = str.replace('\t', '')
str = str.replace(' ', '')
f.close()

#Fill the textbox on the website
str = unicode(str, "big5")
i = 0
while i < len(str):
	if (len(str)-i%numOfWordPerFile)>0 :
		text = str[i:i+numOfWordPerFile]
	else:
		text = str[i:len(str)]	
	browser.find_element_by_name("ctl00$ContentPlaceHolder1$TextBox2").clear()
	browser.find_element_by_name("ctl00$ContentPlaceHolder1$TextBox2").send_keys(text)
	browser.find_element_by_name("ctl00$ContentPlaceHolder1$Button3").click()
	i=i+numOfWordPerFile

time.sleep(5)
browser.quit()

'''Convert all docs into one pdf'''
i = 0
wdFormatPDF = 17
wdOpenFormatXML = 8
file_lst = os.listdir(saveDir)
word = win32com.client.Dispatch('Word.Application')
merger = PdfFileMerger()

for file in file_lst:
	if file.endswith(".doc"):
		pdf_name = os.path.splitext(file_lst[i])[0]+".pdf"
		in_file = os.path.abspath(os.path.join(saveDir, file_lst[i]))
		out_file = os.path.abspath(os.path.join(saveDir, pdf_name))
		doc = word.Documents.Open(in_file, Format = wdOpenFormatXML)
		doc.SaveAs(out_file, FileFormat=wdFormatPDF)
		doc.Close()
		input = open(out_file, "rb")
		merger.append(input)
		i+=1
	else:
		continue
		
time.sleep(5)
word.Quit()

# Writing all the collected pdf pages to a file
output = open("document-output.pdf", "wb")
merger.write(output)

'''Delete all download doc files'''
filelist = glob.glob("WordFile/*.doc")
for f in filelist:
    os.remove(f)

# Creating a routine that appends files to the output file
def append_pdf(input,output):
    [output.addPage(input.getPage(page_num)) for page_num in range(input.numPages)]









