from PyPDF2 import PdfFileWriter, PdfFileReader
import pikepdf
import jsbeautifier
import zlib
import os
import re
import sys

s = None
remove = False

#Check existence and validate of file
def InputPDF():
	path = input("Nhap duong dan: ")
	path = os.path.abspath(path)
	if (os.path.isfile(path)):
		try:
			PdfFileReader(open(path, "rb"))
		except expression as identifier:
			print ("File pdf khong hop le")
			sys.exit()
		f = open(path,"rb")
		global s 
		s = f.read()
		f.close()
		# s = str(s,"utf-8")
		return path
	else:
		print ("Duong dan khong hop le")
		sys.exit()

#Split folder, fileName, extension of file
def SplitPath(path):
	head, tail = os.path.split(path)
	fileName, extension = os.path.splitext(tail)
	folder = os.path.dirname(path)
	return fileName, extension, folder

#Check encrypted file
def CheckEncryption(path):
	global s 
	encryption = re.search(b"\/Encrypt\s+\d+\s+\d+",s)
	if (encryption != None):
		pdf = pikepdf.Pdf.open(path)
		fileName, extension, folder = SplitPath(path)
		newPath = folder + "\\decrypt_" + fileName + extension
		pdf.save(newPath)
		f = open(newPath,"rb")
		s = None
		s = f.read()
		f.close()
		# s = str(s,"utf-8")
		print ("File bi encrypt, dang tien hanh decrypt.....................Done!")
		return newPath
	else:
		return path

#Display js in file
def ExtractJS(path):
	print ("JavaScript extract tu file")
	print ("")
	global s
	fileName, extension, folder = SplitPath(path)
	newPath = folder + "\\text_" + fileName + ".txt"
	f = open(newPath,"w",encoding="utf-8")
	js = re.search(b"\/JS\s+\(.*\;\)",s)
	if  js != None:
		f.write((str(js.group(),"utf-8")))
		print ("Extract thanh cong: " + newPath)
	else:
		stream = re.compile(b'.*?FlateDecode.*?stream(.*?)endstream', re.S)
		for s in stream.findall(s):
			s = s.strip(b'\r\n')
			try:
				decode = zlib.decompress(s)
				decode = str(decode,"utf-8")
				# print (decode)
				if re.search(r"(function)\s+[a-zA-Z0-9_]+(\s+)?\((.*)?\)",decode) != None or re.search(r"(eval)\(.*\)\;",decode) != None:
					f.write(jsbeautifier.beautify(decode))
					f.close()
					print ("Extract thanh cong: " + newPath)
			except:
				pass     

#Check suspicious signal in file
def CheckSuspicious():
	global s
	global remove
	suspicious = ""
	objStm = re.search(b"\/ObjStm",s)            
	js = re.search(b"\/JS",s)                
	javaScript = re.search(b"\/JavaScript",s)     
	aa = re.search(b"\/AA",s)     
	openAction = re.search(b"\/OpenAction",s)     
	acroForm = re.search(b"\/AcroForm",s)     
	jbig2Decode = re.search(b"\/JBIG2Decode",s)    
	richMedia = re.search(b"\/RichMedia",s)    
	launch = re.search(b"\/Launch",s)    
	embeddedFiles = re.search(b"\/EmbeddedFiles",s)    
	xfa = re.search(b"\/XFA",s)    
	colors = re.search(b"\/Colors",s)   
	print ("-"*39)
	print ("Cac dau hieu nguy hiem co trong file")
	if (objStm != None):
		suspicious += " | " + str(objStm.group(),"utf-8").replace("/","")
		remove = True
	if (js != None):
		suspicious += " | " + str(js.group(),"utf-8").replace("/","")
		remove = True
	if (javaScript != None):
		suspicious += " | " + str(javaScript.group(),"utf-8").replace("/","")
		remove = True
	if (aa != None):
		suspicious += " | " + str(aa.group(),"utf-8").replace("/","")
		remove = True
	if (openAction != None):
		suspicious += " | " + str(openAction.group(),"utf-8").replace("/","")
		remove = True
	if (acroForm != None):
		suspicious += " | " + str(acroForm.group(),"utf-8").replace("/","")
		remove = True
	if (jbig2Decode != None):
		suspicious += " | " + str(jbig2Decode.group(),"utf-8").replace("/","")
		remove = True
	if (richMedia != None):
		suspicious += " | " + str(richMedia.group(),"utf-8").replace("/","")
		remove = True
	if (launch != None):
		suspicious += " | " + str(launch.group(),"utf-8").replace("/","")
		remove = True
	if (embeddedFiles != None):
		suspicious += " | " + str(embeddedFiles.group(),"utf-8").replace("/","")
		remove = True
	if (xfa != None):
		suspicious += " | " + str(xfa.group(),"utf-8").replace("/","")
		remove = True
	if (colors != None):
		suspicious += " | " + str(colors.group(),"utf-8").replace("/","")
		remove = True
	print (suspicious)
	print ("-"*39)

#Clean js in file via copy text and image into another file
def RemoveJS(path):
	global remove
	if remove == True:
		print ("-"*39)
		selection = input("Ban co muon remove js [y/n]: ")
		if (selection == 'y'):
			try:
				pdf = pikepdf.Pdf.open(path)
				newPdf = pikepdf.Pdf.new()
				newPdf.pages.extend(pdf.pages)
				fileName, extension, folder = SplitPath(path)
				newPath = folder + "\\clean_" + fileName + extension
				newPdf.save(newPath)
				print ("Remove js thanh cong: " + newPath)
			except expression as identifier:
				print ("Loi remove js")
		print ("-"*39)

#Inject js into fresh file
def InjectJS(path):
	js = re.search(b"\/JS",s)                
	javaScript = re.search(b"\/JavaScript",s)   
	if (js == None or javaScript == None):
		output = PdfFileWriter()
		ipdf = PdfFileReader(open(path, 'rb'))
		jsPath = input("Nhap duong dan file javascript: ")
		if (os.path.isfile(jsPath)):
			f = open(jsPath,"r")
			javascript = f.read()
			f.close()
			# javascript = """app.alert({cMsg: 'Hello from PDF JavaScript', cTitle: 'Testing PDF JavaScript', nIcon: 3});"""
			for i in range(ipdf.getNumPages()):
				page = ipdf.getPage(i)
				output.addPage(page)
			fileName, extension, folder = SplitPath(path)
			newPath = folder + "\\inject_" + fileName + extension
			with open(newPath, 'wb') as f:
				output.addJS(javascript)
				output.write(f)
				print ("Inject js thanh cong: " + newPath)
		else:
			print ("File khong ton tai")
	else:
		print("Vui long nhap file pdf khong co javascript")

#Extract text from file
def ExtractText(path):
	f = open(path, 'rb')
	pdf = PdfFileReader(f)
	f.close()
	pages = pdf.getNumPages()
	fileName, extension, folder = SplitPath(path)
	newPath = folder + "\\text_" + fileName + ".txt"
	f = open(newPath,"w", "utf-8", encoding="utf-8")
	for i in range(0,pages):
		page = pdf.getPage(i)
		content = page.extractText()
		f.write(str(content.encode('utf-8'),"utf-8"))
	f.close()
	print ("Extract text thanh cong: " + newPath)

def main():
	print("1. Kiem tra file pdf co nguy hai khong")
	print("2. Them javascript vao file san co")
	print("3. Extract text")
	selection = input("Moi ban lua chon: ")
	if (selection == "1"):
		path = InputPDF()
		path = CheckEncryption(path)
		CheckSuspicious()
		ExtractJS(path)
		RemoveJS(path)
	elif (selection == "2"):
		path = InputPDF()
		path = CheckEncryption(path)
		InjectJS(path)
	elif (selection == "3"):
		path = InputPDF()
		path = CheckEncryption(path)

if __name__ == "__main__":
    main()