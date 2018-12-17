import oletools.oleid
import msoffcrypto
import os
import re
import sys
import io
from oletools.olevba import VBA_Parser
from win32com.client import DispatchEx, GetObject
# import pyexcel as Workbook

#Suspicious keyword macro auto execute
AUTOEXEC_KEYWORDS = {
    'Runs when the Word document is opened':
        ('AutoExec', 'AutoOpen', 'DocumentOpen'),
    'Runs when the Word document is closed':
        ('AutoExit', 'AutoClose', 'Document_Close', 'DocumentBeforeClose'),
    'Runs when the Word document is modified':
        ('DocumentChange',),
    'Runs when a new Word document is created':
        ('AutoNew', 'Document_New', 'NewDocument'),

    'Runs when the Word or Publisher document is opened':
        ('Document_Open',),
    'Runs when the Publisher document is closed':
        ('Document_BeforeClose',),

    'Runs when the Excel Workbook is opened':
        ('Auto_Open', 'Workbook_Open', 'Workbook_Activate'),
    'Runs when the Excel Workbook is closed':
        ('Auto_Close', 'Workbook_Close'),
}

#Check encrypted file
def CheckEncryption(path):
    oid = oletools.oleid.OleID(path)
    indicators = oid.check()
    for i in indicators:
        if (i.name == "Encrypted" and str(i.value) == "True"):
            return True
    return False

#Split path, fileName, extension, folder
def SplitPath(path):
    fileFullname = os.path.basename(path)
    filename = fileFullname.split('.')[0]
    ext = fileFullname.split('.')[1]
    dirpath = os.path.dirname(path)
    return [dirpath, filename, ext]

#Decrypt file
def Decrypt(path, pwd, dirpath, filename, ext):
    file = msoffcrypto.OfficeFile(open(path, "rb"))
    file.load_key(password=pwd)
    SplitPath(path)
    fileSave = dirpath + '\\' + 'decrypted_' + filename + '.' + ext
    file.decrypt(open(fileSave, "wb"))
    return fileSave

#Detect suspicious siganl in file
def Detect(path, dirpath, filename, ext):
    print ("-"*39)
    flag = False
    vbaparser = VBA_Parser(path)
    if vbaparser.detect_vba_macros():
        flag = True
        print 'VBA Macros found'
        selection = raw_input("Ban co muon dump vba [y/n]: ")
        if (selection == "y"):
            fileSave = dirpath + '\\' + 'vba_' + filename + '.txt'
            f = open(fileSave,"w")
            for (filename, stream_path, vba_filename, vba_code) in vbaparser.extract_macros():
                f.write(str(vba_code).encode("utf-8"))
            f.close()
    else:
        flag = False
        print 'No VBA Macros found'
    # for (filename, stream_path, vba_filename, vba_code) in vbaparser.extract_macros():
    #     pass
    print '-'*39
    vbaparser.analyze_macros()
    print 'Suspicious keywords: %d' % vbaparser.nb_suspicious
    return flag

#Disable macro in file
def Clean(path, dirpath, filename, ext):
    fileSave = dirpath + '\\' + 'clean_' + filename + '.' + ext
    # Workbook.save_book_as(file_name=path, dest_file_name=fileSave)
    f = open(path,'rb')
    s = f.read()
    f.close()
    disableMacro = b''
    for description, keywords in AUTOEXEC_KEYWORDS.items():
        for keyword in keywords:
            s = re.sub(r'(?i)\b' + keyword + r'\b', disableMacro, s)
    f = open(fileSave,"wb")
    f.write(s)
    f.close()
    print "Disable thanh cong: " + fileSave

#Check file is OLE file
def CheckValidation(path, ext):
    word = ['doc']
    excel = ['xls']   

    oid = oletools.oleid.OleID(path)
    indicators = oid.check()
    for i in indicators:
        if (i.name == "OLE format" and str(i.value) == "False"):
            return False
    return True

#Inject macro into fresh file
def Inject(path, dirpath, filename, ext):
    vbaparser = VBA_Parser(path)
    print '-'*39
    if vbaparser.detect_vba_macros() == False:
        pathMacro = raw_input("Nhap duong dan file macro: ")
        f = open(pathMacro,'r')
        fileMacro = f.read()
        f.close()

        word = ['doc']
        excel = ['xls']  
        fileSave = dirpath + '\\' + 'injected_' + filename + '.' + ext

        if ext in word:
            doc = DispatchEx("Word.Application")
            document = doc.Documents.Add(path)
            mod = document.VBProject.VBComponents("ThisDocument")
            mod.CodeModule.AddFromString(fileMacro)
            document.SaveAs(fileSave, FileFormat=0)
            doc.Quit()
        elif ext in excel:
            xl = DispatchEx("Excel.Application")
            wb = xl.Workbooks.Add(path)
            mod = wb.VBProject.VBComponents("ThisWorkbook")
            mod.CodeModule.AddFromString(fileMacro)
            wb.SaveAs(fileSave, FileFormat=56)
            xl.Quit()
        else:
            print "File khong hop le de inject macro"
        print "Inject thanh cong: " + fileSave
    else:
        print "Vui long nhap file sach"

#Extract text from ole
def ExtractText(path, dirpath, filename, ext):
    filePath = path
    doc = GetObject(filePath)
    text = doc.Content.Text
    fileSave = dirpath + '\\' + 'text_' + filename + '.txt'
    f = io.open(fileSave,"w",encoding="utf-8")
    f.write(text)
    f.close()
    print "-"*39
    print "Extract thanh cong: " + fileSave

#Ham main
def main():
    selection = 0
    path = raw_input("Nhap duong dan: ")
    if (os.path.isfile(path)):
        split = SplitPath(path)
        dirpath = split[0]
        filename = split[1]
        ext = split[2]
        validation = CheckValidation(path, ext)

        if (validation):
            encrypted = CheckEncryption(path)
            if (encrypted):
                pwd = raw_input("Nhap password: ")
                path = Decrypt(path, pwd, dirpath, filename, ext)
        else:
            print 'Khong phai file OLE'
            sys.exit(0)

        while selection == 0:
            print "-"*60
            print "1. Disable macro"
            print "2. Inject macro"
            print "3. Extract text"
            selection = raw_input("Vui long lua chon: ")
        
        if (selection == "1"):
            flag = Detect(path, dirpath, filename, ext)
            if (flag):
                disable = raw_input("File co dau hieu nguy hiem cua macro, ban co muon disable macro [y/n]: ")
                if (disable == 'y'):
                    Clean(path, dirpath, filename, ext)
        elif (selection == "2"):
            Inject(path, dirpath, filename, ext)
        elif (selection == "3"):
            ExtractText(path, dirpath, filename, ext)

if __name__ == "__main__":
    main()