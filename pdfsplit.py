import os, errno
import win32com.client as win32
cws = os.getcwd()
from PyPDF2 import PdfFileWriter, PdfFileReader, pdf
from datetime import MAXYEAR, date
from io import RawIOBase

today = date.today()
today = str(today.strftime("%Y%m%d"))

#路徑設定
working_directory = os.getcwd() #檔案在哪就在哪做
source_name = '證書list.xlsx' #資料來源檔案名稱
destination_folder = os.path.join(working_directory,"Cert_PDF") #輸出位置
#   建立文件夾，若已存在則不建立


print("\n\n\n\n\n")
print("************************")
print("PDF Splitter || PDF分割器")
print("************************")
print("\n\n\n")

def automated_mailmerge():
    wordApp = win32.Dispatch('Word.Application')
    wordApp.Visible = True
    sourceDoc = wordApp.Documents.Open(os.path.join(working_directory,'證書範本.docx'))
    mail_merge = sourceDoc.MailMerge
    mail_merge.OpenDataSource(
        Name:=os.path.join(working_directory, source_name),
        sqlstatement:="SELECT * FROM [Sheet1$]")
    record_count = mail_merge.DataSource.RecordCount
    mail_merge.Destination = 0
    mail_merge.Execute(False)
    targetDoc = wordApp.ActiveDocument
    #targetDoc.SaveAs2(os.path.join(destination_folder, today + '.docx'), 16)  #存成 .docx檔
    targetDoc.ExportAsFixedFormat(os.path.join(destination_folder, today), exportformat:= 17) #存成 PDF檔
    targetDoc.Close(False)
    targetDoc = None
    sourceDoc.MailMerge.MainDocumentType = -1
    print('\n合併列印->轉PDF完成\n')

def bytodaydate():
    filename = today
    pdfname = str(filename + ".pdf") 
    filepath = str(destination_folder + "/" + pdfname)
    splitpdf(filepath, today)
    


def byinputdate():
    
    filename =  str(input("FileName || 檔案名稱: "))
    pdfname = str(filename + ".pdf") 
    filepath = str(destination_folder+ "/" + pdfname)
    splitpdf(filepath, filename)


def splitpdf(filepath, filedate):
    #   PDF讀檔
    inputpdf = PdfFileReader(open(filepath, "rb"))
    #   名單匯入
    txt = open("name.txt", encoding="UTF-8")
    names = txt.readlines()

    #   建立文件夾，若已存在則不建立
    try:
        os.makedirs("C://Users/H/Desktop/Cert_PDF/" + today)
    except OSError as e:
        if e.errno != errno.EEXIST:
            raise
    
    #   PDF檔中第【i】個頁數分割出來，儲存為txt檔中第【i】行的名字
    #   txt檔第一行是【王大明】，PDF第一頁分割出來會存為【王大明.pdf】
    for i in range(inputpdf.numPages):
        output = PdfFileWriter()
        output.addPage(inputpdf.getPage(i))
        strname = str(names[i]).strip("\n")
        pdfname = strname + ".pdf"
        pdfoutput = os.path.join(destination_folder, today, pdfname)
        print(pdfoutput)
        with open(pdfoutput, "wb") as outputStream:
            output.write(outputStream)
    #結束關檔
    txt.close()
    print('\n分割完成\n\n')

while(1):
    ans = int(input("1. 合併列印->輸出PDF->切割\n2. 合併列印->輸出PDF\n3. 今日PDF切割\n4. 輸入日期PDF分割\n輸入： "))
    if(ans==1):
        automated_mailmerge()
        bytodaydate()
    elif(ans==2):
        automated_mailmerge()
    elif(ans == 3):
        bytodaydate()
    elif(ans == 4):
        byinputdate()