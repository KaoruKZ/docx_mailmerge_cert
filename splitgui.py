import tkinter as tk
from tkinter.messagebox import showinfo
import os, errno
import win32com.client as win32

cws = os.getcwd()

from PyPDF2 import PdfFileWriter, PdfFileReader, pdf
from datetime import MAXYEAR, date
from io import RawIOBase
from PIL import Image, ImageTk
from openpyxl import load_workbook, workbook

##############################################

wb = load_workbook("C:\\Users\\Ky\\Desktop\\Uber Certificate\\證書list.xlsx")
sheet1 = wb['Sheet1']
col = sheet1["F"]
namelist = []
for row in col:
    namelist.append(str(row.value))
del namelist[0]

##############################################

today = date.today()
today = str(today.strftime("%Y%m%d"))

#路徑設定
working_directory = os.getcwd() #檔案在哪就在哪做
source_name = '證書list.xlsx' #資料來源檔案名稱
destination_folder = "C:\\Users\\Ky\\Desktop\\Uber Certificate\\Cert_PDF" #輸出位置


##############################################
window = tk.Tk()
window.title('證書製作工具')
align_mode = 'nswe'
pad = 5

div_size = 360
img_size = 480

div1 = tk.Frame(window,  width=720 , height=1280 , bg='blue')
div2 = tk.Frame(window,  width=div_size , height=div_size , bg='orange')
div3 = tk.Frame(window,  width=div_size , height=div_size , bg='green')

window.update()
win_size = min( window.winfo_width(), window.winfo_height())

div1.grid(column=0, row=0, padx=pad, pady=pad, rowspan=2, sticky=align_mode)
div2.grid(column=1, row=0, padx=pad, pady=pad, sticky=align_mode)
div3.grid(column=1, row=1, padx=pad, pady=pad, sticky=align_mode)

###################################################
#視窗
def define_layout(obj, cols=1, rows=1):
    
    def method(trg, col, row):
        
        for c in range(cols):    
            trg.columnconfigure(c, weight=1)
        for r in range(rows):
            trg.rowconfigure(r, weight=1)

    if type(obj)==list:        
        [ method(trg, cols, rows) for trg in obj ]
    else:
        trg = obj
        method(trg, cols, rows)

define_layout(window, cols=2, rows=2)
define_layout([div1, div2, div3])

###################################################

def automated_mailmerge():
    wordApp = win32.Dispatch('Word.Application') #打開Word 程式
    wordApp.Visible = True #Word視窗為可見
    sourceDoc = wordApp.Documents.Open('C:\\Users\\Ky\\Desktop\\Uber Certificate\\證書範本.docx') #打開Word文件
    mail_merge = sourceDoc.MailMerge #使用Word合併列印功能
    mail_merge.OpenDataSource(
        Name:='C:\\Users\\Ky\\Desktop\\Uber Certificate\\證書list.xlsx',
        sqlstatement:="SELECT * FROM [Sheet1$]") #設定資料來源之路徑、工作表
    record_count = mail_merge.DataSource.RecordCount 
    mail_merge.Destination = 0 #合併列印重點
    mail_merge.Execute(False) #開始合併列印
    targetDoc = wordApp.ActiveDocument #選取合併列印召喚出來的視窗
    #targetDoc.SaveAs2(os.path.join(destination_folder, today + '.docx'), 16)  #存成 .docx檔
    targetDoc.ExportAsFixedFormat(os.path.join(destination_folder, today), exportformat:= 17) #存成 PDF檔
    targetDoc.Close(False) #關閉新文件視窗
    targetDoc = None #清空
    sourceDoc.MailMerge.MainDocumentType = -1 
    sourceDoc.Close(False) #關閉Word文件，False為不存檔
    wordApp.Quit #關閉Word


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
    #txt = open("name.txt", encoding="UTF-8")
    #names = txt.readlines()

    #   建立文件夾，若已存在則不建立
    try:
        os.makedirs("C://Users/Ky/Desktop/Uber Certificate/Cert_PDF/" + today)
    except OSError as e:
        if e.errno != errno.EEXIST:
            raise
    
    #   PDF檔中第【i】個頁數分割出來，儲存為txt檔中第【i】行的名字
    #   txt檔第一行是【王大明】，PDF第一頁分割出來會存為【王大明.pdf】
    for i in range(inputpdf.numPages):
        output = PdfFileWriter()
        output.addPage(inputpdf.getPage(i))
        strname = str(namelist[i]).strip("\n")
        pdfname = strname + ".pdf"
        pdfoutput = os.path.join(destination_folder, today, pdfname)
        print(pdfoutput)
        with open(pdfoutput, "wb") as outputStream:
            output.write(outputStream)
    #結束關檔
    #txt.close()
    showinfo("做得好", "Social Credit +15分")
    

def merge_and_split():
    automated_mailmerge()
    bytodaydate()

###################################################

im = Image.open('C:\\Users\\Ky\\Desktop\\Code\\SocialCredit.jpg')
imTK = ImageTk.PhotoImage( im.resize( (1280, 720) ) )

image_main = tk.Label(div1, image=imTK)
image_main['height'] = 720
image_main['width'] = 1280

image_main.grid(column=0, row=0, sticky=align_mode)

lbl_title1 = tk.Label(div2, text='Uber費用熊便宜\n計程司機怨無窮\n鬧到政府禁Uber\nUber司機沒錢賺\n計車學院來合作\n訓練課程獲證書\n車行司機為拿證\n電話訊息叫不停\n專任助理頭疼疼', bg='orange', fg='white')
lbl_title1.grid(column=0, row=0, sticky=align_mode)

bt1 = tk.Button(div3, text='工作100年', bg='green', fg='white')
#bt2 = tk.Button(div3, text='Button 2', bg='green', fg='white')
#bt3 = tk.Button(div3, text='Button 3', bg='green', fg='white')
#bt4 = tk.Button(div3, text='Button 4', bg='green', fg='white')

bt1.grid(column=0, row=0, sticky=align_mode)
#bt2.grid(column=0, row=1, sticky=align_mode)
#bt3.grid(column=0, row=2, sticky=align_mode)
#bt4.grid(column=0, row=3, sticky=align_mode)

bt1['command'] = lambda : merge_and_split()

define_layout(window, cols=2, rows=2)
define_layout(div1)
define_layout(div2, rows=2)
define_layout(div3)

window.mainloop()