import os
import re
from tkinter import *
import comtypes.client
import PyPDF2

def init_powerpoint():
   powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
   powerpoint.Visible = 1
   return powerpoint

def ppt_to_pdf(powerpoint, inputFileName, outputFileName, formatType = 32):
   if outputFileName[-3:] != 'pdf':
       outputFileName = outputFileName + ".pdf"
   deck = powerpoint.Presentations.Open(inputFileName)
   deck.SaveAs(outputFileName, formatType) # formatType = 32 for ppt to pdf
   deck.Close()

def convert_files_in_folder(powerpoint, folder):
   files = os.listdir(folder)
   pptfiles = [f for f in files if f.endswith((".ppt", ".pptx"))]
   for pptfile in pptfiles:
       fullpath = os.path.join(folder, pptfile)
       ppt_to_pdf(powerpoint, fullpath, fullpath)

def main():
   powerpoint = init_powerpoint()
   cwd = os.getcwd()
   convert_files_in_folder(powerpoint, cwd)
   powerpoint.Quit()
   main1()


def main1():
    # find all the pdf files in current directory.
    mypath = os.getcwd()
    pattern = r"\.pdf$"
    file_names_lst = [mypath + "\\" + f for f in os.listdir(mypath) if re.search(pattern, f, re.IGNORECASE) 
    and not re.search(r'合并.pdf',f)]

    # merge the file.
    opened_file = [open(file_name,'rb') for file_name in file_names_lst]
    pdfFM = PyPDF2.PdfFileMerger()
    for file in opened_file:
        pdfFM.append(file)

    # output the file.
    with open(mypath + "\\合并.pdf", 'wb') as write_out_file:
        pdfFM.write(write_out_file)

    # close all the input files.
    for file in opened_file:
        file.close()
    aa=('\n-----------------完成---------------------------')
    txt.insert(END, aa)
#-------------------------------------------------------------------------
def xianshiwenjian():
    files = os.listdir(os.getcwd())
    pptfiles = [f for f in files if f.endswith((".ppt", ".pptx"))]
    txt.insert(END, pptfiles)
    return pptfiles

def aaa():
    print(11111111111)
#------------------------界面-----------------------------------------------
root= Tk()
root.title('ppt合并')
root.geometry('500x300') # 这里的乘号不是 * ，而是小写英文字母 x

lb1 = Label(root, text='''请将本程序复制到待合并ppt文件夹内\n确认下框文件后点击继续''',font=("黑体",13))
lb1.place(x=25,rely=0.01, width=450, height=100)

txt = Text(root)
txt.place(relx=0.025, rely=0.5, relwidth=0.95, relheight=0.5)
xianshiwenjian()
btn1 = Button(root, text='继续',font=("黑体",15),command=main)
btn1.place(relx=0.35, rely=0.3, relwidth=0.3, relheight=0.15)

root.mainloop()

#---------如有问题请写邮件                           kila111@126.com
#---------If you have any questions, please email   kila111@126.com