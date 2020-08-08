import comtypes.client
import os
import PyPDF2
import re



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
       fullpath = os.path.join(cwd, pptfile)
       ppt_to_pdf(powerpoint, fullpath, fullpath)

if __name__ == "__main__":
   powerpoint = init_powerpoint()
   cwd = os.getcwd()
   convert_files_in_folder(powerpoint, cwd)
   powerpoint.Quit()


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


main1()

#---------如有问题请写邮件                           kila111@126.com
#---------If you have any questions, please email   kila111@126.com