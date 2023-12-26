import os
import subprocess
import win32com.client
import sys

printer_name = 'Canon G2020 series'
sumatraEXE_file_path = 'C:\\Users\\USER\\AppData\\Local\\SumatraPDF\\SumatraPDF.exe';


def print_pdf(file_path, printer_name):
    # Use subprocess to open SumatraPDF and print the PDF file
    subprocess.run([sumatraEXE_file_path , '-print-to', printer_name, file_path], shell=True)

def print_docx(file_path, printer_name):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    doc = word.Documents.Open(file_path)
    word.Application.ActivePrinter = printer_name
    doc.PrintOut()
    doc.Close()
    word.Quit()

    
if len(sys.argv) != 2:
    print("Usage: python script.py <file_path>")
    sys.exit(1)

file_to_print = sys.argv[1]

file_extension = file_to_print.split('.')[-1].lower()

if file_extension == 'pdf' or file_extension == 'txt':
    print_pdf(file_to_print, printer_name)
elif file_extension == 'docx' or file_extension =='doc':
    print_docx(file_to_print, printer_name)
else:
    print(f"Unsupported file type: {file_extension}")