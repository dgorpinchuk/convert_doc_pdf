import os
import re
import wx
import comtypes.client

def sanitize_filename(filename):
    # Remove unsupported symbols
    filename = re.sub(r'[\/:*?"<>|!,()]', '', filename)

    # Replace spaces with underscores
    filename = filename.replace(" ", "_")

    # Replace '+' with underscores
    filename = filename.replace("+", "_")

    # Replace '-' with underscores
    filename = filename.replace("-", "_")

    return filename

def convert_docx_to_pdf(docx_path, pdf_path):
    word = comtypes.client.CreateObject('Word.Application')
    try:
        doc = word.Documents.Open(docx_path)
        doc.SaveAs(pdf_path, FileFormat=17)  # 17 is the pdf format
        doc.Close()
    except Exception as e:
        print(f"Failed to convert {docx_path} to PDF. Error: {e}")
    finally:
        word.Quit()

def select_files():
    app = wx.App(False)
    style = wx.FD_OPEN | wx.FD_FILE_MUST_EXIST | wx.FD_MULTIPLE
    dialog = wx.FileDialog(None, "Select DOCX files", wildcard="Word files (*.doc;*.docx)|*.doc;*.docx", style=style)

    if dialog.ShowModal() == wx.ID_OK:
        paths = dialog.GetPaths()
    else:
        paths = []
    dialog.Destroy()
    return paths

def main():
    files = select_files()
    for file in files:
        dirname, filename = os.path.split(file)
        name_without_ext, _ = os.path.splitext(filename)
        sanitized_name = sanitize_filename(name_without_ext)
        pdf_path = os.path.join(dirname, sanitized_name + '.pdf')
        convert_docx_to_pdf(file, pdf_path)
        print(f"Converted '{file}' to '{pdf_path}'.")

if __name__ == '__main__':
    main()
