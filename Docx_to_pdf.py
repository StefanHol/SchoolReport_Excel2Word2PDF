# coding: utf-8

# # Konvertiert Word-Dokumente in PDF um
# - Benötigt wird eine Word installation -> <b>'Word.Application'</b>

import win32com
import win32com.client
import os
import logging


class Docx_to_PDF():
    def __init__(self):
        self.wdFormatPDF = 17
        self.logger = logging.getLogger("SR_E2W2P.Docx_to_PDF")
        self.logger.info("init Docx_to_PDF")
        pass

    def set_wdFormatPDF(self, wdFormatPDF=17):
        self.wdFormatPDF = wdFormatPDF

    def read_docx_files_in_folder(self, PDFFolder, Extention=".docx"):
        '''
        input: FolderName with pdf files
        return PyPDF2.pdf.PdfFileReader
        '''
        PDFFileNameArray = []
        for filename in os.listdir(PDFFolder):
            if filename.lower().endswith(Extention):
                PDFFileNameArray.append(filename)
            # print("filename: ", filename)
        # print(PDFFileNameArray)
        AllPDFNames = []
        for ActualPDFFileName in PDFFileNameArray:
            AllPDFNames.append([PDFFolder, ActualPDFFileName])
            # print(PDFFolder + os.path.sep + ActualPDFFileName)
        return AllPDFNames

    def covx_to_pdf(self, infile, outfile):
        """Convert a Word .docx to PDF"""
        try:
            word = win32com.client.Dispatch("Word.Application")
            doc = word.Documents.Open(infile)
            doc.SaveAs(outfile, FileFormat=self.wdFormatPDF)
        except Exception as e:
            self.logger.error("Error in Docx_to_PDF.covx_to_pdf. Input: {}\noutput: {}\nerror: {}".format(infile, outfile, e))
        finally:
            doc.Close()
            word.Quit()

    # Return the File Extention of a sting given FileName
    # e.g. Check for File Extention
    # # if FileExtention(String_to_check) == FileExtention(FileExtention_to_check_for):
    def FileExtention(self, FullString):
        FileParts = FullString.split('.')
        return FileParts[len(FileParts) - 1]

    def Compare_FileExtention(self, FullString, Extention):
        return self.FileExtention(FullString) == self.FileExtention(Extention)

    def convert_all_docx_of_inputFolder_to_PDF_in_outputFolder(self, infolder, outputFolder):
        '''
        Wandel alle Docx Dokumente aus dem "infolder" in PDF um und speichere unter "outputFolder" ab.
        '''
        All_docx_files_in_folder = self.read_docx_files_in_folder(infolder)
        for each_file in All_docx_files_in_folder:
            actual_source_file = each_file[0] + os.path.sep + each_file[1]
            actual_target_file = outputFolder + os.path.sep + each_file[1].replace(self.FileExtention(each_file[1]),
                                                                                   "pdf")
            self.logger.info("Converting: {} ".format(actual_source_file))
            self.covx_to_pdf(actual_source_file, actual_target_file)
            self.logger.info("Converted: {}".format(actual_target_file))
        self.logger.info("convert all docx to pdf done")


if __name__ == "__main__":
    infolder     = r""
    outputFolder = r""
    
    converter = Docx_to_PDF()
    converter.convert_all_docx_of_inputFolder_to_PDF_in_outputFolder(infolder, outputFolder)

