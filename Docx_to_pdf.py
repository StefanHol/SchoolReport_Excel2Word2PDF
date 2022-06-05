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

    def check_file_exist(self, filepath):
        return os.path.exists(filepath)

    def docx_to_pdf(self, infile, outfile) -> bool:
        """Convert a Word .docx to PDF"""
        # Todo: build in retry
        # converting multiple documents failed on slow computers

        # retry = False # retry if write document failed
        # retry_count = 2
        # retry_counter = 0
        success = False
        if self.check_file_exist(outfile):
            os.remove(outfile)
            self.logger.info("Delete File first: {}".format(outfile))
        try:
            self.logger.debug("Word is running before starting: {}".format(self.check_word_running()))
            word = win32com.client.Dispatch("Word.Application")
            self.logger.debug("Started wort application: {}".format(word))
            doc = word.Documents.Open(infile)
            self.logger.debug("Open Document: {}".format(doc))
            doc.SaveAs(outfile, FileFormat=self.wdFormatPDF)
            self.logger.debug("Save Document as PDF: {}".format(doc))
            if self.check_file_exist(outfile):
                self.logger.info("File written: {}".format(outfile))
                success = True
            else:
                self.logger.info("Failed to write file: {}".format(outfile))
        except Exception as e:
            self.logger.error("Error in Docx_to_PDF.docx_to_pdf. Input: {}\noutput: {}\nerror: {}".format(infile, outfile, e))
            success = False
        finally:
            try:
                doc.Close()
            except Exception as e:
                self.logger.error("Error in Docx_to_PDF.docx_to_pdf. doc.Close() not possible")
                success = False
            try:
                word.Quit()
            except Exception as e:
                self.logger.error("Error in Docx_to_PDF.docx_to_pdf. word.Quit() not possible")
                success = False
                self.logger.error("Check Word is running: {}".format(self.check_word_running()))

        if self.check_file_exist(outfile):
            # self.logger.info("".format())
            pass
        else:
            self.logger.info("Failed to write file: {}".format(outfile))
            success = False
        self.logger.debug("Word is running at end: {}".format(self.check_word_running()))
        return success

    def FileExtention(self, FullString) -> str:
        '''
            Return the File Extention of a sting given FileName
            e.g. Check for File Extention
            # # if FileExtention(String_to_check) == FileExtention(FileExtention_to_check_for):
        '''
        FileParts = FullString.split('.')
        return FileParts[len(FileParts) - 1]

    def Compare_FileExtention(self, FullString, Extention):
        return self.FileExtention(FullString) == self.FileExtention(Extention)

    def convert_all_docx_of_inputFolder_to_PDF_in_outputFolder(self, infolder, outputFolder):
        '''
        Wandel alle Docx Dokumente aus dem "infolder" in PDF um und speichere unter "outputFolder" ab.
        '''
        success = None
        All_docx_files_in_folder = self.read_docx_files_in_folder(infolder)
        for each_file in All_docx_files_in_folder:
            actual_source_file = each_file[0] + os.path.sep + each_file[1]
            actual_target_file = outputFolder + os.path.sep + each_file[1].replace(self.FileExtention(each_file[1]),
                                                                                   "pdf")
            self.logger.info("Converting: {} ".format(actual_source_file))
            success = self.docx_to_pdf(actual_source_file, actual_target_file)
            self.logger.info("Converted: {}".format(actual_target_file))
        if success:
            self.logger.info("convert all docx to pdf done")
        else:
            self.logger.warning("convert all docx to pdf failed")

    def convert_selected_docx_file_to_PDF_in_outputFolder(self, wordfile, outputFolder):
        '''
        Wandel alle Docx Dokumente aus dem "infolder" in PDF um und speichere unter "outputFolder" ab.
        '''
        success = None

        actual_source_file = wordfile
        filename = wordfile.split(os.path.sep)[-1]

        actual_target_file = outputFolder + os.path.sep + filename.replace(self.FileExtention(filename),
                                                                               "pdf")
        self.logger.info("Converting: {} ".format(actual_source_file))
        success = self.docx_to_pdf(actual_source_file, actual_target_file)
        self.logger.info("Converted: {}".format(actual_target_file))

        if success:
            self.logger.info("convert selected docx to pdf done")
        else:
            self.logger.warning("convert selected docx to pdf failed")

    def check_word_running(self):
        from subprocess import Popen, PIPE
        command = "tasklist"
        pipe = Popen(command, shell=True, stdout=PIPE, stderr=PIPE)
        pipe.communicate()
        reply = str(pipe.communicate())
        process_to_find = "winword.exe"
        if process_to_find in reply.lower():
            self.logger.info("{} is running".format(process_to_find))
            return True
        else:
            self.logger.info("{} is not running".format(process_to_find))
            return False


if __name__ == "__main__":
    infolder     = r""
    outputFolder = r""
    
    converter = Docx_to_PDF()
    converter.convert_all_docx_of_inputFolder_to_PDF_in_outputFolder(infolder, outputFolder)

