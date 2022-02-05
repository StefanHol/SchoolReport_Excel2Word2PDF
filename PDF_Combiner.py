# coding: utf-8

import PyPDF2
import os
import logging


class Combine_PDF():
    def __init__(self):
        self.logger = logging.getLogger("SR_E2W2P.Combine_PDF")
        self.logger.info("init Combine_PDF")
        pass

    def append_PDFReaderObj_to_PDFWriterObj(self, pdfFileObj, pdfWriter):
        '''
        füge ein PDF Object zu einem PDF object hinzu
        input: 
            pdfFileObj    as PyPDF2.PdfFileReader(pdfFile_1)
            pdfWriter     as PyPDF2.PdfFileWriter()
        return:
            PdfFileWriter as PyPDF2.PdfFileWriter()
        '''
        for currentPageNumber in range(pdfFileObj.numPages):
            pageObject = pdfFileObj.getPage(currentPageNumber)
            pdfWriter.addPage(pageObject)
        return pdfWriter
        
    def read_pdf_files_in_folder(self, PDFFolder, Extention = ".pdf"):
        '''
        Funktion finded alle PDFs im Ordner "PDFFolder"
        input: FolderName with pdf files
        return PyPDF2.pdf.PdfFileReader
        '''
        PDFFileNameArray=[]
        for filename in os.listdir(PDFFolder):
            if filename.lower().endswith(Extention):
                PDFFileNameArray.append(filename)
            # print("filename: ", filename)
        # print(PDFFileNameArray)
        AllPDFNames=[]
        for ActualPDFFileName in PDFFileNameArray:
            AllPDFNames.append(PDFFolder + os.path.sep + ActualPDFFileName)
            # print(PDFFolder + os.path.sep + ActualPDFFileName)
        return AllPDFNames
        
    def PDF_Combiner(self, PDFFolder, PDFOutputName):
        '''
        Funktion um alle PDFs zusammenzufügen
        
        
        '''
        # Suche alle PDFs im Ordner "PDFFolder" 
        AllPDFNames = self.read_pdf_files_in_folder(PDFFolder)

        # erstelle ein PDFObject um alle seiten zu sammeln
        pdfWriter = PyPDF2.PdfFileWriter()
        pdfFile = []

        # Füge alles zusammen
        self.logger.info("Füge alle PDFs aus dem Ordner '{}' zusammen in '{}'".format(PDFFolder, PDFOutputName))
        for number, eachPDFName in enumerate(AllPDFNames):
            self.logger.info("Datei '{}' wird hinzugefügt".format(eachPDFName))
            pdfFile.append(open(eachPDFName, 'rb'))
            eachPDFObject = PyPDF2.PdfFileReader(pdfFile[number])
            pdfWriter = self.append_PDFReaderObj_to_PDFWriterObj(eachPDFObject, pdfWriter)

        # write Output and close
        pdfOutputFile = open(PDFOutputName, 'wb')
        self.logger.info("Schreibe: {}".format(PDFOutputName))
        pdfWriter.write(pdfOutputFile)
        pdfOutputFile.close()

        # schließe alle gelesenen dokumente
        for number, eachpdfFile in enumerate(pdfFile):
            pdfFile[number].close()


if __name__ == "__main__":
    PDFFolder     = r"C:\Temp\PDF"
    PDFOutputName = r"C:\Temp\PDF\Output.pdf"
    
    PDF_combiner = Combine_PDF()
    PDF_combiner.PDF_Combiner(PDFFolder, PDFOutputName)

