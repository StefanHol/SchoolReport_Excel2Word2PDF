# python3
# coding: utf-8

# import threading
import openpyxl as px
import pandas as pd

from helper_Word import helper_Word
from Docx_to_pdf import Docx_to_PDF
from PDF_Combiner import Combine_PDF
from shutil import copyfile
import os
# from collections import OrderedDict
import time
import json_helper
# import Excel_to_Word_in_arbeit
# from helper_Word import helper_Word as hw

__version__ = "0.0.16"
__author__ = "Stefan Holstein"
__repo__ = "http://git......"
__config_json__ = "config.json"
__Anleitung__ = "SchoolReport Excel2Word2PDF.pdf"
# rebuild_GUI = True
rebuild_GUI = False


# # ToDo: Add certificate
# # Could not find a suitable TLS CA certificate bundle, invalid path
# # https://github.com/pyinstaller/pyinstaller/issues/6352

used_Qt_Version = 0
try:
    from PyQt5.QtWidgets import QMainWindow, QInputDialog, QMessageBox
    from PyQt5.QtWidgets import QWidget, QComboBox, QCheckBox, QLineEdit, QFileDialog, QTableWidgetItem
    from PyQt5.QtGui import QFont, QIcon, QPixmap
    from PyQt5.QtWidgets import QTableView, QAbstractItemView, QCompleter, QLineEdit, QHeaderView

    from PyQt5.QtCore import QAbstractTableModel, Qt

    used_Qt_Version = 5
except Exception as e:
    print(e)
    exit()
    pass


def compile_GUI():
    if used_Qt_Version == 4:
        print("Compile QUI for Qt Version: " + str(used_Qt_Version))
        os.system("pyuic4 -o GUI\Converter_ui.py GUI\Converter.ui")
    elif used_Qt_Version == 5:
        print("Compile QUI for Qt Version: " + str(used_Qt_Version))
        os.system("pyuic5 -o GUI\Converter_ui.py GUI\Converter.ui")


if rebuild_GUI:
    compile_GUI()

from GUI.Converter_ui import Ui_MainWindow


import logging

logging.basicConfig(level=logging.WARNING)


class PandasModel(QAbstractTableModel):
    """
    Class to populate a table view with a pandas dataframe
    https://stackoverflow.com/questions/31475965/fastest-way-to-populate-qtableview-from-pandas-data-frame

    """
    def __init__(self, data, parent=None):
        QAbstractTableModel.__init__(self, parent)
        self._data = data
        logger = logging.getLogger(__name__)
        logger.debug("init Pandas Model")

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                return str(self._data.iloc[index.row(), index.column()])
        return None

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._data.columns[col]
        return None


class create_letter_of_reference():
    def __init__(self):
        self.logger = logging.getLogger("SR_E2W2P.create_letter_of_reference")
        self.logger.info("init create_letter_of_reference")

        self.Array_CheckBox_Namen = None
        self.Text_Feld_Namen = None
        self.wb = None
        self.df = None

        self.srcFileName = ""
        self.destination_Folder = ""
        self.Excel_Filename = ""

        self.wichtige_meldung = ""

    def read_word_information_fields(self, srcFileName):
        self.srcFileName = srcFileName
        self.logger.info("read_word_information_fields")
        word_doc = helper_Word()
        word_doc.init_file(srcFileName, False)
        self.logger.info("get_all_fieldnames_from_word")
        word_doc.get_all_fieldnames_from_word()
        self.logger.info("find_arrays_and_text_field_names")
        word_doc.find_arrays_and_text_field_names()
        self.logger.info("find_arrays_and_text_field_names: done")
        self.Array_CheckBox_Namen = word_doc.Text_FieldArray_Row_Names
        self.Text_Feld_Namen = word_doc.Text_Field_Names
        self.logger.info("Array_CheckBox_Namen:")
        self.logger.info(self.Array_CheckBox_Namen)
        self.logger.info("Text_Feld_Namen:")
        self.logger.info(self.Text_Feld_Namen)
        try:
            word_doc.doc.Close()
            self.logger.info("Word Document Closed")
        except Exception as e:
            self.logger.error(e)
        return self.Array_CheckBox_Namen, self.Text_Feld_Namen

    def read_excel_file_to_pandas(self, Excel_Filename):
        self.logger.debug("read_excel_file_to_pandas")
        # print("Excel_Filename")
        self.Excel_Filename = Excel_Filename
        # print("self.Excel_Filename: ", self.Excel_Filename)
        self.wb = px.load_workbook(self.Excel_Filename, data_only=True)
        self.logger.info("loaded Excel File: {}".format(Excel_Filename))
        self.logger.info("Array_CheckBox_Namen\n{}".format(self.Array_CheckBox_Namen))
        self.df = self.create_dataframe_for_InputFields(self.Array_CheckBox_Namen, self.Text_Feld_Namen)
        # print("self.df", self.df)
        self.df = self.remove_empty_data(self.df)
        self.logger.debug("self.df {}".format(self.df))
        return self.df

    def get_values_By_range_name(self, range_name):
        value_array = []
        self.logger.debug("check: {}".format(range_name))
        try:
            address = list(self.wb.defined_names[range_name].destinations)
            #     address
            for sheetname, cellAddress in address:
                cellAddress = cellAddress.replace('$', '')
            cellAddress
            worksheet = self.wb[sheetname]
            for i in range(0, len(worksheet[cellAddress])):
                for item in worksheet[cellAddress][i]:
                    #         print(item.value)
                    value_array.append(item.value)
        except Exception as e:
            self.logger.error("Error: Textfeld {} ist in Worddokument definiert,".format(range_name) +
                  "kann aber nicht im Exceldokument gefunden werden.\n{}".format(e))

            self.wichtige_meldung = ("Warnung: Textfeld {} ist in Worddokument definiert,".format(range_name) +
                                    "kann aber nicht im Exceldokument gefunden werden.\n{}\n".format(e) +
                                    "Vorlagen überprüfen!\nDie entsprechenden Daten werden nicht übertragen.\n")
        return value_array

    def create_folder_if_not_exist(self, folder):
        if not os.path.exists(folder):
            os.mkdir(folder)

    def ExtractFileName(self, mypath):
        name = mypath.split(os.path.sep)
        return name[len(name) - 1]

    def zahl_zu_bewertung(self, value):
        if str(value) == "1":
            return "sehr gut"
        if str(value) == "2":
            return "gut"
        if str(value) == "3":
            return "schlecht"
        if str(value) == "4":
            return "sehr schlecht"

    def Copy_word_file_and_write_data(self, df, targetname, df_index):
        # copy File and create target file object
        self.create_folder_if_not_exist(self.destination_Folder)
        FileName = self.srcFileName.replace(".docx", "_-_" + (df["Name"].iloc[df_index]).replace(" ", "_") + ".docx")
        FileName = self.ExtractFileName(
            self.srcFileName.replace(".docx", "_-_" + (df["Name"].iloc[df_index]).replace(" ", "_") + ".docx"))
        #    dstFileName = srcFileName.replace(".docx", "_-_" + (df["Name"].iloc[df_index]).replace(" ", "_") + ".docx")
        logging.info(FileName)
        dstFileName = self.destination_Folder + os.path.sep + FileName
        #     print(srcFileName,"\n"+ dstFileName)
        copyfile(self.srcFileName, dstFileName)
        # wordFile.doc.Close()
        wordFile = helper_Word()
        wordFile.init_file(dstFileName, False)
        self.logger.info(FileName)

        for ColumnName in df:
            # print(df["Name"].iloc[df_index], ":\t",ColumnName , "\t",  df[ColumnName].iloc[df_index])
            # ToDo: Call Input Function here
            SchülerName = df["Name"].iloc[df_index]
            ZellenNameOderArray = ColumnName
            Cellen_Oder_Array_Wert = df[ColumnName].iloc[df_index]
            self.logger.info("Name: {} \t {} \t {}".format(SchülerName, ColumnName, Cellen_Oder_Array_Wert))
            # ToDo: Call Input Function here

        try:
            counter = 0
            for TextElement in self.Text_Feld_Namen:
                # print(counter)
                # print(TextElement)
                counter += 1
                outputText = df[TextElement].iloc[df_index]
                try:
                    try:
                        self.logger.info("int: TextElement {} {}".format(TextElement, int(outputText)))
                        outputText = int(outputText)
                    except Exception as e:
                        self.logger.info("str: TextElement {} {}".format(TextElement, str(outputText)))
                        outputText = str(outputText)
                        # print(e, str(outputText), str(TextElement))
                    wordFile.set_WordField_Text(TextElement, str(outputText))
                    self.logger.info("TextElement {} {}".format(TextElement, wordFile.get_WordField_Text(TextElement)))
                except Exception as e:
                    self.logger.debug("{} {} {}".format(e, str(outputText), str(TextElement)))
                    # print(e, str(outputText), str(TextElement))
            counter = 0
            for ArrayElement in self.Array_CheckBox_Namen:
                #             print(counter)
                counter += 1
                outputText = df[ArrayElement].iloc[df_index]
                #########################################################################
                # Noten CheckBox und Text
                # korrigieren
                #########################################################################
                SetCheckBox = False
                realerror = True
                error_message = ""
                try:
                    outputText = int(outputText)
                    self.logger.debug("int: ArrayElement {} {}".format(outputText, ArrayElement))
                    self.logger.debug(str(ArrayElement) + " = " + self.zahl_zu_bewertung(outputText))
                    realerror = False
                except Exception as e:
                    try:
                        outputText = str(outputText)
                        self.logger.debug("str: ArrayElement {} {}".format(outputText, ArrayElement))
                        self.logger.debug(str(ArrayElement) + " = " + self.zahl_zu_bewertung(outputText))
                        realerror = False
                    except Exception as e:
                        realerror = True
                        self.logger.debug("{} try of try outputText: {} {}".format(e,  ArrayElement, outputText))
                        error_message = "{}: try of try outputText: {}, {}".format(e, ArrayElement, outputText)
                        #                     print(e, ArrayElement, outputText)
                        pass
                try:
                    if not (outputText is None):
                        CheckBox = ArrayElement + str(outputText - 1)
                        SetCheckBox = True
                        realerror = False
                    else:
                        self.logger.debug("Skip Checkbox: {} {}".format(outputText, ArrayElement))
                        pass
                except Exception as e:
                    realerror = True
                    self.logger.debug("tryexcept 2: {} {} {}".format(e, outputText, ArrayElement))
                    error_message = "warning: nichts in checkbox eingetragen {}, {}, {}".format(e, outputText,
                                                                                                ArrayElement)
                    SetCheckBox = False
                # print(CheckBox)
                if SetCheckBox:
                    realerror = False
                    wordFile.check_checkbox(CheckBox)
                    # save & close target file here

                if realerror:
                    self.logger.debug("Hinweis: {}".format(error_message))
        except Exception as e:
            self.logger.error("Fehler hier: {}".format(e))

        finally:
            try:
                wordFile.save_and_close()
            except Exception as e:
                print(e)

        return FileName

    def create_wordDocxs_from_dataframe(self, df):
        #     print(df.head())
        if True:
            start_time = time.time()
        file_name_list = []
        for index in range(len(df)):
            #         print(index)
            #    #ToDo: Call Input Function here
            self.logger.debug(str(df["Name"].iloc[index]))
            file_name = self.Copy_word_file_and_write_data(df, df["Name"].iloc[index], index)
            file_name_list.append(file_name)
        if True:
            self.logger.info("Creation takes %.2f seconds" % (time.time() - start_time))
        # print("create_wordDocxs_from_dataframe", file_name_list)
        return file_name_list

    def create_dataframe_for_InputFields(self, Array_CheckBox_Namen, Text_Feld_Namen):
        self.logger.debug("create_dataframe_for_InputFields")
        df = pd.DataFrame({})
        # df.to_excel("tempout.xlsx")
        # df.to_csv("tempout.csv", index=False, header=True)
        self.logger.info("Text_Feld_Namen\n{}".format(Text_Feld_Namen))
        for index, HeaderName in enumerate(Text_Feld_Namen):
            # print(index, HeaderName)
            if index == 0:
                # print(HeaderName)
                # print(HeaderName, self.get_values_By_range_name(HeaderName),
                #                   len(self.get_values_By_range_name(HeaderName)))
                try:
                    df = pd.DataFrame(self.get_values_By_range_name(HeaderName), columns=[HeaderName])
                except Exception as e:
                    self.logger.error("create_dataframe_for_InputFields_1 {}".format(e))
                # df = pd.DataFrame.from_dict([(HeaderName, self.get_values_By_range_name(HeaderName))])
                #         print(HeaderName, len(df))
            else:
                self.logger.debug("{}, {}, {}".format(HeaderName, str(self.get_values_By_range_name(HeaderName)),
                      len(self.get_values_By_range_name(HeaderName))))
                try:
                    se = pd.Series(self.get_values_By_range_name(HeaderName))
                    #         print(HeaderName ,len(df), len(se))
                    df[HeaderName] = se.values
                except Exception as e:
                    self.logger.error("create_dataframe_for_InputFields_2 {}".format(e))
        # print(df.head())
        #
        # df.to_excel("tempout.xlsx")

        for index, HeaderName in enumerate(Array_CheckBox_Namen):
            # if index ==0:
            #    print(HeaderName)
            #    df = pd.DataFrame.from_items([(HeaderName, get_values_By_range_name(HeaderName))])
            #    print(HeaderName, len(df))
            # else:
            try:
                # print(index, HeaderName)
                se = pd.Series(self.get_values_By_range_name(HeaderName))
                self.logger.debug("HeaderName: {} len(df): {} len(se): {}".format(HeaderName, len(df), len(se)))
                df[HeaderName] = se.values
            except Exception as e:
                self.logger.debug("checkbox not found: {}".format(e))
                self.logger.debug("{} {}".format(index, HeaderName))
        # print("end for loop Array_CheckBox_Namen")
        return df

    def remove_empty_data(self, df):
        Empty_List = df.loc[df['Name'].isin([None])].index.tolist()
        df.drop(df.index[Empty_List], inplace=True)
        df = df.reset_index(drop=True)
        return df


class MainProg(QMainWindow):
    """Scanstation GUI"""

    def __init__(self, parent=None):
        QWidget.__init__(self, parent)

        self.logger = logging.getLogger("SR_E2W2P.MainProg")
        self.logger.info("init MainProg")

        # load default json config if not existing
        self.check_or_create_default_json_config()
        self.config_information = json_helper.read_json_file(__config_json__)
        self.logger.debug(self.config_information)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.work_folder = self.config_information.get("dokumenten_start_ordner")

        self.converter = Docx_to_PDF()
        self.creator_class = create_letter_of_reference()
        self.output_pdf_folder = ""
        self.auto_pdf_folder = self.config_information.get("Autotomatisch_PDF_Ordner_anlegen")
        self.ui.actionAutomatisch_PDF_Ordner_anlegen.setChecked(self.auto_pdf_folder)
        self.ui.pushButton_PDF_Folder.setEnabled(not self.auto_pdf_folder)

        self.word_selected = False
        self.excel_selected = False
        self.output_selected = False
        self.pdf_output_selected = False

        self.ui.pushButton_Zeugnisvorlage.clicked.connect(self.select_new_Word_Template_File)
        self.ui.pushButton_Excel_Zeugnisse.clicked.connect(self.select_Excel_Zeugnis_File)
        self.ui.pushButton_Ausgabe_Ordner.clicked.connect(self.select_output_folder)
        self.ui.pushButton_OpenOutputFolder.clicked.connect(self.open_output_folder)
        self.ui.pushButton_OpenPDFOutputFolder.clicked.connect(self.open_PDF_output_folder)

        self.ui.pushButton_create_only_selected_docx.clicked.connect(self.create_single_selected_reference_word)
        self.ui.pushButton_create_only_selected_docx_pdf.clicked.connect(
                                self.create_single_selected_reference_word_and_pdf)
        self.ui.pushButton_create_all_docx.clicked.connect(self.create_all_references_word)
        self.ui.pushButton_create_all_docx_pdf.clicked.connect(self.create_all_references_word_and_pdf)

        self.ui.pushButton_Dummy.clicked.connect(self.dummy)
        self.ui.pushButton_combine_pdfs.clicked.connect(self.combine_all_PDFs_in_selected_output_folder)
        self.ui.pushButton_convert_all_docx_tp_pdf.clicked.connect(self.convert_all_docx_to_pdf)

        self.ui.pushButton_PDF_Folder.clicked.connect(self.select_PDF_folder)

        self.ui.actionSet_search_folder.triggered.connect(self.set_dokument_search_folder)
        self.ui.actionAutomatisch_PDF_Ordner_anlegen.triggered.connect(self.Toggle_actionAutomatisch_PDF_Ordner_anlegen)

        self.ui.actionAnleitung_oeffnen.triggered.connect(self.Anleitung_oeffnen)
        self.ui.actionkill_winword.triggered.connect(self.kill_winword)

        self.ui.actionAbout.triggered.connect(self.show_about)

        self.ui.pushButton_Dummy.setHidden(True)

        # add banner image
        banner_name = "img" + os.path.sep + "Logo.png"
        banner_path = os.path.dirname(__file__) + os.path.sep + banner_name
        self.pixmap = QPixmap(banner_path)
        # self.pixmap = self.pixmap.scaledToWidth(300)
        self.ui.banner.setPixmap(self.pixmap)
        self.ui.banner.setScaledContents(True)

        # Dummy Data
        data = ([["Name 1", "Demo Daten", "Demo Daten", "Demo Daten", "Demo Daten"],
                 ["Name 2", "Demo Daten", "Demo Daten", "Demo Daten", "Demo Daten"],
                 ["Name 3", "Demo Daten", "Demo Daten", "Demo Daten", "Demo Daten"]])
        your_pandas_data = pd.DataFrame(data)
        your_pandas_data.columns = ['Name', 'Spalte 1', 'Spalte 2', 'Spalte 3', 'Spalte 4']

        TableModel = PandasModel(your_pandas_data)
        self.ui.tableView.setModel(TableModel)

    def check_or_create_default_json_config(self):
        if not os.path.isfile(__config_json__):
            data = {
                "dokumenten_start_ordner": r"C:\work\Excel2Word2PDF_SchoolReport\Demo",
                "logging_level": "Debug",
                "Autotomatisch_PDF_Ordner_anlegen": True
            }
            json_helper.write_json_file(__config_json__, data)

    def write_json_config_information(self):
        json_helper.write_json_file(__config_json__, self.config_information)

    def select_output_folder(self):
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.FileMode())
        self.output_folder = dialog.getExistingDirectory(None, 'Wähle einen leeren Ausgabeordner', self.work_folder)
        temp = str(self.output_folder.replace('/', os.path.sep))
        selected = False
        if os.path.exists(temp):
            self.output_folder = temp
            self.logger.debug(self.output_folder)
            selected = True
        if selected:
            self.ui.lineEdit_ZeugnisAusgabe.setText(self.output_folder)
            self.creator_class.destination_Folder = self.output_folder
            self.output_selected = True
            self.enable_buttons()
            if self.auto_pdf_folder:
                self.set_PDF_Folder(self.output_folder + os.path.sep + "_PDF_")
            self.enable_buttons()

    def open_output_folder(self):
        self.logger.debug("open_output_folder")
        import subprocess
        # Open Script File Location
        self.logger.debug('open =', self.output_folder)
        subprocess.Popen(r'explorer ' + self.output_folder)

    def open_PDF_output_folder(self):
        self.logger.debug("open_PDF_output_folder")
        import subprocess
        # Open Script File Location
        self.logger.debug('open =', self.output_pdf_folder)
        subprocess.Popen(r'explorer ' + self.output_pdf_folder)

    def select_PDF_folder(self):
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.FileMode())
        folder = dialog.getExistingDirectory(None, 'Wähle den PDF Ausgabeordner', self.work_folder)
        folder = str(folder.replace('/', os.path.sep))
        if os.path.exists(folder):
            self.logger.info("PDF_folder", folder)
            selected = True
            self.set_PDF_Folder(folder)

    def set_PDF_Folder(self, folder):
        self.output_pdf_folder = folder
        self.creator_class.create_folder_if_not_exist(self.output_pdf_folder)
        self.ui.lineEdit_PDF_output.setText(self.output_pdf_folder)

        self.pdf_output_selected = True
        self.enable_buttons()

    def enable_buttons(self):
        if self.word_selected and not self.excel_selected:
            self.ui.pushButton_Excel_Zeugnisse.setEnabled(True)
        if self.excel_selected and self.word_selected:
            self.ui.pushButton_Ausgabe_Ordner.setEnabled(True)

        if self.excel_selected and self.word_selected and self.output_selected:
            self.ui.pushButton_create_only_selected_docx.setEnabled(True)
            self.ui.pushButton_create_all_docx.setEnabled(True)

        if self.excel_selected and self.word_selected and self.output_selected and self.pdf_output_selected:
            self.ui.pushButton_create_only_selected_docx.setEnabled(True)
            self.ui.pushButton_create_only_selected_docx_pdf.setEnabled(True)
            self.ui.pushButton_create_all_docx.setEnabled(True)
            self.ui.pushButton_create_all_docx_pdf.setEnabled(True)
            self.ui.pushButton_convert_all_docx_tp_pdf.setEnabled(True)
            self.ui.pushButton_combine_pdfs.setEnabled(True)

        if self.output_selected:
            self.ui.pushButton_OpenOutputFolder.setEnabled(True)

        if self.pdf_output_selected:
            self.ui.pushButton_OpenPDFOutputFolder.setEnabled(True)

    def set_dokument_search_folder(self):
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.FileMode())
        temp = dialog.getExistingDirectory(None, 'Wähle deinen haupt Dokumentenordner', self.work_folder)
        temp = str(temp.replace('/', os.path.sep))
        self.work_folder = temp
        self.config_information["dokumenten_start_ordner"] = self.work_folder
        self.logger.info(self.config_information)
        json_helper.write_json_file(__config_json__, self.config_information)

    def select_new_Word_Template_File(self):
        # ToDo: Inputfile
        # ToDo: filter richtig setzen nur *.docx anzeigen
        self.srcFileName = ""

        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.FileMode())
        selected = False
        try:
            filter = "Office File (*.docx)"
            word_template = dialog.getOpenFileName(self, 'Wähle die Zeugnis Vorlage *.docx', self.work_folder, filter,
                                                   "*.docx")
            word_template = str(word_template[0].replace('/', os.path.sep))
            self.srcFileName = word_template
            if os.path.exists(self.srcFileName):
                self.Array_CheckBox_Namen, self.Text_Feld_Namen = self.creator_class.read_word_information_fields(
                    self.srcFileName)
                selected = True
            else:
                self.logger.debug("kein Word Dokument gewählt")

        except Exception as e:
            print("select_new_Word_Template_File"  + str(e))
        if selected:
            self.word_selected = True
            self.logger.info(self.srcFileName)
            self.ui.lineEdit_Zeugnisvorlage.setText(str(self.srcFileName))
            # ToDo: reactivate
            self.logger.info("{}\n{}".format(str(self.Text_Feld_Namen), str(self.Array_CheckBox_Namen)))
            self.ui.pushButton_Excel_Zeugnisse.setEnabled(True)
            self.enable_buttons()
            self.creator_class.wichtige_meldung = ""

    def select_Excel_Zeugnis_File(self):
        # ToDo: filter richtig setzen nur *.xlsx anzeigen
        self.Excel_Filename = ""
        selected = False
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.FileMode())
        try:
            filter = "Office Excel File (*.xlsx)"
            excel_file = dialog.getOpenFileName(None, 'Wähle die Excle Zeugnis Datei', self.work_folder, filter,
                                                "*.xlsx")
            excel_file = str(excel_file[0].replace('/', os.path.sep))
            if os.path.exists(excel_file):
                self.Excel_Filename = excel_file
                selected = True
        except Exception as e:
            self.logger.error(e)
        if selected:
            self.logger.info(self.Excel_Filename)
            self.ui.lineEdit_Zeugnisse_xlsx.setText(self.Excel_Filename)

            self.creator_class.read_excel_file_to_pandas(self.Excel_Filename)

            # print(self.creator_class.df)
            self.excel_selected = True
            # self.enable_buttons()
            self.pandas_to_tableview(self.creator_class.df)

            # self.ui.pushButton_Ausgabe_Ordner.setEnabled(True)
            if not self.creator_class.wichtige_meldung == "":
                self.critical_messagebox("Critical", self.creator_class.wichtige_meldung)

    def critical_messagebox(self, title="Title", text="text"):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        # setting message for Message Box
        msg.setText(text)
        # setting Message box window title
        msg.setWindowTitle(title)
        # declaring buttons on Message Box
        msg.setIcon(QMessageBox.Critical)
        msg.setStandardButtons(QMessageBox.Ok)
        # start the app
        retval = msg.exec_()


    def pandas_to_tableview(self, pandas_data):
        self.logger.debug("pandas_to_tableview")
        TableModel = PandasModel(pandas_data)
        self.ui.tableView.setModel(TableModel)

    def getSelecet_Row(self):
        row = -1
        data = ""
        if self.ui.tableView.currentIndex().isValid():
            row = self.ui.tableView.currentIndex().row()
            data = self.ui.tableView.currentIndex().data()
        self.logger.info("row {} data {}".format(row, data))
        return row

    def create_single_selected_reference_word(self):
        self.logger.debug("create_single_selected_reference_word")
        row = self.getSelecet_Row()
        self.logger.info("row {}".format(row))
        word_file_names = []
        if not self.creator_class.destination_Folder == "" and \
                not self.creator_class.Excel_Filename == "" and \
                not self.creator_class.srcFileName == "":
            Generate_Data = self.creator_class.df.iloc[[row]].reset_index(drop=True)
            word_file_names = self.creator_class.create_wordDocxs_from_dataframe(Generate_Data)
        self.logger.info(word_file_names[0])
        self.write_infoline(str(word_file_names))

    def create_single_selected_reference_word_and_pdf(self):
        row = self.getSelecet_Row()
        word_file_names = []
        if not self.creator_class.destination_Folder == "" and \
                not self.creator_class.Excel_Filename == "" and \
                not self.creator_class.srcFileName == "":
            Generate_Data = self.creator_class.df.iloc[[row]].reset_index(drop=True)
            word_file_names = self.creator_class.create_wordDocxs_from_dataframe(Generate_Data)
        self.logger.info("create single selected word {}; row {}".format(str(word_file_names), str(row)))
        for each in word_file_names:
            self.logger.info(each)
            infile = self.output_folder + os.path.sep + each
            outfile = self.output_pdf_folder + os.path.sep + each.replace(self.converter.FileExtention(each), "pdf")
            self.logger.info(outfile)
            self.converter.covx_to_pdf(infile, outfile)
        self.write_infoline(outfile)

    def create_all_references_word(self):
        # row = self.getSelecet_Row()
        word_file_names = []
        if not self.creator_class.destination_Folder == "" and \
                not self.creator_class.Excel_Filename == "" and \
                not self.creator_class.srcFileName == "":
            # Generate_Data = self.creator_class.df.iloc[[row]].reset_index(drop=True)
            word_file_names = self.creator_class.create_wordDocxs_from_dataframe(self.creator_class.df)
        self.logger.info("create_all_references_word {}".format(str(word_file_names)))
        self.write_infoline(str(word_file_names))  # crashed if active

    def create_all_references_word_and_pdf(self):
        # row = self.getSelecet_Row()
        word_file_names = []
        if not self.creator_class.destination_Folder == "" and \
                not self.creator_class.Excel_Filename == "" and \
                not self.creator_class.srcFileName == "":
            # Generate_Data = self.creator_class.df.iloc[[row]].reset_index(drop=True)
            word_file_names = self.creator_class.create_wordDocxs_from_dataframe(self.creator_class.df)
        self.logger.info("create_all_references_word_and_pdf {}".format(str(word_file_names)))

        for each in word_file_names:
            self.logger.info(each)
            infile = self.output_folder + os.path.sep + each
            outfile = self.output_pdf_folder + os.path.sep + each.replace(self.converter.FileExtention(each), "pdf")
            self.logger.info(outfile)
            # self.ui.lineEdit_infoline.setText(str(outfile.split(os.path.sep)[-1]))
            self.converter.covx_to_pdf(infile, outfile)
        self.write_infoline("Alle Word und PDF-Dokumente erzeugt.")

    def convert_all_docx_to_pdf(self):
        self.logger.info("convert_all_docx_to_pdf {} {}".format(self.output_folder, self.output_pdf_folder))
        converter = Docx_to_PDF()
        converter.convert_all_docx_of_inputFolder_to_PDF_in_outputFolder(self.output_folder, self.output_pdf_folder)
        self.write_infoline("Alle Word Dokumente zu PDF konvertiert.")

    def combine_all_PDFs_in_selected_output_folder(self):
        PDFOutputName = self.output_pdf_folder + os.path.sep + "Output.pdf"
        # self.ui.lineEdit_infoline.setText(PDFOutputName)
        self.write_infoline(PDFOutputName)
        if os.path.exists(PDFOutputName):
            self.logger.debug("lösche PDF: {}".format(str(PDFOutputName)))
            os.remove(PDFOutputName)
        PDF_combiner = Combine_PDF()
        PDF_combiner.PDF_Combiner(self.output_pdf_folder, PDFOutputName)
        self.write_infoline("Alle PDFs zu {} zusammengefügt.".format(str(PDFOutputName)))

    def Toggle_actionAutomatisch_PDF_Ordner_anlegen(self):
        # self.show_hide_channel_buttons
        if self.auto_pdf_folder:
            self.auto_pdf_folder = False
            self.ui.actionAutomatisch_PDF_Ordner_anlegen.setChecked(self.auto_pdf_folder)
            self.ui.pushButton_PDF_Folder.setEnabled(not self.auto_pdf_folder)
            self.config_information["Autotomatisch_PDF_Ordner_anlegen"] = self.auto_pdf_folder
            self.write_json_config_information()
        else:
            self.auto_pdf_folder = True
            self.ui.actionAutomatisch_PDF_Ordner_anlegen.setChecked(self.auto_pdf_folder)
            self.ui.pushButton_PDF_Folder.setEnabled(not self.auto_pdf_folder)
            self.config_information["Autotomatisch_PDF_Ordner_anlegen"] = self.auto_pdf_folder
            self.write_json_config_information()

    def show_about(self):
        msgBox = QMessageBox()
        msgBox.setWindowTitle("About")
        msgBox.setWindowIcon(QIcon("Logo.png"))

        msgBox_Text = "SchoolReport_Excel2Word2PDF" + "\n" + \
                      "" + "\n" + \
                      "Version: " + __version__ + "\n" + \
                      "Author: " + __author__ + "\n" + \
                      "Repo: " + __repo__ + "\n"
        msgBox.setText(msgBox_Text)

        size = 14
        # print("font size", size)
        new_font = QFont("", size) # , QFont.Bold)
        msgBox.setFont(new_font)

        msgBox.setStandardButtons(QMessageBox.Ok)
        result = msgBox.exec()

    def Anleitung_oeffnen(self):
        # __Anleitung__
        os.startfile(__Anleitung__)

    def kill_winword(self):
        '''
            https://www.instructables.com/Python-Script-to-Stop-Windows-ApplicationsShutdown/
        '''
        button_reply = QMessageBox.question(self,
                                            "Word beenden?",
                                            "Alle Word Prozesse beenden?\n"
                                            "- Bitte vorher alle offenen Dokumente speichern!\n"
                                            "",
                                            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if button_reply == QMessageBox.Yes:
            reply = os.system(r"taskkill /f /im winword.exe")
            if reply == 0:
                self.write_infoline("Der Prozess WINWORD.EXE wurde beendet.")
            elif reply == 128:
                self.write_infoline("Der Prozess winword.exe wurde nicht gefunden.")
            else:
                self.write_infoline("Unbekannte Rückmeldung")
        else:
            # print("Nö", QMessageBox.Yes)
            pass

    def write_infoline(self, text):
        self.ui.lineEdit_infoline.setText(text)

    def dummy(self):
        # https://stackoverflow.com/questions/37035756/how-to-select-multiple-rows-in-qtableview-using-selectionmodel
        row = self.getSelecet_Row()
        indexes = self.ui.tableView.selectionModel().selectedRows()
        self.logger.info(row, indexes)
        # outfile = r"W:\jdsfjkhkdfsfjk\hijfsdhjfdskhk\file.xlsx"
        # print(outfile.split(os.path.sep))
        # print(outfile.split(os.path.sep)[-1])
        # self.ui.lineEdit_infoline.setText(outfile.split(os.path.sep)[-1])
