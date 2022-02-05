#!/usr/bin/env python
# -*- coding: iso-8859-1 -*-
'''
Auslesen aller nötigen Word Informationen (Textfelder, CheckBox Namen, ...)
 - Word-Doropdown auswählen
 - CheckBox abhaken
 - Text ausfüllen

'''
import win32com
import win32com.client
import os
import os.path
import WordDocument
import re
import logging

class helper_Word():
    def __init__(self):
        # Word-Application referenzieren
        self.logger = logging.getLogger("SR_E2W2P.helper_Word")
        self.logger.info("init helper_Word")

        self.app = win32com.client.Dispatch("Word.Application")

        self.filename = None # os.path.abspath("*.docx")
        self.doc = None
        self.Field_dict = {}
        self.Text_FieldArray_Row_Names = []
        self.Text_Field_Names = []
        pass

    def init_file(self, filename, app_visible = False):
        # Word sichtbar machen. Kann später weggelassen werden
        self.app.Visible = app_visible

        self.filename = os.path.abspath(filename)

        # Quelldatei öffnen
        self.doc = self.app.Documents.Open(self.filename, ReadOnly=False)
        self.WD = WordDocument.WordDocument(self.filename)

    def save_and_close(self):
        try:
            self.doc.Save()
        except Exception as e:
            self.logger.error(e)

        try:
            self.doc.Save()
            self.doc.Close()
            return True
        except Exception as e:
            self.logger.error(e)


    def set_WordField_Text(self, Fieldname, Text):
        # set DoropDown
        # word_doc.doc.FormFields("Note_Musik").DropDown.Value = 4
        # 1 = not selected
        # 2 = sehr gut
        # 3 = gut
        # ...
        # 7 = ungenügend
        if "Note" in Fieldname:
            if Text == "sehr gut":
                Text = 1
            elif Text == "gut":
                Text = 2
            elif Text == "befriedigend":
                Text = 3
            elif Text == "ausreichend":
                Text = 4
            elif Text == "mangelhaft":
                Text = 5
            elif Text == "ungenügend":
                Text = 6
            elif Text == "nicht erteilt":
                Text = 7
            else:
                Text = 0
            # print("Note")
            note = int(Text) + 1
            # print("Note =", note - 1 )
            self.doc.FormFields(Fieldname).DropDown.Value = note
        else:
            if Text == None or Text == "None" :
                Text = ""

            if len(Text) >= 254:
                # print("TextForm: {} {}".format(Fieldname,Text))
                self.doc.FormFields(Fieldname).Result = ""  # clear first or remove to append
                self.doc.FormFields(Fieldname).Select()
                self.app.Selection.MoveRight()
                self.app.Selection.Text = Text
            else:
                self.doc.FormFields(Fieldname).Result = Text

    def get_WordField_Text(self, Fieldname):
        return self.doc.FormFields(Fieldname).Result

    def check_checkbox(self, check_box_name):
        try:
            self.doc.FormFields(check_box_name).CheckBox.Value = True
        except Exception as e:
            self.logger.error("check_checkbox {} {} ".format(str(check_box_name), e))

    def uncheck_checkbox(self, check_box_name):
        try:
            self.doc.FormFields(check_box_name).CheckBox.Value = False
        except Exception as e:
            self.logger.error("uncheck_checkbox {} {} ".format(str(check_box_name), e))

    def get_all_fieldnames_from_word(self):
        '''
        return Dict {Field_Name_1: [None, None], #for single field
                     Field_Name_2: [0, 0], #for multiple checkboxes field
                    }
        '''
        self.WD = WordDocument.WordDocument(self.filename)
        Documentelements = self.WD.get_form_data()
        # print(Documentelements)
        Dokument_Element_Dict = {}
        for each in Documentelements:
            #t = re.search(r"( - |- | -| |)(|\()[0-9]{3,4}p(\)|)", FolderName, re.IGNORECASE)
            t = re.findall(r"([0-9]{1,2}[_][0-9])", each, re.IGNORECASE)
            row = None
            column= None
            regresult = (t[0] if t else None)
            Element_name = each
            if not regresult == None:
                row = regresult.split("_")[0]
                column = regresult.split("_")[1]
                Element_name = each.replace(regresult, "")
            if Element_name in Dokument_Element_Dict.keys():
                #print(Element_name, Dokument_Element_Dict[Element_name])
                templist = Dokument_Element_Dict[Element_name]
                templist.append([row, column])
                Dokument_Element_Dict[Element_name] = templist
            else:
                Dokument_Element_Dict[Element_name] = [[row, column]]
        self.logger.info(Dokument_Element_Dict)
        self.Field_dict = Dokument_Element_Dict
        return Dokument_Element_Dict

    def find_CheckBox_Array_dimentions(self, arrayname=""):
        if self.Field_dict == {}:
            self.Field_dict = self.get_all_fieldnames_from_word()

        row = 0
        column = 0
        for each in self.Field_dict[arrayname]:
            if int(each[0]) > row:
                row = int(each[0])
            if int(each[1]) > int(column):
                column = int(each[1])
        return [row, column]

    def find_arrays_and_text_field_names(self):
        self.Text_FieldArray_Row_Names = []
        self.Text_Field_Names = []
        for each in self.Field_dict:
            if self.Field_dict[each][0][0] == None and self.Field_dict[each][0][1] == None:
                self.Text_Field_Names.append(each)
            else:
                Dimensions = self.find_CheckBox_Array_dimentions(each)
                for i in range(Dimensions[0]+1):
                    self.Text_FieldArray_Row_Names.append(each + str(i) + "_")
                if Dimensions[0] == 0:
                    self.Text_FieldArray_Row_Names.append(each + "0" + "_")
            # self.logger.info(self.Text_Field_Names)
            # self.logger.info(self.Text_FieldArray_Row_Names)

if __name__ == '__main__':
    mydoc = helper_Word()
    mydoc.init_file("Zeugnis_Stufe2_2018 - Kopie.docx")
    mydoc.set_WordField_Text("Name", "Fritz Fratz")
    print(mydoc.doc.SaveAs("Zeugnis_Stufe2_2018 - Fritz Fratz.docx"))

