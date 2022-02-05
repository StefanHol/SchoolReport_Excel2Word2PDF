
'''
Auslesen aller nötigen Word Informationen (Textfelder, CheckBox Namen, ...)
XML Informationen Lesen.
'''

import zipfile
from lxml import etree
from collections import OrderedDict
import logging
# import glob
import os
import csv


class WordDocument:

    NSMAP = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    def __init__(self, filename):
        """Read a Word document in .docx format."""
        self.logger = logging.getLogger("SR_E2W2P.WordDokument")
        self.logger.info("init WordDocument")

        self.filename = filename
        with open(filename, 'rb') as f:
            archive = zipfile.ZipFile(f)
            xml_content = archive.read('word/document.xml')
        self.xml_content = xml_content
        self.root = None
        self.cells = None
        self.NSMAP = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    def get_root(self):
        """Get the root element of the XML tree of a Word document."""
        if self.root is None:
            self.root = etree.fromstring(self.xml_content)
        return self.root
    
    def get_name(self, elt):
        """Get the name of a form input, or None if the name is not found."""
        name = elt.find('.//w:ffData/w:name', self.NSMAP)
        if name is not None:
            return name.get('{%(w)s}val' % self.NSMAP)
    
    def get_checked(self, elt):
        """Returns 1 if a checkbox is checked, 0 if it is not checked,
        or None if the element is not a checkbox."""
        checkBox = elt.find('.//w:ffData/w:checkBox', self.NSMAP)
        if checkBox is not None:
            return len(checkBox.findall('w:checked', self.NSMAP))
    
    def get_text_input(self, cell):
        """Get the value of a text input, or None if the cell 
        is not a text input."""
        for elt in cell.findall('w:r', self.NSMAP):
            if elt.find('w:rPr/w:noProof', self.NSMAP) is not None:
                text = elt.find('w:t', self.NSMAP).text
                return text
        elt = cell.find('w:pPr/w:pStyle', self.NSMAP)
        if elt is not None and elt.get('{%(w)s}val' % self.NSMAP) == 'FieldText':
            return cell.find('w:r/w:t', self.NSMAP).text
    
    def get_form_elements(self):
        """Return all form elements as an iterator."""
        root = self.get_root()
        return root.iterfind('.//w:p', self.NSMAP)
        
        
    def get_form_data(self):
        """Return all form data as an ordered dictionary."""
        data = OrderedDict()
        
        for p in self.get_form_elements():
            name = self.get_name(p)
            # print(name)
            try:
                if name is not None and not(name in data):
                    value = self.get_checked(p)
                    if value is None:
                        value = self.get_text_input(p)
                    if value == u'\u2002':
                        value = ''
                    data[name] = value
            except Exception as e:
                self.logger.error("error in WordDocument.py -> get_form_data: {} ".format(e))
        return data

def write_csv(data_list, csv_file_name, **kwargs):
    """Write data from multiple Word documents to a single CSV file"""
    if not isinstance(data_list, list):
        data_list = [data_list]
    columns = OrderedDict()
    for data in data_list:
        for key in data.keys():
            if not columns.has_key(key):
                columns[key] = 1
       
    with open(csv_file_name, 'wb') as f:
        writer = csv.writer(f, kwargs)
        writer.writerow(columns.keys())
        for data in data_list:
            writer.writerow([data.get(key, '') for key in columns])

def read_dir(path = '.'):
    data_list = []
    for filename in os.listdir(path):
        if filename.lower().endswith('.docx') and not filename[0] == '~':
            full_filename = os.path.join(path, filename)
            doc = WordDocument(full_filename)
            data = doc.get_form_data()
            data_list.append(data)
    return data_list
