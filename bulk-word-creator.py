import os, sys
from docxtpl import DocxTemplate
import pandas as pd
os.chdir(sys.path[0])

class ContractDoc():

    def __init__(self):
        pass
    def make_contract(self, template, csv):
        '''
        Creates contract file(s) from a template and context files.  
        Template is a MS Word document and csv is a CSV file
        containing a list of placeholders labels and values.

        template: MS Word document (docx).
        csv: CSV file containing the list of placeholders labels as the header, 
        placeholder values as the fields.

        Will create a file for each row in the csv file (minus the header).
        '''
        placeholders = pd.read_csv(csv)

        for record in placeholders.to_dict(orient="records"):
            doc = DocxTemplate(template)
            doc.render(record)
            doc.save('test.docx')
