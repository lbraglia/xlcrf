import pandas as pd

class Column:
    def __init__(self):
        self.variable      = variabile
        self.description   = descrizione_e_unita_misura
        self.validate      = tipo
        self.list_name     = id_elenco
        self.criteria      = criterio
        self.value         = valore
        self.minimum       = minimo
        self.maximum       = massimo
        self.ignore_blank  = ignora_celle_vuote
        self.drop_down     = elenco_nella_cella
        self.show_input    = input_mostra
        self.input_title   = input_titolo
        self.input_message = input_messaggio
        self.show_error    = errore_mostra
        self.error_type    = errore_tipo
        self.error_title   = errore_titolo
        self.error_message = errore_messaggio

    

class Sheet:
    def __init__(self, xl, sheet):
        self.df = xl.parse(sheet)
        self.sheetname = sheet
        # cicla sulle righe del data.frame
        self.columns = []
        # parse columns


class CRF:
    def __init__(self):
        # a CRF is a list of sheets
        self.sheets = []
    
    def read_structure(self, f):
        """
        Import a dataset structure
        """
        print("Reading structure file: " + f)
        xl = pd.ExcelFile(f)
        # individua gli sheet chiave
        # print(xl.sheet_names)  # see all sheet names
        data_sheets = [s for s in xl.sheet_names if s not in
                       ['modalita_output', 'modalita_struttura']]
        for s in data_sheets:
            self.sheets.append(Sheet(xl, s))


    def create(self, f):
        """
        Create an xlsx template according to the structure
        """
        print("Creating CRF file: " + f)


    
if __name__ == '__main__':
    infile = '/home/l/xlcrf/structs/example1.xlsx'
    outfile = '/tmp/example1_crf.xlsx'
    ex1 = CRF()
    ex1.read_structure(infile)
    ex1.create(outfile)
    print(ex1.sheets[0].df)
