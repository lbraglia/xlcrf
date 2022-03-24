from pprint import pprint
import numpy as np
import pandas as pd
import xlsxwriter

# numero di righe per cui è possibile inserire dati. Le altre son bloccate
n_fillable_rows = 10000 
protection_pw = 'asd'

# def excel_col(n):
#     # https://stackoverflow.com/questions/19153462
#     excel_col = str()
#     div = n
#     while div:
#         (div, mod) = divmod(div-1, 26) # will return (x, 0 .. 25)
#         excel_col = chr(mod + 65) + excel_col
#     return excel_col

def import_validate(x):
    if (x == 'intero'):
        return "integer"
    elif (x == 'decimale'):
        return "decimal"
    elif (x == 'elenco'):
        return "list"
    elif (x == 'data'):
        return "date"
    elif (x == 'ora'):
        return "time"
    else:
        return "any" # default

def import_criteria(x):
    if (x == 'tra'):
        return 'between'
    elif (x == 'non compreso tra'):
        return 'not between'
    else:
        return x

def import_sino_deftrue(x):
    if (x != ''):
        return x == 'Sì'
    else:
        return True # default

def import_id_elenco(x):
    if (pd.isna(x)):
        return None
    else:
        return x

class Column:
    def __init__(self, prog, struct):
        self.index         = prog
        self.variable      = struct['variabile']
        self.description   = struct['descrizione_e_unita_misura']
        self.validate      = import_validate(struct['tipo'])
        self.id_elenco     = import_id_elenco(struct['id_elenco'])
        self.criteria      = import_criteria(struct['criterio'])
        self.value         = struct['valore']
        self.minimum       = struct['minimo']
        self.maximum       = struct['massimo']
        self.ignore_blank  = import_sino_deftrue(struct['ignora_celle_vuote'])
        self.drop_down     = import_sino_deftrue(struct['elenco_nella_cella'])
        self.show_input    = import_sino_deftrue(struct['input_mostra'])
        self.input_title   = struct['input_titolo']
        self.input_message = struct['input_messaggio']
        self.show_error    = import_sino_deftrue(struct['errore_mostra'])
        self.error_type    = struct['errore_tipo']
        self.error_title   = struct['errore_titolo']
        self.error_message = struct['errore_messaggio']

    def export(self, ws, modalita):
        prog_col = self.index
        # title
        ws.write(0, prog_col, self.variable)
        # data validation from 1 to fillable_rows
        valid_dict = {
            "validate"       : self.validate     ,
            "criteria"       : self.criteria     ,
            "value"          : self.value        ,
            "minimum"        : self.minimum      ,
            "maximum"        : self.maximum      ,
            "ignore_blank"   : self.ignore_blank ,
            "dropdown"       : self.drop_down    ,
            "show_input"     : self.show_input   ,
            "input_title"    : self.input_title  ,
            "input_message"  : self.input_message,
            "show_error"     : self.show_error   ,
            "error_type"     : self.error_type   ,
            "error_title"    : self.error_title  ,
            "error_message"  : self.error_message
        }
        # Add modalita per le categoriche
        if (self.id_elenco == None):
            # print("niente")
            pass
        else:
            # print(modalita[self.id_elenco])
            valid_dict['source'] = modalita[self.id_elenco]

        ws.data_validation(1, prog_col, n_fillable_rows, prog_col, valid_dict)

        
class Sheet:
    def __init__(self, xl, sheetname):
        sheet = xl.parse(sheetname)
        sheet = sheet.reset_index() 
        self.sheetname = sheetname
        # estrai i dati delle colonne dell'output
        # ciclando sulle righe del data.frame di input
        self.columns = []
        for index, row in sheet.iterrows():
            self.columns.append(Column(index, row))

    def export(self, ws, formats, modalita):
        # Data
        for c in self.columns:
            c.export(ws, modalita)
        # formatting the worksheet
        ws.set_row(0, cell_format = formats['title'])
        # blocca riquadri
        ws.freeze_panes('A2')
        # protezione della prima riga da modifiche
        for r in range(1, n_fillable_rows):
            ws.set_row(row = r, cell_format = formats['unlocked'])
        # Meglio farlo a mano se no non si riescono a fare modifiche alla
        # larghezza colonne
        # ws.protect(password = protection_pw)


def parse_modalita(df):
    grouped = df.groupby('list_name')
    modalita = {}
    for name, group in grouped:
        modalita[name] = list(group['modalita'])
    return modalita
        
        
class CRF:
    def __init__(self):
        # a CRF is a dict of sheets (each called by its name)
        self.sheets = {}
        
    def read_structure(self, f):
        """
        Import a dataset structure
        """
        print("Reading structure file: " + f)
        xl = pd.ExcelFile(f)

        # importa le modalità impiegate
        self.modalita = parse_modalita(df = xl.parse('modalita_output'))
        
        # importa gli sheet e dai le modalita come dict
        data_sheets = [s for s in xl.sheet_names if s not in
                       ['modalita_output', 'modalita_struttura']]
        for s in data_sheets:
            self.sheets[s] = Sheet(xl, s)
        

    def create(self, f):
        """
        Create an xlsx template according to the structure
        """
        print("Creating CRF file: " + f)
        wb = xlsxwriter.Workbook(f)
        formats = {
            'title' : wb.add_format({'bold': True}),
            'unlocked' : wb.add_format({'locked': False})
        }
        # raw data
        for k, s in self.sheets.items():
            ws = wb.add_worksheet(k)
            s.export(ws, formats, self.modalita)
        wb.close()



    
if __name__ == '__main__':
    infile = '/home/l/xlcrf/structs/example1.xlsx'
    outfile = '/tmp/example1_crf.xlsx'
    ex1 = CRF()
    ex1.read_structure(infile)
    ex1.create(outfile)
