from pprint import pprint
import pandas as pd
import xlsxwriter


# XLSX supports 1048576 rows and 16384 columns
xlsx_rows = 1048576
# numero di righe per cui è possibile inserire dati. Le altre son bloccate
n_fillable_rows = 10000 
protection_pw = 'asd'

def excel_col(n):
    # https://stackoverflow.com/questions/19153462
    excel_col = str()
    div = n
    while div:
        (div, mod) = divmod(div-1, 26) # will return (x, 0 .. 25)
        excel_col = chr(mod + 65) + excel_col
    return excel_col

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
    
class Column:
    def __init__(self, prog, struct):
        self.index         = prog
        self.variable      = struct['variabile']
        self.description   = struct['descrizione_e_unita_misura']
        self.validate      = import_validate(struct['tipo'])
        self.list_name     = struct['id_elenco']
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

    def export(self, ws):
        prog_col = self.index
        # title
        ws.write(0, prog_col, self.variable)
        

        
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

    def export(self, ws, formats):
        # Data
        for c in self.columns:
            c.export(ws)
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

        
class CRF:
    def __init__(self):
        # a CRF is a dict of sheets (each called by its name)
        self.sheets = {}
        # dict di modalità
        self.modalita = {}
        
    def read_structure(self, f):
        """
        Import a dataset structure
        """
        print("Reading structure file: " + f)
        xl = pd.ExcelFile(f)
        # importa gli sheet
        data_sheets = [s for s in xl.sheet_names if s not in
                       ['modalita_output', 'modalita_struttura']]
        for s in data_sheets:
            self.sheets[s] = Sheet(xl, s)

        # importa le modalità e crea un dict
        modalita = xl.parse('modalita_output')

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
            s.export(ws, formats)

        wb.close()

        

    
if __name__ == '__main__':
    infile = '/home/l/xlcrf/structs/example1.xlsx'
    outfile = '/tmp/example1_crf.xlsx'
    ex1 = CRF()
    ex1.read_structure(infile)
    ex1.create(outfile)
