import os
from xlcrf import CRF

def script(inpath):
    outdir = '/tmp'
    outfile = os.path.basename(os.path.splitext(inpath)[0] + "_CRF.xlsx")
    outpath = os.path.join(outdir, outfile)
    ex1 = CRF()
    ex1.read_structure(inpath)
    ex1.create(outpath)
    print("output file: "+ outpath)
        
    
