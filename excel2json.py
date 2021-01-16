# -*- coding: utf-8 -*-
"""
Created on Mon Jan 11 21:58:22 2021

@author: olang
"""
import csv, json, os, sys
import pandas as pd

def xlsx2csv(file):
    filename = os.path.splitext(file)[0]
    if file.endswith('.xlsx'):
        data_xls = pd.read_excel(file, dtype=str, index_col=None)
        file = (filename+'.csv')
        data_xls.to_csv(file, encoding='utf-8', index=False)
        return file
    elif file.endswith('.xls'):
        data_xls = pd.read_excel(file, dtype=str, index_col=None)
        file = (filename+'.csv')
        data_xls.to_csv(file, encoding='utf-8', index=False)
        return file     
    else:
        return file

def main(argv):
    fields = []
    if len(argv) == 2:
        dire = argv[1]
        
        for file in os.listdir(dire):
            file = xlsx2csv(file)
            if file.endswith('.csv'):
                filename = os.path.splitext(file)[0]
                
                if not os.path.exists(filename):
                    os.mkdir(filename)
                
                fn = open(file,'r',encoding="utf8")
                r = csv.DictReader(fn, fieldnames = ("SPEC","COD. ET","Descrição","NM","DN"))
                
                for row in r:
                    frame = {"SPEC":row["SPEC"],
                             "COD. ET":row["COD. ET"],
                             "Descrição":row["Descrição"],
                             "NM":row["NM"],
                             "DN":row["DN"]}
                    if row["NM"] not in fields:
                        fields.append(row["NM"])
                        out = json.dumps(frame, indent=4, ensure_ascii=False)
                        f = open(filename+'/'+filename[0:-1]+'_'+fields[-1]+'.json','w',encoding='utf8')
                        f.write(out)
                        f.close()
                    else:
                        continue
                fields = []
                fn.close()
                if os.path.exists((filename)+'.xlsx') or os.path.exists((filename)+'.xls'):
                    os.remove(file)
            else:
                continue
    else:
        print('Error! Too many args')
    
if __name__ == '__main__':
    main(sys.argv)




        

                
                
        