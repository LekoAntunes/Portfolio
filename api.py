import json
import pandas as pd 

with open('setup.json') as json_file:
    json_data = json.load(json_file)

path = json_data['main_path']    
report = json_data['reports'] ['br_ifrs15']
source = path + report['file']

df = pd.read_excel(source, sheet_name='HistRecReceita')


def add_market():
    
    for r in df.index:
        e = df.loc[r, 'Mercado']
        
        if pd.isnull(e):
            k = df.loc[r, 'TipFt']
            vl = 'BR'
            
            if k == 'ZVEX' or k == 'ZEXT':
                vl = 'EXPO'
                
            df.loc[r, 'Mercado'] = vl


def add_accounting_type():
    
    for r in df.index:
        e = df.loc[r, 'Contabilização Fatura 2']
        
        if pd.isnull(e):
            t = df.loc[r, 'Texto']
            u = df.loc[r, 'Usuário Fatura 2']
            vl = 'INTEGRACAO'
            
            if t.startswith('Finalizado Manualmente (AANTUNES') or t.startswith('Finalizado Manualmente (PYKOSZ'):
                vl = 'ZTSD401'
                
            elif u == 'AANTUNES' or u == 'PYKOSZ':
                vl = 'VF01'
                
            df.loc[r, 'Contabilização Fatura 2'] = vl
    

def add_carrier_nickname():
    
    source_temp = report['support_data']['carrire_nickname']
    df_temp = pd.read_excel(source_temp, sheet_name='Query')
    
    test = pd.merge(df[pd.isnull(df['Apelido'])], df_temp, how='left', on='Cod Transp')
    


# API
# add_market()
#GetUserInvoiceROL_Q1() ### Verificar possibilidade de add esse item na query principal ou criar um novo job

# add_accounting_type()
#GetDocTranspInfo_Q1() ### Verificar possibilidade de add esse item na query principal ou criar um novo job

add_carrier_nickname()