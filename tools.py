import pandas as pd

from datetime import datetime
from workalendar.america import Brazil

cal = Brazil()
cal.holidays(datetime.now().year)


def txt_to_df(source, cols, fcol, split_element='|', encode='iso-8859-1', h=1, type=str):
    df = pd.read_csv(source, usecols=cols, index_col=False, sep=split_element, encoding=encode, header=h, error_bad_lines=False, decimal=',', dtype=type)
    df = df[pd.to_numeric(df[fcol], errors='coerce').notnull()]
    df.columns = [x.strip() for x in df.columns]

    try:
        df[df.columns] = df.apply(lambda x: x.str.strip())
    except:
        pass

    return df

def my_merge(df, df_from, df_key, cols):  
    df_from_key = cols[0]
    df_from = df_from.astype({df_from_key: str})
    
    if len(cols) == 1:
        cols = df_from.columns
    
    df = df.astype({df_key: str})    
    df = pd.merge(df, df_from[cols], left_on=df_key, right_on=df_from_key, how='left')
    df.drop(columns=[df_from_key], inplace=True)

    return df


def extract_sap(pathfile, h, index_col=None):
    df = pd.read_table(pathfile, sep="|", header=h, encoding='ISO-8859-1', low_memory=False)
    df = df.drop_duplicates(subset=index_col, keep='last', inplace=False)
    
    # Limpando apóstrofes/aspas nos valores
    df = df.replace("'","", regex=True).replace('"','"', regex=True)
    df = df.astype(str)
    
    # Limpando espaços iniciais e finais dos valores e nome de campos
    for col in df.columns:
        df[col] = df[col].str.strip()
        
    df.rename(columns=lambda x: x.strip(), inplace=True)
    # columns = pd.io.parsers.ParserBase({'names':df.columns, 'usecols':None})._maybe_dedup_names(df.columns)
    # df.set_axis(columns, axis=1, inplace=True)
    
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique():
        cols[cols[cols == dup].index.values.tolist()] = [dup + '.' + str(i) if i != 0 else dup for i in range(sum(cols == dup))]
    
    # rename the columns with the cols list.
    df.columns = cols    
    
    # Apenas para arquivos SAP
    # Primeira e ultima coluna sao nulas
    df.drop(columns=df.columns[[0,-1]], inplace=True)
    
    # Eliminando linhas nulas e de somatório
    df.drop(index=df.index[df.isnull().all(1)],inplace=True, errors='ignore')
    df.drop(index=df.index[df.iloc[:,0]=='*'], inplace=True, errors='ignore')
    df.drop(index=df.index[df.iloc[:,0].str.contains('---')], inplace=True, errors='ignore')
    df.drop(index=df.index[df.iloc[:,0] == 'nan'], inplace=True, errors='ignore')
    df.drop(index=df.index[df.iloc[:,0] == df.columns[0]], inplace=True, errors='ignore')
    df = df.reset_index().drop(columns='index')
    
    return df


def convert_to_number(df, cols):
    for col in cols:
        df[col] = df[col].str.replace('.', '')
        df[col] = df[col].str.replace(',', '.')
        df[col] = pd.to_numeric(df[col])


def convert_date_sap_to_br(df, cols):
    for col in cols:
        df_col = df[col]        
        dt_col = pd.to_datetime(df_col, dayfirst=True, errors = 'coerce')          
        df[col] = dt_col


def calculate_work_days(dt, d):
    try:
        return cal.add_working_days(dt, d)
    except:
        return ''




def cnpj_string(cnpj):
    if len(cnpj) < 14:
        cnpj = '0' + cnpj

    return cnpj