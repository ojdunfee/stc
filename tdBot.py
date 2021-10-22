import pandas as pd
import numpy as np
import os
import calendar
import string

from co_data import ACCTS, KEYWORDS, EMPLOYEE, COS

def load_sheet(filename):
    df = pd.read_excel(filename, converters={'MCC/SIC Code': str, 'Originating Account Number': str})
    df = df[df['Merchant Name'] != 'AUTO PAYMENT DEDUCTION']
    return df

def get_company(df):
    acct = str(df.iloc[0, -1][-4:])
    return COS[acct]['sheet']

def accounts_by_mcc(cell):
    if cell in ACCTS.keys():
        return ACCTS[cell]
    elif str(cell) == 'nan':
        return '63004'
    else:
        return np.nan

def get_state_code(cell):
    if str(cell)[-4:] in EMPLOYEE.keys():
        return EMPLOYEE[str(cell)[-4:]]['state']
    else:
        return np.nan

def get_branch_code(cell):
    if str(cell)[-4:] in EMPLOYEE.keys():
        return EMPLOYEE[str(cell)[-4:]]['branch']
    else:
        return np.nan

def get_description(df):
    L = [cell for cell in df['Merchant Name']]
    for i, cell in enumerate(df['Originating Account Name']):
        if type(cell) == np.nan:
            continue
        else:
            if len(str(cell).split()) == 2 and cell != 'COMMERCIAL DEPARTMENT':
                name = str(cell).split()
                initials = name[0][0] + name[1][0]
                L[i] = '{}-{}'.format(initials, L[i])
            else:
                continue
    df['Description/Comment'] = [x for x in L]
    return df

def get_ic_partner_code(cell):
    if str(cell)[-4:] in EMPLOYEE.keys():
        return EMPLOYEE[str(cell)[-4:]]['ic_code']
    else:
        return np.nan

def get_department_code(cell):
    if str(cell)[-4:] in EMPLOYEE.keys():
        return EMPLOYEE[str(cell)[-4:]]['dept']
    else:
        return np.nan

def get_save_file(dir):
    for file in os.listdir(dir):
        if file.endswith('TD CARD.xlsx'):
            df = pd.read_excel(os.path.join(dir, file))
            date = df.iloc[-1, 0]
            date = '{}/{}/{}'.format(date.month, calendar.monthrange(date.year, date.month)[1], date.year)
            return '{} TD_Statement.xlsx'.format(date.replace('/', '_'))
            break

def generate_keywords(df):
    for i, cell in enumerate(df['Description/Comment']):
        cell = cell.translate(str.maketrans(' ', ' ', string.punctuation))
        for word in cell.split():
            for k, v in KEYWORDS.items():
                if word in v:
                    df.loc[i, 'No'] = k
    return df

def fix_sheet(df):
    df.loc[df[df['Direct Unit Cost'] < 0].index.tolist(), 'No'] = '19999'
    df.loc[df[df['No'] == '63000'].index.tolist(), 'Dept Code'] = '00'
    df.loc[df[df['No'] == '63003'].index.tolist(), 'Dept Code'] = '00'
    df.loc[df[df['Description/Comment'] == 'STANDARD VCF 4.4 100'].index.tolist(), 'No'] = '63004'
    df.loc[df[df['Description/Comment'] == 'STANDARD VCF 4.4 100'].index.tolist(), 'State'] = '00'
    df.loc[df[df['Description/Comment'] == 'STANDARD VCF 4.4 100'].index.tolist(), 'Branch Code'] = '000'
    df.loc[df[df['Description/Comment'] == 'STANDARD VCF 4.4 100'].index.tolist(), 'Dept Code'] = '00'
    return df

def generate_other_columns(df):
    df['Type'] = ['G/L Account'] * len(df)
    df['IC Partner Ref Type'] = ['G/L Account'] * len(df)
    df['Quantity'] = [1] * len(df)
    df['Direct Unit Cost'] = [round(cell, 2) for i, cell in enumerate(df['Billed Amount'])]
    df['IC Partner Reference'] = [np.nan] * len(df)
    return df[['Type', 'No', 'State', 'Branch Code', 'Dept Code', 'Description/Comment', 'Quantity', 'Direct Unit Cost', 'IC Partner Ref Type', 'IC Partner Code', 'IC Partner Reference']]

def generate_sheet(filename):
    df = load_sheet(filename)
    company = get_company(df)
    df['No'] = df['MCC/SIC Code'].apply(accounts_by_mcc)
    df['State'] = df['Originating Account Number'].apply(get_state_code)
    df['Branch Code'] = df['Originating Account Number'].apply(get_branch_code)
    df = get_description(df)
    df['IC Partner Code'] = df['Originating Account Number'].apply(get_ic_partner_code)
    df['Dept Code'] = df['Originating Account Number'].apply(get_department_code)
    df = generate_keywords(df)
    df = generate_other_columns(df)
    df = fix_sheet(df)
    return df, company

def create_spreadsheet(filename, frames):
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        for i in range(len(frames)):
            frames[i][0].to_excel(writer, sheet_name=frames[i][1], index=False)
            workbook = writer.book
            worksheet = writer.sheets[frames[i][1]]
            num_format = workbook.add_format({'num_format':'##0.00'})
            worksheet.set_column_pixels('A:A', 80)
            worksheet.set_column_pixels('B:B', 42)
            worksheet.set_column_pixels('C:C', 40)
            worksheet.set_column_pixels('D:D', 83)
            worksheet.set_column_pixels('E:E', 72)
            worksheet.set_column_pixels('F:F', 202)
            worksheet.set_column_pixels('G:G', 61)
            worksheet.set_column_pixels('H:H', 105, cell_format=num_format)
            worksheet.set_column_pixels('I:I', 127)
            worksheet.set_column_pixels('J:J', 104)
            worksheet.set_column_pixels('K:K', 137)
            

        
def generate_journal(dir, save_file):
    frames = []
    filename = ''
    for file in os.listdir(dir):
        if file.endswith('.xlsx'):
            df, sheet = generate_sheet(os.path.join(dir, file))
            if not filename:
                filename = save_file
            frames.append((df, sheet))

    frames = sorted(frames, key=lambda x: x[1])
    create_spreadsheet(filename, frames)