import pandas as pd
import numpy as np
import re
from co_data import *

def get_state(cell):
    '''Generate state codes based off file numbers'''
    m = re.match('\d+(\w{2,3}-\d{2})*', cell)
    try:
        return m.groups()[0].split('-')[-1]
    except AttributeError:
        return '01'

def get_dept(cell):
    '''Generate department codes based off existing accounting codes'''
    if cell.startswith('6'):
        return '02'
    return '00'

def debits(cell):
    '''Generate debits if amount of invoice is less than 0'''
    if cell < 0:
        return round(abs(cell), 2)
    return np.nan

def credits(cell):
    '''Generate credits if amount is greater than 0'''
    if cell >= 0:
        return round(cell, 2)
    return np.nan

def report(escrow, df):
    '''
    Group the cash receipt report, if there is a shortage concatenate the file number and closing agents initials for review
    Total the accounts in totals and generate the proper business central account code 
    '''
    cash = df[~df.AcctCode.isin(['66302','66300'])]
    shorts = df[df.AcctCode.isin(['66302', '66300'])]
    shorts.loc[(shorts.CloseAgent.isna()), 'CloseAgent'] = 'None'

    frame = cash.groupby(['TitleCoNum','Branch','St','AcctCode']).agg({
        'Invoice Line Total':'sum',
        'Dept':'first',
        'CloseAgent':'first',
        'Date': 'first'
    }).reset_index()

    df = pd.concat([frame, shorts], ignore_index=True)
    df['Type'] = ['G/L Account'] * len(df)
    df['Account Desr'] = ['{} {}'.format(df['File Number'][i], df['CloseAgent'][i]) if df.AcctCode[i] in ['66300','66302'] else np.nan for i in range(len(df))]
    df['Description Reference'] = ['{} RQ DEP'.format(df['Date'][i]) for i in range(len(df))]
    df['Debits'] = df['Invoice Line Total'].apply(debits)
    df['Credits'] = df['Invoice Line Total'].apply(credits)

    totals = pd.DataFrame({
        'Date': df['Date'][0],
        'Type': ['Bank Account'],
        'AcctCode': [ACCOUNTS[escrow]['bank']],
        'St': ['00'],
        'Branch': ['000'],
        'Dept': ['00'],
        'Account Desr': [np.nan],
        'Description Reference': ['{} RQ DEP'.format(df['Date'][0])],
        'Debits': [round(df['Invoice Line Total'].sum(), 2)],
        'Credits': [np.nan],
    })

    df = pd.concat([df, totals], ignore_index=True)
    df.rename(columns={'AcctCode':'Account'}, inplace=True)

    return df[['Date','Type','Account','St','Branch','Dept','Account Desr','Description Reference','Debits','Credits']]

def totals(escrow, df):
    '''
    Generate revenue and count totals for each account based off OrderCategory
    If additional revenue is being added to an existing invoice filter out to only count invoices with a premium being charged
    '''
    debits, credits, states = list(), list(), list()
    accounts, branches = list(), list()
    dates, description_references = list(), list()

    for (tco, date, branch, st, oc), frame in df.groupby(['TitleCoNum', 'Date', 'Branch', 'St','OrderCategory']):
        if oc in [1,4]:
            debits.append(round(frame['Invoice Line Total'].sum(), 2))
            debits.append(len(set(frame[frame.AcctCode == '40000']['File Number'])) - len(set(frame[(frame.SortField == 2) & (frame.AcctCode == '40000')]['File Number'])))
        elif oc in [2,5]:
            debits.append(round(frame['Invoice Line Total'].sum(), 2))
            debits.append(len(set(frame[frame.AcctCode == '40002']['File Number'])) - len(set(frame[(frame.SortField == 2) & (frame.AcctCode == '40002')]['File Number'])))
        else:
            debits.append(round(frame['Invoice Line Total'].sum(), 2))
            debits.append(len(set(frame['File Number'].tolist())))
            
        accounts.append(CLOSING[oc]['revenue'])
        accounts.append(CLOSING[oc]['count'])
        credits += [np.nan] * 2
        states += [st] * 2
        branches += [branch] * 2
        dates += [date] * 2
        description_references += ['{} {}'.format(date, 'RQ DEP')] * 2
    
    dates.append(date)

    description_references.append('{} {}'.format(date, 'RQ DEP'))
    branches.append('000')
    states.append('00')
    accounts.append('99998')
    debits.append(np.nan)
    credits.append(np.nan)
    
    report =  pd.DataFrame({'Date': dates,
                            'Type': ['G/L Account'] * len(debits),
                            'Account': accounts,
                            'St': states,
                            'Branch': branches,
                            'Dept': ['00'] * len(debits),
                            'Description Reference': description_references,
                            'Debits': debits,
                            'Credits': credits})
    
    report = report[~report.Account.str.startswith('?')]
    report = report[report.Debits != 0]
    return report

def clean_data(filename):
    '''
    Upload the spreadsheet into pandas, generate branch codes based off TitleCoNum, departments based off accounting codes,
    fix entries with errors.
    '''
    if filename.endswith('csv'):
        df = pd.read_csv(filename, converters={'AcctCode':str, 'TitleCoNum':str})
    else:
        df = pd.read_excel(filename, converters={'AcctCode':str, 'TitleCoNum':str})

    df.Description = df.Description.str.lower()

    df.dropna(subset=['Invoice Line Total'], inplace=True)
    df = df[df['Invoice Line Total'] != 0]
    
    df['Branch'] = df.TitleCoNum.map(lambda x: BRANCHES[x] if x in BRANCHES.keys() else '000')
    df['St'] = df['File Number'].apply(get_state)

    for k, v in CASH_DESCRIPTIONS.items():
        descr = '|'.join(v)
        df.loc[(df.Description.str.contains(descr)) & (df.AcctCode != k) & (df.AcctCode != '19999'), 'AcctCode'] = k
    
    df.Description = df.Description.str.title()
    df.loc[(df.AcctCode == '40000') & (df.OrderCategory.isin([2,5])), 'AcctCode'] = '40002'

    df.OrderCategory.replace(25, 8, inplace=True)

    df['Dept'] = df.AcctCode.apply(get_dept)
    df['Date'] = pd.to_datetime(df['PaymentDate']).dt.strftime('%m/%d/%Y')

    filename = filename.split('.')[0] + '_edits' + '.xlsx'
    try:
        df.to_excel(filename, engine='xlsxwriter', index=False)
    except PermissionError:
        filename = 'cash_edits1.xlsx'
        df.to_excel(filename, engine='xlsxwriter', index=False)

    return df

def get_filename(filename):
    '''
    Generate the filanem for the cash receipt journal based off the date of posting
    '''
    df = clean_data(filename)
    return '{} cash_receipts.xlsx'.format(df['Date'][0].replace('/','_'))

def fix_accounts(escrow, df):
    """Fix Account numbers for specific companies before creating sheet."""
    if escrow == 146:
        df.loc[(df.Account == '96021'), 'Account'] = '96024'
    elif escrow == 219:
        df.loc[(df.Account == '96023'), 'Account'] = '96020'
    elif escrow in [15, 115, 219]:
        df.loc[(df.Account == '43502'), 'Account'] = '43501'
    return df

def report_data(escrow, df):
    '''
    return the joined journal, locate the indices to add the formula into the excel worksheet that sums up the closing entries
    '''
    frame1 = report(escrow, df)
    frame2 = totals(escrow, df)
    journal = pd.concat([frame1, frame2], ignore_index=True)
    journal = fix_accounts(escrow, journal)
    s = journal[journal.Type == 'Bank Account'].index.values[0] + 3
    e = journal[journal.Account == '99998'].index.values[0] + 1
    f = '=SUM(I{}:I{})'.format(s,e)
    sheet = ACCOUNTS[escrow]['sheet']
    return journal, ACCOUNTS[escrow]['sheet'], f, e

def create_spreadsheet(filename, arr):
    '''
    Generate the excel workbook.
    '''
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        for i in range(len(arr)):
            arr[i][0].to_excel(writer, sheet_name=arr[i][1], index=False)
            workbook = writer.book
            worksheet = writer.sheets[arr[i][1]]
            num_format = workbook.add_format({'num_format':'##0.00'})
            worksheet.set_column(0, 2, 12)
            worksheet.set_column(3, 3, 3)
            worksheet.set_column(4, 5, 7)
            worksheet.set_column(6, 6, 15)
            worksheet.set_column(7, 7, 21)
            worksheet.set_column('I:J', 9, cell_format=num_format)
            worksheet.write_formula('J{}'.format(arr[i][-1] + 1), arr[i][-2])

def check_sheet(cash, fees):
    """
    Scans the worksheet and returns errors if any. if master is set to True it accounts that each file number in the master
    worksheet has been accounted for and the proper amount has been recorded.
    """
    errors = []
    fees = pd.read_excel(fees)
    cash = pd.read_excel(cash)

    fee_files = set(fees['File'])
    cash_files = set(cash['File Number'])
    files = fee_files.union(cash_files)

    if cash.OrderCategory.isna().sum() > 0:
        no_order_category = cash[cash.OrderCategory.isna()]['File Number'].unique()
        for x in no_order_category:
            errors.append('{} has no order category'.format(x))

    for file in files:
        posted = cash[cash['File Number'] == file]['Invoice Line Total'].sum()
        transferred = fees[fees['File'] == file]['Amount'].sum()
        if not round(posted, 2) == round(transferred, 2):
            errors.append(f'{file} is off {round(transferred - posted, 2)}')
    return errors

def get_revisions(cash):
    revisions = []
    cash = pd.read_excel(cash)

    second_invoices = cash[(cash.SortField == 2) & (cash.AcctCode.isin(['40000','40002']))]
    for x in second_invoices['File Number'].unique():
        revisions.append('{} is a second invoice with a premium, review'.format(x))
    
    return revisions

def check_balances(filename):
    '''
    Check the accounts to make sure the totals balance.
    '''
    xl = pd.ExcelFile(filename)
    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        rev_accts = [CLOSING[x]['revenue'] for x in CLOSING.keys()]
        x = round(df[df.Account.isin(rev_accts)]['Debits'].sum(), 2)
        y = round(df[df.Type == 'Bank Account']['Debits'].sum(), 2)
        if not x == y:
            print(sheet, 'Revenue Accounts:', x, 'Bank Account:', y)

def generate_journal(cash_file, save_file, fee_sheet):
    '''
    Main function of the program, first check sheet for errors and return any errors that wont get accounted for in the journal,
    If there are revisions create the sheet anyway but print out any revisions that need to be made,
    group the data by EscrowBank then run it through the other functions to generate the worksheet.
    '''
    errors = check_sheet(cash_file, fee_sheet)
    if errors:
        for error in errors:
            print(error)
        return None

    revisions = get_revisions(cash_file)
    for revision in revisions:
        print(revision)

    df = clean_data(cash_file)

    arr = list()
    for escrow, frame in df.groupby('EscrowBank'):
        arr.append(report_data(escrow, frame))

    arr = sorted(arr, key=lambda x: x[1])
    
    create_spreadsheet(save_file, arr)

    check_balances(save_file)