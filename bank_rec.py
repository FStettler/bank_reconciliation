from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from datetime import date, datetime, timedelta
import calendar
import pandas as pd
import glob
import numpy as np
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font

def read_static():

    static = open('static.txt','r')
    
    lineas= static.readlines()

    month = lineas[0][ lineas[0].strip().find(':') + 1 :].strip()
    year = lineas[1][ lineas[1].strip().find(':') + 1 :].strip()
    opening_balance = float(lineas[2][ lineas[2].strip().find(':') + 1 :].strip())

    static.close()

    return month, year, opening_balance

def login():

    url = "https:///admin/login.php"
    driver = webdriver.Chrome(ChromeDriverManager().install())
    driver.get(url)
    driver.find_element(By.NAME,"username").send_keys('')
    driver.find_element(By.NAME,"pass").send_keys('')
    driver.find_element(By.CSS_SELECTOR,"input[value='Login']").click()
    return driver

def book_bank_charges(driver):

    df = pd.read_excel('bank_rec.xlsx',sheet_name='System')
    bank_charges = float(df.columns[-1])
    bank_charges = round(bank_charges ,2)
    check = (bank_charges * 100 / 15)

    if check % 1 == 0 and bank_charges > 0:

        url = 'https:///admin/bank/user_line_edit.php'
        driver.get(url)

        driver.find_element(By.NAME,'transaction_ref').send_keys('Bank charges SO')
        
        nom_code = Select(driver.find_element(By.NAME, "transaction_type"))
        nom_code.select_by_value('6')

        today = datetime.today()
        
        yesterday = today - timedelta(days=1)
        yesterday = yesterday.strftime("%d-%m-%Y")

        transaction_date = driver.find_element(By.NAME,'transaction_date')
        transaction_date.clear()
        transaction_date.send_keys(yesterday)

        driver.find_element(By.NAME,'amount_out').send_keys(bank_charges)

        driver.find_element(By.NAME,'nom_code').send_keys('371')

        driver.find_element(By.XPATH, "//input[@value='Create New']").click()

def get_system_bookings(driver, month, year):


    last_day = calendar.monthrange(date(int(year), int(month), 1).year, date(int(year), int(month), 1).month)[1]

    url = "https:///admin/bank/search.php?id=&t_type=&acc_id=7&t_ref=&pay_ref=&sd=01-" + month + "-" + year + "&ed=" + str(last_day) + "-" + month + "-" + year + "&amount=&ul=&reconciled="
    driver.get(url)
    tbody = driver.find_element(By.XPATH,"//th[text()='ID']/../..")
    rows = tbody.find_elements(By.TAG_NAME,"tr")

    id = []
    line_date = []
    ref = []
    outs = []
    ins = []

    for row in rows:

        tds = row.find_elements(By.TAG_NAME,"td")
        
        if len(tds) != 10:
            continue

        try:
            if tds[0].text.strip() == '':
                continue
        except IndexError:
            continue

        id.append(tds[0].find_element(By.TAG_NAME,"a").text)
        line_date.append(
            date(
            day = int(tds[3].text[:2]),
            month = int(tds[3].text[3:5]),
            year = int(tds[3].text[6:10])
            ))
        ref.append(tds[5].text)
        outs.append(tds[7].text)
        ins.append(tds[8].text)

    data = {'id':id, 'date':line_date, 'ref':ref, 'out':outs, 'in': ins}
    system_data = pd.DataFrame(data)

    system_data['out'] = system_data['out'].str.replace(' ','0')
    system_data['out'] = system_data['out'].str.replace(',','')

    system_data['in'] = system_data['in'].str.replace(' ','0')
    system_data['in'] = system_data['in'].str.replace(',','')

    system_data['id'] = pd.to_numeric(system_data['id'])
    system_data['out'] = pd.to_numeric(system_data['out'], downcast='float')
    system_data['in'] = pd.to_numeric(system_data['in'], downcast='float')

    system_data['out'] = np.round(system_data['out'],decimals=2)
    system_data['in'] = np.round(system_data['in'],decimals=2)

    system_data['net'] = system_data['in'] - system_data['out']
    system_data['net'] = np.round(system_data['net'],decimals=2)

    driver.quit()

    return system_data

def merges_system_data(df):

    old_data = pd.read_excel('bank_rec.xlsx',sheet_name='System')
    
    df_desc = df[ df['id'].isin(old_data['id'].to_list())]
    df_desc = df_desc[['id', 'ref']]

    old_data = pd.merge(old_data, df_desc, how='left', on='id')

    old_data.rename(columns={'ref_y':'ref'}, inplace=True)
    old_data.drop(['ref_x'], axis=1, inplace=True)
    
    
    df = df[ ~df['id'].isin(old_data['id'].to_list())]

    new_data = pd.concat([old_data[['id','date','ref', 'out', 'in', 'net', 'status', 'obs']],df])

    new_data = new_data.reset_index(drop=False)
    new_data = new_data.drop('index', axis=1)

    new_data.loc[ new_data['ref'] == 'Bank charges SO' , 'status'] = 'OK'

    return new_data

def get_bank_info():

    my_csv_files = glob.glob('*.csv')

    bank_fresh_data = pd.read_csv(my_csv_files[0])
    bank_fresh_data = bank_fresh_data.fillna(0)
    bank_fresh_data['net'] = bank_fresh_data['Credit'] - bank_fresh_data['Debit']
    bank_fresh_data['Date'] = pd.to_datetime(bank_fresh_data['Date'], format="%d/%m/%Y")

    balances = bank_fresh_data[['Date','Balance']]
    balances = balances[balances['Balance'] != 0]

    bank_fresh_data = bank_fresh_data.drop(['Balance'],axis=1)

    date_max = bank_fresh_data['Date'].max()

    bank_fresh_data = bank_fresh_data[bank_fresh_data['Date'] < date_max]

    return bank_fresh_data, balances

def concat_bank_data(df):

    old_data = pd.read_excel('Bank_rec.xlsx',sheet_name='Bank', parse_dates=True)

    old_data = old_data[['Date','Details','Debit','Credit','net','status','obs','ANU','Comm']]

    date_max = old_data['Date'].max()

    df = df[df['Date'] > date_max]

    new_data = pd.concat([old_data, df])

    new_data = new_data.reset_index(drop=False)
    new_data = new_data.drop('index', axis=1)

    return new_data

def check_closing_balance(bank_data, balances, opening_balance):
    
    closing_date = bank_data['Date'].max()

    closing_balance = float(balances[balances['Date'] == closing_date]['Balance'])

    transactions_net = bank_data['net'].sum()

    closing_balance_2 = round(transactions_net + opening_balance,2)

    if closing_balance == closing_balance_2:
        return True
    else:
        return False
 
def check_ok_status(bank_data, system_data):
    
    bank_net_ok = bank_data[bank_data['status'] == 'OK']['net'].sum().round(2)
    system_net_ok = system_data[system_data['status'] == 'OK']['net'].sum().round(2)

    difference = system_net_ok - bank_net_ok

    check = round((difference * 100 / 15) , 2)

    bank_charges = round(check / 100 * 15 ,2)

    if check % 1 == 0:
        return True, bank_charges
    else:
        return False, None

def match(bank_data, system_data):
    
    for i,r in bank_data.iterrows():

        bank_status = r['status']

        if bank_status == 'OK':
            continue

        bank_date = r['Date']
        bank_desc = r['Details'].strip().replace(' ','')
        bank_net = r['net']

        if bank_net > 0:
            continue

        for index,row in system_data.iterrows():

            system_status = row['status']

            if system_status == 'OK':
                continue

            system_date = row['date']
            system_desc = row['ref'].strip().replace(' ','')
            system_net = np.round(row['net'],decimals=2)

            if system_net > 0:
                continue

            if system_desc in ['xxx','yyy','zzz']:
                system_net -= 0.15

            if bank_desc == system_desc and bank_net == system_net:

                bank_data.at[i, 'status'] = 'OK'
                system_data.at[index, 'status'] = 'OK'
                break
    
    return bank_data, system_data

def to_excel(bank_data, system_data):
    workbook = Workbook()

    system = workbook.active
    system.title = 'System'
    bank = workbook.create_sheet("Bank")

    for r in dataframe_to_rows(bank_data, index=False, header=True):
        bank.append(r)

    for r in dataframe_to_rows(system_data, index=False, header=True):
        system.append(r)

    bank['J1'] = '=sumif(System!G:G, "OK", System!F:F)-sumif(F:F, "OK", E:E)'
    bank.freeze_panes = 'A2'
    bank.sheet_view.showGridLines = False

    system['I1'] = '=sumif(G:G, "OK", F:F)-sumif(Bank!F:F, "OK", Bank!E:E)'
    system.freeze_panes = 'A2'
    system.sheet_view.showGridLines = False

    date_format_code = 'DD/MM/YYYY'
    for cell in system['B']:
        cell.number_format = date_format_code

    for cell in bank['A']:
        cell.number_format = date_format_code

    cols = ['A','B','C','D','E','F','G','H']   
    for c in cols:
        system[c+'1'].alignment = Alignment(horizontal='center')
        system[c+'1'].font = Font(bold=True)

    cols = ['A','B','C','D','E','F','G','H', 'I']   
    for c in cols:
        bank[c+'1'].alignment = Alignment(horizontal='center')
        bank[c+'1'].font = Font(bold=True)

    final_row = system_data.shape[0]
    filter_range = f'A1:H{final_row}'
    system.auto_filter.ref = filter_range

    final_row = bank_data.shape[0]
    filter_range = f'A1:I{final_row}'
    bank.auto_filter.ref = filter_range

    workbook.save("bank_rec.xlsx")



if __name__ == '__main__':

    month, year, opening_balance = read_static()

    driver = login()

    book_bank_charges(driver)

    system_data = get_system_bookings(driver, month, year)
    system_data = merges_system_data(system_data)

    bank_data, balances = get_bank_info()
    bank_data = concat_bank_data(bank_data)

    check_cb = check_closing_balance(bank_data, balances, opening_balance)
    check_ok, bank_charges = check_ok_status(bank_data, system_data)

    print('Check closing balance: ',check_cb)
    print('Check OKs: ',check_ok)
    print('Check bank charges: ',bank_charges)

    if check_cb and check_ok:
        
        bank_data, system_data = match(bank_data, system_data)
        
        to_excel(bank_data, system_data)



