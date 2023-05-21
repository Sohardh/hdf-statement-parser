import csv
import os
import re
import time

import pandas as pd

STATEMENT_XL_FILES = [
  '/Users/sohardhchobera/Downloads/50100243021840_1684655192040.xls',
  '/Users/sohardhchobera/Downloads/115528583_1684655208443.xls']


def excel_to_json(xlfile):
  dataframe1 = pd.read_excel(xlfile)
  df = []
  for index, row in dataframe1.iterrows():
    if re.search('^[0-9]{1,2}\\/[0-9]{1,2}\\/[0-9]{2}$', str(row[0]).strip()):
      dict = {
        'date': str(row[0]), "description": str(row[1]),
        "debit": (str(row[4])) if str(row[4]) != 'nan' else '',
        "credit": (str(row[5])) if str(row[5]) != 'nan' else '',
        "ref_no": str(row[2])
      }
      df.append(dict)
  return df


def save_csv(transactions):
  keys = ['date', 'description', 'debit', 'credit', 'ref_no',
          'opposite_account', 'asset_account', 'tag', 'category']
  if not os.path.exists('output'):
    os.makedirs('output')
  csv_file = os.path.join('output',
                          'transactions_' + str(int(time.time())) + '.csv')
  with open(csv_file, 'w',
            newline='') as output_file:
    dict_writer = csv.DictWriter(output_file, keys)
    dict_writer.writeheader()
    dict_writer.writerows(transactions)


def transaction_analyser(transactions):
  for transaction in transactions:
    parse_upi(transaction)
    parse_hdfc(transaction)
    parse_salary(transaction)
    parse_misc(transaction)
    cleanup(transaction)
  save_csv(transactions)


def cleanup(transaction):
  for k, v in transaction.items():
    transaction[k] = v.replace(',', ';')


def parse_misc(transaction):
  desc = transaction['description']
  if desc.startswith('POS'):
    split = desc.split(' ')[2:]
    op_acc_name = ' '.join(split)
    transaction['opposite_account'] = op_acc_name


def parse_hdfc(transaction):
  desc = transaction['description']
  if desc.startswith('FD THROUGH') or desc.startswith(
      'PRIN AND INT AUTO_REDEEM'):
    transaction['opposite_account'] = 'HDFC FD'
    transaction['tag'] = 'bank'

  if desc.find('RELOAD FOREX CARD') != -1:
    transaction['opposite_account'] = 'HDFC Forex Card'
    transaction['tag'] = 'forex_card wallet bank'

  if desc.startswith('CC') and desc.find('AUTOPAY') != -1:
    transaction['opposite_account'] = 'HDFC Bank'
    transaction['tag'] = 'credit_card bank'
    transaction['category'] = 'bill'
  if desc.startswith('CREDIT INTEREST CAPITALISED'):
    transaction['opposite_account'] = 'HDFC Bank'
    transaction['tag'] = 'interest bank'
  if desc.startswith('INT. AUTO_REDEMPTION'):
    transaction['opposite_account'] = 'HDFC Bank'
    transaction['tag'] = 'interest bank'


def parse_salary(transaction):
  if transaction['description'].find('BLUEOPTIMA SAAS') != -1:
    transaction['opposite_account'] = 'Blueoptima'
    transaction['tag'] = 'salary'
    transaction['category'] = 'salary'


def parse_upi(transaction):
  if not transaction['description'].startswith('UPI-'):
    return
  transaction['tag'] = 'UPI'
  desc = transaction['description']
  op_acc_name = desc.split('-')[1]
  if op_acc_name == 'ADD MONEY TO WALLET':
    transaction['opposite_account'] = 'PAYTM'
    transaction['tag'] = transaction['tag'] + ' wallet'
  else:
    transaction['opposite_account'] = op_acc_name

  if op_acc_name.upper().startswith('ZOMATO') or (
      op_acc_name.upper().startswith('SWIGGY') and op_acc_name.upper().find(
      'SWIGGYINSTAMART') == -1):
    transaction['category'] = 'food'

  if op_acc_name.upper().startswith(
      'BLINKIT') or op_acc_name.upper().startswith('SWIGGYINSTAMART') \
      or op_acc_name.upper().startswith('GROFERS'):
    transaction['category'] = 'grocery'

  if op_acc_name.upper().startswith(
      'JIOMOBILITY') or op_acc_name.upper().startswith('PAYTM RECHARGE'):
    transaction['category'] = 'mobile_recharge'


# opposite_account
# tag
# category

def main():
  transactions = []
  for xl_file in STATEMENT_XL_FILES:
    transactions += excel_to_json(xl_file)
  transaction_analyser(transactions)


if __name__ == '__main__':
  main()
