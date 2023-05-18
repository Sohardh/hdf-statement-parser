import json
import re
from os import path
import pandas as pd

ERROR = []


def excel_to_json(xlfile):
  dataframe1 = pd.read_excel(xlfile)
  start_index = 0
  df = []
  for index, row in dataframe1.iterrows():
    # print(row[0])
    if re.search('^[0-9]{1,2}\\/[0-9]{1,2}\\/[0-9]{2}$', str(row[0]).strip()):
      dict = {
        'date': str(row[0]), "description": str(row[1]),
        "debit": str(row[3]),
        "credit": str(row[4]), "ref_no": str(row[5])
      }
      df.append(dict)
  return json.dumps(df, indent=2)


def SetTransactionDetail(newtran, tran, merchant, mode):
  if tran['Debit Amount'] > 0:
    newtran['Type'] = "Debit"
    newtran['Amount'] = tran['Debit Amount']
  else:
    newtran['Type'] = "Credit"
    newtran['Amount'] = tran['Credit Amount']
  newtran['Merchant'] = merchant
  newtran['Mode'] = mode


def main():
  STATEMENT_XL_FILE = '/Users/sohardhchobera/Downloads/50100243021840_1684433436965.xls'
  if path.exists(STATEMENT_XL_FILE):
    csvfile = open(STATEMENT_XL_FILE, 'r')
  else:
    ERROR.append("Error - Statement file not found")
    return ERROR
  transactions = excel_to_json(csvfile)
  return ERROR


if __name__ == '__main__':
  main()
