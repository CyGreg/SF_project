import pandas as pd
import gurobipy as gp
from gurobipy import GRB
from openpyxl import load_workbook


dane = pd.read_excel('Test_param.xlsm',sheet_name='Arkusz1',skiprows=1)
print(len(dane))

parametry_in = {}

for i in range(len(dane)):
    parametry_in[dane['Nazwa'][i]] = dane['Wartosc'][i]

print(parametry_in['P1'])

parametry_out = load_workbook('Test_param.xlsx')
aktywny = parametry_out.active

for i in range(6,9):
    aktywny['E'+str(i)] = 2
"""


aktywny['E6'] = 0
aktywny['E7'] = 0
aktywny['E8'] = 0
aktywny['E9'] = 0
"""

parametry_out.save('Test_param.xlsx')
