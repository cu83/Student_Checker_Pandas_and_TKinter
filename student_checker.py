import pandas as pd

file = 'blocklists.xlsx'

name = str(input('Name: '))
number = str(input('Number: '))
sheets = None

sheet1 = 'BSEE3A'
sheet2 = 'BSEE3B'
sheet3 = 'BSEE3C'
sheet4 = 'BSEE3D'

df = pd.read_excel(file, sheet_name=sheet1)
df2 = pd.read_excel(file, sheet_name=sheet2)
df3 = pd.read_excel(file, sheet_name=sheet3)
df4 = pd.read_excel(file, sheet_name=sheet4)

if name and number in df.values:
    try:
        print(sheet1)
    except ValueError:
        print('Value does not exist')

elif name and number in df2.values:
    try:
        print(sheet2)
    except ValueError:
        print('Value does not exist')

elif name and number in df3.values:
    try:
        print(sheet3)
    except ValueError:
        print('Value does not exist')

elif name and number in df4.values:
    try:
        print(sheet4)
    except ValueError:
        print('Value does not exist')

else:
    print('Match Not Found')
    

