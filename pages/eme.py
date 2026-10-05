import pandas as pd

file_path = r'C:\Users\SPM\Downloads\BPI\Template\TEMPLATE DB\ACTIVE 05.15.xlsx'
df = pd.read_excel(file_path)
print(df.head())