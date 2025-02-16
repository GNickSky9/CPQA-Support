import sys
import argparse
import pandas as pd
import os
import glob
import magic
import openpyxl


'''
The data being read in will be every weekly spreadsheet we have at the moment. We will combine the data by test name
and sort by date while keeping track of the result value.
'''
if __name__ == "__main__":
	if len(sys.argv) > 1:
		print("We have args")
	else:
		print("We don't have args")
	
	excel_files = glob.glob("*.xlsx") + glob.glob("*.xls")
	
	if not excel_files:
		print("No Excel files found in the directory.")
	else:
		for file in excel_files:
			
			print(f"Reading {file}...")
			print("File Type\t" + magic.from_file(file, mime=True))
			if file.endswith(".xlsx"):
				df = pd.read_excel(file, engine="openpyxl", header=0, skiprows=4).dropna(how='all')
				df_clean = df.dropna().reset_index(drop=True)
			else:
				df = pd.read_excel(file, engine="xlrd")
	
	columns_list = df_clean.columns.tolist()
	pd.set_option('display.max_rows', None)
	print(df_clean[['ReceiveDate', 'OrderName', 'RcvResMn']])
