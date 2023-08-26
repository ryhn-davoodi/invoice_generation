import pandas as pd
import glob
paths=glob.glob("invoices/*.xlsx")

for i in paths:
    df=pd.read_excel(i,sheet_name="Sheet 1")
    print(df)