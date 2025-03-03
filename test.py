import pandas as pd

data = pd.read_excel("customer_list.xlsx", sheet_name="Main", dtype=object)

for row in data.itertuples():
    print(row)


    