# verifying employee records across different database and/or platforms
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows as dtr

def init_cleaning(df, keep_cols, concat_cols, shrink_cols, cap_cols):
    # todo for cleaning the initially read dataframes
    # pick desired columns, and restructure the columns if necessary
    # input: df, [desired columns], [[columns that need concat], delimiter, new_col_names], [[cols that need shrinking], shrink digit], [cols that needs cap]
    data = df[keep_cols]

    ccols, delimiter, new_col_names = concat_cols[0], concat_cols[1], concat_cols[2]
    for i in range(len(ccols)):
        data[new_col_names[i]] = data[ccols[i]].apply(lambda row: ''.join(str(row)), axis=1)

    scols, shrink_digits = shrink_cols[0], shrink_cols[1]
    for i in range(len(scols)):
        if shrink_digits[i] < 0:
            data[scols[i]] = data[scols[i]].apply(lambda x: x[shrink_digits[i]:])
        elif shrink_digits[i] > 0:
            data[scols[i]] = data[scols[i]].apply(lambda x: x[:shrink_digits[i]])

    for c2 in cap_cols:
        data[c2] = data[c2].apply(lambda x: str(x).upper())

    return data


if __name__ == "__main__":
    print("Initialized! ")
    # import data files and read as pandas df
    data_doc1 = "emp_sage.xlsx"
    data_doc2 = "emp_win.xls"
    df1 = pd.read_excel("emp_sage.xlsx", sheet_name="Employee List", header=3)
    df2 = pd.read_excel("emp_win.xls", header=1)
    data1 = init_cleaning(df1,
                        ["Employee ID", "First Name", "Last Name", "Address 1", "Address 2", "City", "State", "Zip"],
                        [[["Address 1", "Address 2"], ["First Name", "Last Name"]], " ", ["Address", "Name"]],
                        [["Employee ID"], [-5]],
                        ["Name", "Address", "City", "State"])
    print(data1["Employee ID"].head(10))
