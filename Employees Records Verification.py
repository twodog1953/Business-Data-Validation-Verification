# verifying employee records across different database and/or platforms
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows as dtr
import Levenshtein as lev


def init_cleaning(df, keep_cols, concat_cols, shrink_cols, cap_cols):
    # for cleaning the initially read dataframes
    # pick desired columns, and restructure the columns if necessary
    # input: df, [desired columns], [[columns that need concat], delimiter, new_col_names], [[cols that need shrinking], shrink digit], [cols that needs cap]
    data = df
    ccols, delimiter, new_col_names = concat_cols[0], concat_cols[1], concat_cols[2]
    for i in range(len(ccols)):
        data[new_col_names[i]] = data[ccols[i][0]].fillna("")
        for j in ccols[i][1:]:
            data[new_col_names[i]] = pd.Series(map(lambda x, y: str(x) + delimiter + str(y), data[new_col_names[i]], data[j].fillna("")))
            # data[new_col_names[i]] = data[new_col_names[i]].apply(lambda x: str(x) + delimiter + str(data[j]))
    if shrink_cols != "":
        scols, shrink_digits = shrink_cols[0], shrink_cols[1]
        for i in range(len(scols)):
            if shrink_digits[i] < 0:
                data[scols[i]] = data[scols[i]].apply(lambda x: x[shrink_digits[i]:])
            elif shrink_digits[i] > 0:
                data[scols[i]] = data[scols[i]].apply(lambda x: x[:shrink_digits[i]])

    for c2 in cap_cols:
        data[c2] = data[c2].apply(lambda x: str(x).upper())

    data = df[keep_cols]

    return data


def rename_merge(df1, df2, keycol1, keycol2):
    # a function to merge and rename in one stop
    # rename all cols for distinction, and merge based on keycols
    for c in df1.columns:
        df1.rename(columns={c: c + '1'}, inplace=True)
    for c in df2.columns:
        df2.rename(columns={c: c + '2'}, inplace=True)
    df2 = df2.rename(columns={keycol2 + '2': keycol1 + '1'})
    df = pd.merge(df1, df2, how="left", on=keycol1 + '1')
    return df



def fuzz_score(f_dic, id_col):
    # To generate fuzzy score for all desired str cols
    # format: {"output col name": [col1, col2]}
    out = pd.DataFrame({"Employee ID1": id_col})
    for s in f_dic:
        col1, col2 = f_dic[s][0], f_dic[s][1]
        print("col1 len: {0}, col2 len: {1}".format(len(col1), len(col2)))
        # out[s] = pd.Series(map(lambda x, y: lev.ratio(x, y), col1, col2))
        out[s] = col1.combine(col2, lambda x, y: lev.ratio(x, y))
    return out




if __name__ == "__main__":
    pd.set_option('display.max_columns', None)
    pd.options.mode.chained_assignment = None  # default='warn'
    print("Initialized! ")
    # import data files and read as pandas df
    data_doc1 = "emp_sage.xlsx"
    data_doc2 = "emp_win.xls"
    df1 = pd.read_excel("emp_sage.xlsx", sheet_name="Employee List", header=3).drop_duplicates("Employee ID")
    df2 = pd.read_excel("emp_win.xls", header=1).drop_duplicates("EmployeeNumber")

    # initial cleaning and renaming
    keep_col1 = ["Employee ID", "Name", "Address", "City", "State", "Zip"]
    keep_col2 = ["EmployeeNumber", "Name", "Address", "City", "State", "Zip"]
    data1 = init_cleaning(df1,
                        keep_col1,
                        [[["Address 1", "Address 2"], ["First Name", "Last Name"]], " ", ["Address", "Name"]],
                        [["Employee ID"], [-5]],
                        ["Name", "Address", "City", "State"])
    data2 = init_cleaning(df2,
                          keep_col2,
                          [[["Address1", "Address2"], ["FirstName", "LastName"]], " ", ["Address", "Name"]],
                          [["Zip"], [5]],
                          ["Name", "Address", "City", "State"])
    data2["Name"] = data2["Name"].apply(lambda x: x.replace("  ", " "))
    # change col names for better distinction
    # data2.rename(columns={}, inplace=True)

    # print(data1.head(10))
    # print('----------------------')
    # print(data2.head(10))
    df1 = ''
    df2 = ''

    # merge the two sets and get fuzzy score
    key1 = "Employee ID"
    key2 = "EmployeeNumber"
    data1[key1] = data1[key1].astype(pd.StringDtype())
    data2[key2] = data2[key2].astype(pd.StringDtype())

    # print(data1.describe())
    # print('-----')
    # print(data2.describe())

    # df = rename_merge(data1, data2, key1, key2).fillna('').drop_duplicates("Employee ID1")
    df = rename_merge(data1, data2, key1, key2)
    df = df.dropna(subset=["Name2"]).fillna('').drop_duplicates("Employee ID1")
    dff = df.copy()
    print(df.describe())
    print("-----")

    # format: {"output col name": [col1, col2]}
    fuzz_dic = {
        "Name_Score": [df["Name1"], df["Name2"]],
        "Address Score": [df["Address1"], df["Address2"]],
        "City Score": [df["City1"], df["City2"]],
        "Zip Score": [df["Zip1"], df["Zip2"]]
    }
    fuzz_df = fuzz_score(fuzz_dic, df["Employee ID1"])
    print(fuzz_df.describe())
    df = pd.merge(df, fuzz_df, on="Employee ID1", how="left")

    # output result to excel
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in dtr(df, index=False, header=True):
        ws.append(r)
    wb.save('check_out.xlsx')
    openpyxl.workbook.Workbook.close(wb)