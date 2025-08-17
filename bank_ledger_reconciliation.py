import os
import pandas
import pathlib
from openpyxl import load_workbook
import sys

# bas = bank account statement
# deb = day end balance
# rt = result template

def extract_bas_deb_by_time(directory: str) -> dict:
    for file_name in os.listdir(directory):
        if file_name.startswith("So phu Ngan hang") and (file_name.endswith("xls") or file_name.endswith("xlsx")):
            bas_file_path = os.path.join(directory, file_name)
            print(bas_file_path)
            bas = pandas.read_excel(bas_file_path, dtype=str)
            print(bas)
            bas["transaction_date_time"] = pandas.to_datetime(bas["Thời gian giao dịch"])
            bas["transaction_date"] = bas["transaction_date_time"].dt.date
            deb_idx = bas.groupby("transaction_date")["transaction_date_time"].idxmax()
            print(deb_idx)
            print(bas.groupby("transaction_date")["transaction_date_time"].max())
            deb_per_date = bas.loc[deb_idx, ["transaction_date", "Số dư cuối"]]
            print(deb_per_date)
    return dict(zip(deb_per_date.iloc[:, 0], deb_per_date.iloc[:, 1]))

def extract_bas_deb_by_order(directory: str) -> dict:
    for file_name in os.listdir(directory):
        if file_name.startswith("So phu Ngan hang") and (file_name.endswith("xls") or file_name.endswith("xlsx")):
            bas_file_path = os.path.join(directory, file_name)
            print(bas_file_path)
            bas = pandas.read_excel(bas_file_path, dtype=str)
            bas["transaction_date_time"] = pandas.pandas.to_datetime(bas["Thời gian giao dịch"]).dt.date
            deduplicated_transaction_dates =  bas["transaction_date_time"].drop_duplicates()
            deb_per_date = {}
            if (deduplicated_transaction_dates.is_monotonic_decreasing):
                deb_per_date = bas.groupby("transaction_date_time", as_index=False).first()[["transaction_date_time", "Số dư cuối"]]
            elif (deduplicated_transaction_dates.is_monotonic_increasing):
                deb_per_date = bas.groupby("transaction_date_time", as_index=False).last()[["transaction_date_time", "Số dư cuối"]]
            else:
                return None
            print(deb_per_date)
    return dict(zip(deb_per_date.iloc[:, 0], deb_per_date.iloc[:, 1]))

def calculate_bas_deb(directory):
    return

def extract_evn_deb(directory: str) -> dict:
    for file_name in os.listdir(directory):
        if file_name.startswith("So TGNH") and (file_name.endswith("xls") or file_name.endswith("xlsx")):
            evn_file_path = os.path.join(directory, file_name)
            print(evn_file_path)
            evn = pandas.read_excel(evn_file_path, header=None, dtype=str, skiprows=17, skipfooter=8)
            deb_per_date = evn.loc[evn[7].notna()].iloc[:, [4, 7]]
            deb_per_date[4] = pandas.to_datetime(deb_per_date[4].str[-10:], format="%d/%m/%Y").dt.date
            print(deb_per_date)
    return dict(zip(deb_per_date.iloc[:, 0], deb_per_date.iloc[:, 1]))

def export_results(bas_deb: dict, evn_deb: dict):
    rt_file_name = "KQ doi soat So phu NH - EVN_CM_009.xlsx"

    rt_wb = load_workbook(rt_file_name)
    rt_ws = rt_wb.active

    # print(bas_deb)
    # print(evn_deb)
    
    results = []
    for idx, (key, value) in enumerate(bas_deb.items()):
        results.append(dict(index=idx+1, transaction_date=key, bas_deb_res=value, evn_deb_res=evn_deb[key], diff=(int(evn_deb[key].replace(" ", "")) - int(value.replace(",", "")))))
    print(results)

    start_row = 7
    for i, row in enumerate(results, start=start_row):
        for j, key in enumerate(row, start=1):
            rt_ws.cell(row=i, column=j, value=row[key])
    
    rt_wb.save("KQ doi soat So phu NH - EVN_CM_009_test.xlsx")

    return results

# exe_path = sys.argv[0]
# exe_dir = os.path.dirname(exe_path) # pathlib.Path(__file__).parent.resolve()
# bas_deb = extract_bas_deb_by_time(exe_dir)
# evn_deb = extract_evn_deb(exe_dir)
# export_results(bas_deb, evn_deb)

bas_deb = extract_bas_deb_by_order(pathlib.Path(__file__).parent.resolve())
evn_deb = extract_evn_deb(pathlib.Path(__file__).parent.resolve())
export_results(bas_deb, evn_deb)