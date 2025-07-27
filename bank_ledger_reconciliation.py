import os
import pandas
import pathlib

# bas = bank account statement
# deb = day end balance

def extract_bas_deb(directory):
    for file_name in os.listdir(directory):
        if file_name.startswith("So phu Ngan hang") and (file_name.endswith("xls") or file_name.endswith("xlsx")):
            bas_file_path = os.path.join(directory, file_name)
            bas = pandas.read_excel(bas_file_path, sheet_name=None, dtype=str)["Sheet 1"]
            print(bas)
            bas["transaction_date_time"] = pandas.to_datetime(bas["Thời gian giao dịch"])
            bas["transaction_date"] = bas["transaction_date_time"].dt.date
            deb_idx = bas.groupby("transaction_date")["transaction_date_time"].idxmax()
            deb_per_date = bas.loc(deb_idx, ["transaction_date", "Số dư cuối"]).set_index("transaction_date")
            print(deb_per_date)
    return deb_per_date

def calculate_bas_deb(directory):
    return

def extract_evn_deb():
    return

def export_results():
    return

extract_bas_deb(pathlib.Path(__file__).parent.resolve())