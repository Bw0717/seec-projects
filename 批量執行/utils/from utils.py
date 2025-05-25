from utils.excel_reader import read_excel_and_process
file_path = r"C:\Users\14574\Desktop\自動匯入表單.xlsx"
data = read_excel_and_process(file_path)
for data_row in data:
    print(data_row)