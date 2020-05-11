import pandas as pd
import math

filepath = r'C:\Users\karthi\Downloads\Template Linear Regression Sheet.xlsx'
results_file = r'C:\Users\karthi\Downloads\results.xlsx'

range_low = 3139.2
range_hi = 3836.8
print('Range : {} to {}'.format(range_low, range_hi))
trade_price_col_start_offset = 13

hiLo_sheets = ['HH', 'LL', 'HL', 'LH']
results = {}

for sheet in hiLo_sheets:
    data = pd.read_excel(filepath, sheet_name=sheet)
    df = pd.DataFrame(data)

    all_cols = []
    col_index = trade_price_col_start_offset
    while True:
        col =  df.iloc[:, col_index]
        trade_val_list = [val for val in col if isinstance(val, (int, float)) and not math.isnan(val)]
        if not len(trade_val_list):
            break
        all_cols = all_cols + trade_val_list
        col_index = col_index + 9

    results[sheet] = []
    for val in all_cols:
        if range_low < val < range_hi:
            new_val = int(val * 100)
            lastdigit = 0
            if new_val%10 >= 5:
                lastdigit = 5
            val = float((int(new_val/10)*10 + lastdigit)/100)
            results[sheet].append(val)
    print('SHEET : {} => Total no of entries : {}'.format(sheet, len(results[sheet])))


'''
Write to file
'''
writer = pd.ExcelWriter(results_file, engine = 'openpyxl')

data = {}
max_len = max([len(results[sheet]) for sheet in hiLo_sheets])
for sheet in hiLo_sheets:
    data[sheet] = results[sheet]
    if len(results[sheet]) < max_len:
        dummy_elems = [math.nan for val in range(0,max_len-len(results[sheet]))]
        data[sheet] = data[sheet] + dummy_elems

df = pd.DataFrame(data)

df.to_excel(writer, sheet_name='results')
writer.save()
writer.close()
