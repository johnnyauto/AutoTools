import pandas as pd
from openpyxl import load_workbook

# 讀取Excel檔案
excel_file = 'VIU_TR05_信號定義_V04_20240201_source.xlsx'
sheet_name = 'CAN01_Matrix'
df = pd.read_excel(excel_file)


# 選取指定欄位
# selected_columns = ['Transmitter', 'Receiver', 'Signal Name', 'Message Name', 'Message ID', 'Message Type', 'Signal Type', 'period\n(ms)']
# selected_df = df[selected_columns]
'''
# 開啟CAN.dbc檔案
dbc_file = open('path/to/your/CAN.dbc', 'w')

# 逐行寫入CAN.dbc檔案
#for index, row in selected_df.iterrows():
for index, row in df.iterrows():
    sig_name = row['SigName']
    msg = row['Msg']
    msg_id = row['MsgID']
    
    # 格式化輸出並寫入檔案
    line = f'B0_ {msg_id} {msg}\n SG_ {sig_name}\n'
    dbc_file.write(line)

# 關閉CAN.dbc檔案
dbc_file.close()
'''

'''output_seg01'''
# VERSION
output_seg01 = 'VERSION ""\n\n\n'

# NS_
output_seg01 += 'NS_ :'
output_seg01 += '    NS_DESC_\n    CM_\n    BA_DEF_\n    BA_\n    VAL_\n'
output_seg01 += '	CAT_DEF_\n    CAT_\n	FILTER\n	BA_DEF_DEF_\n    EV_DATA_\n'
output_seg01 += '    ENVVAR_DATA_\n    SGTYPE_\n    SGTYPE_VAL_\n    BA_DEF_SGTYPE_\n'
output_seg01 += '    BA_SGTYPE_\n    SIG_TYPE_REF_\n    VAL_TABLE_\n    SIG_GROUP_\n'
output_seg01 += '    SIG_VALTYPE_\n    SIGTYPE_VALTYPE_\n    BO_TX_BU_\n    BA_DEF_REL_\n'
output_seg01 += '    BA_REL_\n    BA_DEF_DEF_REL_\n    BU_SG_REL_\n    BU_EV_REL_\n'
output_seg01 += '    BU_BO_REL_\n    SG_MUL_VAL_\n\n'

# BS_
output_seg01 += 'BS_\n\n'

'''output_seg02'''
# BU_
workbook = load_workbook(excel_file, data_only=True)
worksheet = workbook[sheet_name]

Transmiter = []
Tx_index = 4    # column index of 'Transmiter'
Msg_index = 10  # column index of 'Message Name'

# for row_index in worksheet.iter_rows(values_only=True):
for row_index in range(2, worksheet.max_row+1):
    Msg_value = worksheet.cell(row=row_index, column=Msg_index).value
    Msg_value = worksheet.cell(row=row_index, column=Msg_index).font.strike
    Tx_value = worksheet.cell(row=row_index, column=Tx_index).value
    Tx_strike = worksheet.cell(row=row_index, column=Tx_index).font.strike

    # if not empty and not strikethrough format
    if Tx_value and not Tx_strike:
        Transmiter.append(Tx_value)

Transmiter = sorted(set(Transmiter))
output_seg02 = 'BU_: '
for value in Transmiter:
    output_seg02 += f'{value} '
output_seg02 += '\n\n'


'''output_seg03'''
# BO_ 




# create a DBC file
output_text = output_seg01 + output_seg02
with open('MyDBC.dbc', 'w', encoding='utf-8') as f:
    f.write(output_text)

print("MyDBC.dbc is generated.")
