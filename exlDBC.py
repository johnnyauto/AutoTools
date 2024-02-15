import pandas as pd
import openpyxl

'''Process Data (remove Empty and Strikethrough format data)'''
# 讀取Excel檔案
excel_file = 'VIU_TR05_信號定義_V04_20240201_source.xlsx'
sheet_name = 'CAN01_Matrix'
workbook = openpyxl.load_workbook(excel_file, data_only=True)
worksheet = workbook[sheet_name]

pData = []      # processed Data
Tx_index = 4    # column index of 'Transmiter'
Msg_index = 10  # column index of 'Message Name'

# for row_index in worksheet.iter_rows(values_only=True):
for row_index in range(2, worksheet.max_row+1):
    Msg_value = worksheet.cell(row=row_index, column=Msg_index).value
    Msg_strike = worksheet.cell(row=row_index, column=Msg_index).font.strike
    #Tx_value = worksheet.cell(row=row_index, column=Tx_index).value
    #Tx_strike = worksheet.cell(row=row_index, column=Tx_index).font.strike

    # if not empty and not strikethrough format
    if Msg_value and not Msg_strike:
        # 如果 Msg 欄位不為空白且未被格式為strike，將整列資料加入 pData 中
        pData.append([worksheet.cell(row=row_index, column=col).value for col in range(1, worksheet.max_column + 1)])

# 取得欄位名稱
columns = [worksheet.cell(row=1, column=col).value for col in range(1, worksheet.max_column + 1)]

# 將資料轉換為 DataFrame
df = pd.DataFrame(pData, columns=columns)
# 將'Message ID'欄位的資料轉換為10進制資料
df['Message ID'] = df['Message ID'].apply(lambda x: int(x, 16))

#selected_columns = ['Transmitter', 'Receiver', 'Signal Name', 'Message Name', 'Message ID', 'Message Type', 'Signal Type', 'DLC']
#selected_df = df[selected_columns]

'''output_seg01'''
# VERSION
output_seg01 = 'VERSION ""\n\n\n'

# NS_
output_seg01 += 'NS_ :\n'
output_seg01 += '    NS_DESC_\n    CM_\n    BA_DEF_\n    BA_\n    VAL_\n'
output_seg01 += '	CAT_DEF_\n    CAT_\n	FILTER\n	BA_DEF_DEF_\n    EV_DATA_\n'
output_seg01 += '    ENVVAR_DATA_\n    SGTYPE_\n    SGTYPE_VAL_\n    BA_DEF_SGTYPE_\n'
output_seg01 += '    BA_SGTYPE_\n    SIG_TYPE_REF_\n    VAL_TABLE_\n    SIG_GROUP_\n'
output_seg01 += '    SIG_VALTYPE_\n    SIGTYPE_VALTYPE_\n    BO_TX_BU_\n    BA_DEF_REL_\n'
output_seg01 += '    BA_REL_\n    BA_DEF_DEF_REL_\n    BU_SG_REL_\n    BU_EV_REL_\n'
output_seg01 += '    BU_BO_REL_\n    SG_MUL_VAL_\n\n'

# BS_
output_seg01 += 'BS_:\n\n'


'''output_seg02'''
# BU_
output_seg02 = 'BU_: '

df_group = df.groupby('Transmitter')
Transmitter = list(df_group.groups.keys())
for value in Transmitter:
    output_seg02 += f'{value} '

output_seg02 += '\n\n'


'''output_seg03'''
# BO_ 
# 選取指定欄位
selected_columns = ['Transmitter', 'Message Name', 'Message ID', 'DLC']
selected_df = df[selected_columns]
selected_df_group = selected_df.groupby('Message Name')

df_group = df.groupby('Message Name')

output_seg03 =""
for group_lable, group_data in df_group:
    Transmitter = group_data['Transmitter'].iloc[0]
    MsgName = group_data['Message Name'].iloc[0]
    MsgID = group_data['Message ID'].iloc[0]
    DataLen = group_data['DLC'].iloc[0]
    output_seg03 += f'BO_ {MsgID} {MsgName}: {DataLen} {Transmitter}\n'

    for index in range(len(group_data)):
        signal_name = group_data['Signal Name'].iloc[index]
        signal_size = group_data['size(bit)'].iloc[index]
        factor = int(group_data['Factor'].iloc[index])
        offset = int(group_data['Offset'].iloc[index])
        minimum = int(group_data['P-Minimum'].iloc[index])
        maximum = int(group_data['P-Maximum'].iloc[index])
        

        ByteOrder = group_data['Byte Order'].iloc[index]
        if ByteOrder is 'Motorola':
            start_bit = group_data['Msb'].iloc[index]
            byte_order = '0'
        else:   # Intel
            start_bit = group_data['Lab'].iloc[index]
            byte_order = '1'

        DataType = group_data['Data Type'].iloc[index]
        if DataType is 'unsigned':
            value_type = '+'
        else:   # signed
            value_type = '-'
        
        unit = group_data['Unit'].iloc[index]
        if pd.isna(unit):
            unit = ''

        receiver = group_data['Receiver'].iloc[index]
        receiver = receiver.replace('\n',',')

        #SG_ {signal_name} {multiplexer_indicator} : {start_bit}|{signal_size}@{byte_order}{value_type} ({factor},{offset}) [{minimum}|{maximum}] "{unit}" {receiver}
        output_seg03 += f' SG_ {signal_name} : {start_bit}|{signal_size}@{byte_order}{value_type} ({factor},{offset}) [{minimum}|{maximum}] "{unit}" {receiver}\n'

        if index == len(group_data)-1:
            output_seg03 += '\n'



'''MsgName = list(selected_df_group.groups.keys())

for MsgName in MsgName:
    print(selected_df_group.get_group(MsgName))'''



# create a DBC file
output_text = output_seg01 + output_seg02 + output_seg03
with open('MyDBC.dbc', 'w', encoding='utf-8') as f:
    f.write(output_text)

print("MyDBC.dbc is generated.")
