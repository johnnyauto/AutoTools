import pandas as pd
import openpyxl
import re

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
# BU_: {node_name_1} {node_name_2} ...
output_seg02 = 'BU_: '

df_group = df.groupby('Transmitter')
Transmitter = list(df_group.groups.keys())
for node_name in Transmitter:
    output_seg02 += f'{node_name} '

output_seg02 += '\n\n\n'


'''output_seg03'''
# BO_ 
# 選取指定欄位
'''selected_columns = ['Transmitter', 'Message Name', 'Message ID', 'DLC']
selected_df = df[selected_columns]
selected_df_group = selected_df.groupby('Message Name')'''

output_seg03 =""
df_group = df.groupby('Message Name')
for group_index, group_data in df_group:
    transmitter = group_data['Transmitter'].iloc[0]
    if pd.isna(transmitter):
        transmitter = 'Vector__XXX'
    message_name = group_data['Message Name'].iloc[0]
    message_id = group_data['Message ID'].iloc[0]
    message_size = group_data['DLC'].iloc[0]
    # BO_ {message_id} {message_name}: {message_size} {transmitter}
    output_seg03 += f'BO_ {message_id} {message_name}: {message_size} {transmitter}\n'

# SG_
    for dataIndex in range(len(group_data)):
        # signal_name
        signal_name = group_data['Signal Name'].iloc[dataIndex]
        signal_name = signal_name.replace(' ','')
        signal_name = signal_name.replace('\n','')
        signal_name = signal_name.replace('(PS:自定义)','')
        # multiplexer_indicator
        multiplexer_indicator = ''
        # signal_size
        signal_size = group_data['size(bit)'].iloc[dataIndex]
        # factor
        factor = int(group_data['Factor'].iloc[dataIndex])
        # offset
        offset = int(group_data['Offset'].iloc[dataIndex])
        # minimum
        minimum = int(group_data['P-Minimum'].iloc[dataIndex])
        # maximum
        maximum = int(group_data['P-Maximum'].iloc[dataIndex])
        # start_bit | byte_order
        ByteOrder = group_data['Byte Order'].iloc[dataIndex]
        if ByteOrder == 'Motorola':
            start_bit = group_data['Mab'].iloc[dataIndex]
            byte_order = '0'
        else:   # Intel
            start_bit = group_data['Lab'].iloc[dataIndex]
            byte_order = '1'
        # value_type
        DataType = group_data['Data Type'].iloc[dataIndex]
        if DataType == 'unsigned':
            value_type = '+'
        else:   # signed
            value_type = '-'
        # unit
        unit = group_data['Unit'].iloc[dataIndex]
        if pd.isna(unit):
            unit = ''
        # receiver
        receiver = group_data['Receiver'].iloc[dataIndex]
        receiver = receiver.replace('\n',',')
        receiver = receiver.replace('/',',')

        # SG_ {signal_name} {multiplexer_indicator} : {start_bit}|{signal_size}@{byte_order}{value_type} ({factor},{offset}) [{minimum}|{maximum}] "{unit}" {receiver}
        output_seg03 += f' SG_ {signal_name} {multiplexer_indicator}: {start_bit}|{signal_size}@{byte_order}{value_type} ({factor},{offset}) [{minimum}|{maximum}] "{unit}" {receiver}\n'

        if dataIndex == len(group_data)-1:
            output_seg03 += '\n'
output_seg03 += '\n\n'


'''output_seg04 & output_seg05'''
# BA_DEF_
output_seg04 = ''
output_seg04 += 'BA_DEF_ BO_  "GenMsgFastOnStart" INT 0 0;\n'
output_seg04 += 'BA_DEF_ SG_  "GenSigInactiveValue" INT 0 0;\n'
output_seg04 += 'BA_DEF_ BU_  "ILUsed" ENUM  "Yes","No";\n'
output_seg04 += 'BA_DEF_ SG_  "GenSigStartValue" FLOAT 0 100000000000;\n'
output_seg04 += 'BA_DEF_ SG_  "GenSigSendType" ENUM  "Cyclic","OnWrite","OnWriteWithRepetition","OnChange","OnChangeWithRepetition","IfActive","IfActiveWithRepetition","NoSigSendType","OnChangeAndIfActive","OnChangeAndIfActiveWithRepetition";\n'
output_seg04 += 'BA_DEF_ BO_  "GenMsgNrOfRepetition" INT 0 999999;\n'
output_seg04 += 'BA_DEF_ BO_  "GenMsgDelayTime" INT 0 1000;\n'
output_seg04 += 'BA_DEF_ BO_  "GenMsgCycleTime" INT 2 50000;\n'
output_seg04 += 'BA_DEF_ BO_  "GenMsgSendType" ENUM  "Cyclic","not_used","not_used","not_used","not_used","not_used","not_used","IfActive","NoMsgSendType";\n'
output_seg04 += 'BA_DEF_ BO_  "GenMsgCycleTimeFast" INT 2 100000;\n'
output_seg04 += 'BA_DEF_ BO_  "GenMsgILSupport" ENUM  "No","Yes";\n'
output_seg04 += 'BA_DEF_ BO_  "GenMsgStartDelayTime" INT 0 100000;\n'
output_seg04 += 'BA_DEF_ BU_  "NodeLayerModules" STRING ;\n'
# BA_DEF_DEF_
output_seg04 += 'BA_DEF_DEF_  "GenMsgFastOnStart" 0;\n'
output_seg04 += 'BA_DEF_DEF_  "GenSigInactiveValue" 0;\n'
output_seg04 += 'BA_DEF_DEF_  "ILUsed" "Yes";\n'
output_seg04 += 'BA_DEF_DEF_  "GenSigStartValue" 0;\n'
output_seg04 += 'BA_DEF_DEF_  "GenSigSendType" "Cyclic";\n'
output_seg04 += 'BA_DEF_DEF_  "GenMsgNrOfRepetition" 0;\n'
output_seg04 += 'BA_DEF_DEF_  "GenMsgDelayTime" 0;\n'
output_seg04 += 'BA_DEF_DEF_  "GenMsgCycleTime" 100;\n'
output_seg04 += 'BA_DEF_DEF_  "GenMsgSendType" "Cyclic";\n'
output_seg04 += 'BA_DEF_DEF_  "GenMsgCycleTimeFast" 0;\n'
output_seg04 += 'BA_DEF_DEF_  "GenMsgILSupport" "Yes";\n'
output_seg04 += 'BA_DEF_DEF_  "GenMsgStartDelayTime" 0;\n'
output_seg04 += 'BA_DEF_DEF_  "NodeLayerModules" "CANoeILNLVector.dll";\n'
# BA_
output_seg05 = ''
output_seg06 = ''

for group_index, group_data in df_group:
    message_id = group_data['Message ID'].iloc[0]
    message_type = group_data['Message Type'].iloc[0]

    if message_type == 'P':
        # setup for "GenMsgCycleTime", "GenMsgSendType" use default
        MsgCycleTime = int(group_data['period\n(ms)'].iloc[0])
        output_seg04 += f'BA_ "GenMsgCycleTime" BO_ {message_id} {MsgCycleTime};\n'
    elif message_type == 'E':
        # setup for "GenMsgSendType"
        MsgSendType = '8'   # NoMsgSendType
        output_seg04 += f'BA_ "GenMsgSendType" BO_ {message_id} {MsgSendType};\n'

        for dataIndex in range(len(group_data)):
            signal_name = group_data['Signal Name'].iloc[dataIndex]
            SigSendType = 1 # OnWrite
            output_seg05 += f'BA_ "GenSigSendType" SG_ {message_id} {signal_name} {SigSendType};\n'
# VAL_
    for dataIndex in range(len(group_data)):
        # 只對Unit為empty的signal建立value table
        if pd.isna(group_data['Unit'].iloc[dataIndex]):
            ValTable = group_data['Coding'].iloc[dataIndex]
            ValTable = ValTable.replace(' : ',':')
            ValTable = ValTable.replace('：',':')
            ValTable = ValTable.replace(': ',':')
            ValTable = ValTable.replace(' :',':')
            ValTable = ValTable.replace(' ~ ','~')
            ValTable = ValTable.replace(':~','~')
            ValTable = ValTable.replace('-0','~0')
            ValTable = ValTable.replace('Ox','0x')
            ValTable = ValTable.replace('"','')
            ValTable = ValTable.replace('0x16-1F','0x16~1F')
            #ValTable = ValTable.replace(' ','_')
            # if ValTable is not empty
            if not pd.isna(ValTable):
                signal_name = group_data['Signal Name'].iloc[dataIndex]
                signal_name = signal_name.replace(' ','')
                signal_name = signal_name.replace('\n','')
                signal_name = signal_name.replace('(PS:自定义)','')
                # split data by '\n'
                ValTable_split = ValTable.split("\n")
                print('ValTable_split size', len(ValTable_split))

                Value_Description = ''
                for data in ValTable_split:
                    if '~' in data:
                        '''
                        # 使用正則表達式將資料 0xAA~0xBB:CCCC 分割為 '0xAA', '0xBB' 和 'CCCC' 三部分
                        print(data)
                        pattern = re.compile(r'(\w+x\w+)~(\w+x\w+):(\w+)')
                        regular_data = pattern.match(data)
                        start, end, Description = regular_data.groups()
                        for Value in range(int(start, 16), int(end, 16)+1):
                            Value_Description = f'{Value} "{Description}" '
                            output_seg06 += f'VAL_ {message_id} {signal_name} {Value_Description};\n'
                        '''
                        # 不處理帶有~的value table, 因為有些值的範圍太大
                        continue
                    else:
                        # split data to two elements (Value & Description) by ':'
                        data_split = data.split(":", 1)
                        # due to some data lacks ':', use " " to split instead
                        if len(data_split) == 1:
                            data_split = data.split(" ", 1)
                        print(data_split)
                        #print('data_split size = ', len(data_split))
                        #Value = data_split[0]
                        Value = int(data_split[0], 16)
                        Description = data_split[1]
                        Value_Description += f'{Value} "{Description}" '
                output_seg06 += f'VAL_ {message_id} {signal_name} {Value_Description};\n'


# create a DBC file
output_text = output_seg01 + output_seg02 + output_seg03 + output_seg04 + output_seg05 + output_seg06
with open('MyDBC.dbc', 'w', encoding='utf-8') as f:
    f.write(output_text)

print("MyDBC.dbc is generated.")
