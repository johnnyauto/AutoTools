import pandas as pd
import openpyxl

##### fun(): process data and generate data frame #####
def process_data(sheet_name):
    worksheet = workbook[sheet_name]

    # Process Data (remove Empty and Strikethrough format data) 
    pData = []      # processed Data
    Sig_index = 6   # column index of 'Signal Name'
    Msg_index = 10  # column index of 'Message Name'
    Msb_index = 20  # column index of 'Msb'

    # for row_index in worksheet.iter_rows(values_only=True):
    for row_index in range(2, worksheet.max_row+1):
        Msg_value = worksheet.cell(row=row_index, column=Msg_index).value
        Msb_value = worksheet.cell(row=row_index, column=Msb_index).value
        Msg_strike = worksheet.cell(row=row_index, column=Msg_index).font.strike
        Sig_strike = worksheet.cell(row=row_index, column=Sig_index).font.strike

        if Msg_value and Msb_value and not Msg_strike and not Sig_strike:
            # generate a processed Data
            pData.append([worksheet.cell(row=row_index, column=col).value for col in range(1, worksheet.max_column + 1)])

    # get column name
    columns = [worksheet.cell(row=1, column=col).value for col in range(1, worksheet.max_column + 1)]

    # convert pData to DataFrame
    df = pd.DataFrame(pData, columns=columns)
    # convert 'Message ID' from Hex to Dec format
    df['Message ID'] = df['Message ID'].apply(lambda x: int(x, 16))
    return df


##### fun(): output_seg01 [LDF config] #####
def ldf_cfg(df):
    output_seg01 = '\n\nLIN_description_file;\n'
    output_seg01 += 'LIN_protocol_version = "2.1";\n'
    output_seg01 += 'LIN_language_version = "2.1";\n'
    output_seg01 += 'LIN_speed = 19.2 kbps;\n'
    return output_seg01


##### fun(): output_seg02 [Nodes] #####
def ldf_notes(df, slave_node_list):
    notes_list = []
    time_base = '10'
    jitter = '0.1'

    # add node_name from 'Transmitter'
    df_group_tx = df.groupby('Transmitter')
    df_group_rx = df.groupby('Receiver')
    Transmitter = list(df_group_tx.groups.keys())
    Receiver = list(df_group_rx.groups.keys())
    notes_list = Transmitter + Receiver
    for index, node in enumerate(notes_list):
        print(f'{index}: {node}')

    index = input('請選擇一個Node作為Master node: ')
    master_node = notes_list[int(index)]

    slave_nodes = ''
    for node in notes_list:
        if node != master_node:
            slave_nodes += f'{node}, '
            slave_node_list.append(node) # for general used
    slave_nodes = slave_nodes[:-2]

    output_seg02 = '\nNodes {\n'
    output_seg02 += f'  Master: {master_node}, {time_base} ms, {jitter} ms ;\n'
    output_seg02 += f'  Slave: {slave_nodes} ;\n'
    output_seg02 += '}\n'
    return output_seg02


##### fun(): output_seg03 [Signals] #####
def ldf_sig_def(df):
    output_seg03 = '\nSignals {\n'
    for index, row in df.iterrows():
        signal_name = row['Signal Name']
        size_bit = int(row['size(bit)'])
        init_val = row['Default Initialised value']
        if '0x' in str(init_val):
            init_val = pd.Series(init_val).apply(lambda x: int(x, 16)) # .apply() only for DataFrame
            init_val = init_val.values # init_val become a list[] with a element
            init_val = init_val[0] # get the value for the list
        tx_node = row['Transmitter']
        rx_node = row['Receiver']
        if '\n' in rx_node:
            rx_node = rx_node.replace('\n',', ')
        
        output_seg03 += f'  {signal_name}: {size_bit}, {init_val}, {tx_node}, {rx_node} ;\n'
    output_seg03 += '}\n'
    return output_seg03
        

##### fun(): output_seg04 [Diagnostic_signals] #####
def ldf_diag_sig(df):
    output_seg04 = '\nDiagnostic_signals {\n'
    output_seg04 += '  MasterReqB0: 8, 0 ;\n'
    output_seg04 += '  MasterReqB1: 8, 0 ;\n'
    output_seg04 += '  MasterReqB2: 8, 0 ;\n'
    output_seg04 += '  MasterReqB3: 8, 0 ;\n'
    output_seg04 += '  MasterReqB4: 8, 0 ;\n'
    output_seg04 += '  MasterReqB5: 8, 0 ;\n'
    output_seg04 += '  MasterReqB6: 8, 0 ;\n'
    output_seg04 += '  MasterReqB7: 8, 0 ;\n'
    output_seg04 += '  SlaveRespB0: 8, 0 ;\n'
    output_seg04 += '  SlaveRespB1: 8, 0 ;\n'
    output_seg04 += '  SlaveRespB2: 8, 0 ;\n'
    output_seg04 += '  SlaveRespB3: 8, 0 ;\n'
    output_seg04 += '  SlaveRespB4: 8, 0 ;\n'
    output_seg04 += '  SlaveRespB5: 8, 0 ;\n'
    output_seg04 += '  SlaveRespB6: 8, 0 ;\n'
    output_seg04 += '  SlaveRespB7: 8, 0 ;\n'
    output_seg04 += '}\n'
    return output_seg04


##### fun(): output_seg05 [Frames] #####
def ldf_data_frame_def(df):
    output_seg05 = '\nFrames {\n'
    df_group = df.groupby('Message Name')
    for group_index, group_data in df_group:
        frame_name = group_data['Message Name'].iloc[0]
        frame_id = group_data['Message ID'].iloc[0]
        tx_node = group_data['Transmitter'].iloc[0]
        frame_size = group_data['DLC'].iloc[0]
        output_seg05 += f'  {frame_name}: {frame_id}, {tx_node}, {frame_size} '
        output_seg05 += '{\n'
        for data_index in range(len(group_data)):
            signal_name = group_data['Signal Name'].iloc[data_index]
            start_bit = group_data['Lsb'].iloc[data_index]
            output_seg05 += f'    {signal_name}, {start_bit} ;\n'
        output_seg05 += '  }\n'
    output_seg05 += '}\n'
    return output_seg05


##### fun(): output_seg06 [Diagnostic_frame] #####
def ldf_diag_frame(df):
    output_seg06 = '\nDiagnostic_frame {\n'
    output_seg06 += '  MasterReq: 0x3c {\n'
    output_seg06 += '    MasterReqB0, 0 ;\n'
    output_seg06 += '    MasterReqB1, 8 ;\n'
    output_seg06 += '    MasterReqB2, 16 ;\n'
    output_seg06 += '    MasterReqB3, 24 ;\n'
    output_seg06 += '    MasterReqB4, 32 ;\n'
    output_seg06 += '    MasterReqB5, 40 ;\n'
    output_seg06 += '    MasterReqB6, 48 ;\n'
    output_seg06 += '    MasterReqB7, 56 ;\n'
    output_seg06 += '  }\n'
    output_seg06 += '  SlaveResp: 0x3d {\n'
    output_seg06 += '    SlaveRespB0, 0 ;\n'
    output_seg06 += '    SlaveRespB1, 8 ;\n'
    output_seg06 += '    SlaveRespB2, 16 ;\n'
    output_seg06 += '    SlaveRespB3, 24 ;\n'
    output_seg06 += '    SlaveRespB4, 32 ;\n'
    output_seg06 += '    SlaveRespB5, 40 ;\n'
    output_seg06 += '    SlaveRespB6, 48 ;\n'
    output_seg06 += '    SlaveRespB7, 56 ;\n'
    output_seg06 += '  }\n'
    output_seg06 += '}\n'
    return output_seg06


##### fun(): output_seg07 [Node_attributes] #####
def ldf_node_attr(df):
    output_seg07 = '\nNode_attributes {\n'
    for slave_node in slave_node_list:
        output_seg07 += f'  {slave_node} '
        output_seg07 += '{\n'
        output_seg07 += '    LIN_protocol = "2.1" \n'
        output_seg07 += '    configured_NAD = 0 \n'
        output_seg07 += '    initial_NAD = 0x0 \n'
        output_seg07 += '    product_id = 0xFFFF, 0xFFFF, 0xFFF \n'
        output_seg07 += '    P2_min = 50 ms \n'
        output_seg07 += '    ST_min = 0 ms \n'
        output_seg07 += '    N_As_timeout = 1000 ms \n'
        output_seg07 += '    N_Cr_timeout = 1000 ms \n'
        output_seg07 += '    configurable_frames {\n'

        df_group = df.groupby('Message Name')
        for group_index, group_data in df_group:
            frame = group_data['Message Name'].iloc[0]
            tx_node = group_data['Transmitter'].iloc[0]
            rx_node_list = group_data['Receiver'].iloc[0].split('\n')
            # check whether the slave node is related to the Transmitter or Receiver of the message
            # if so, the message belongs configurable_frames
            if slave_node == tx_node or slave_node in rx_node_list:
                output_seg07 += f'      {frame} ;\n'

        output_seg07 += '    }\n'
        output_seg07 += '  }\n'
    output_seg07 += '  }\n'
    return output_seg07


##### fun(): output_seg08 [Schedule_tables] #####
def ldf_sch_table():
    df_sch_table = pd.read_excel(excel_file, sheet_name='LIN_Schedule Table')
    '''new_column = df_sch_table.iloc[1]
    df_sch_table = df_sch_table.iloc[2:].rename(columns=new_column)
    df_sch_table.reset_index(drop=True, inplace=True)'''

    output_seg08 = '\nSchedule_tables {\n'
    output_seg08 += ' Table1 {\n' # schedule table name
    df_sch_table = df_sch_table.drop([0])
    for index, row in df_sch_table.iterrows():
        frame = row.iloc[3]
        delay_time = row.iloc[2]
        output_seg08 += f'    {frame} delay {delay_time} ms ;\n'
    
    output_seg08 += '  }\n'
    output_seg08 += '}\n'
    return output_seg08



##### Main #####
while True:
    try:
        # load Excel file
        print('This program will generate LDF files from excel.\n')
        excel_file = input('請輸入欲轉換的Excel檔名: ')
        if not '.xlsx' in excel_file:
            excel_file += '.xlsx'
        workbook = openpyxl.load_workbook(excel_file, data_only=True)
        break
    except:
        print('\nError! 檔名錯誤或找不到檔案路徑')
        input('Press [Enter] to continue.\n')

while True:
    # get sheetName and sheet_index from workbook
    sheet_name_list = workbook.sheetnames
    print('\n\n[Sheet list]')
    for index, sheetName in enumerate(sheet_name_list):
        print(f'{index}: {sheetName}')

    sheet_index = input("選擇一個sheet(輸入數字)生成LDF, 或輸入'q'結束程式: ")
    
    if sheet_index.lower() == 'q':
        break
    else:
        sheetName = sheet_name_list[int(sheet_index)]
        # process data and generate df(data frame)
        df = process_data(sheetName)
        slave_node_list = []
        output_01 = ldf_cfg(df)
        output_02 = ldf_notes(df, slave_node_list)
        output_03 = ldf_sig_def(df)
        output_04 = ldf_diag_sig(df)
        output_05 = ldf_data_frame_def(df)
        output_06 = ldf_diag_frame(df)
        output_07 = ldf_node_attr(df)     
        output_08 = ldf_sch_table()  
        
        output_text = output_01 + output_02 + output_03 + output_04 + output_05 + output_06 + output_07 + output_08
        with open('testLDF.ldf', 'w', encoding='utf-8') as f:
            f.write(output_text)
        print('\nLDF is generated!!\n')
        #print(slave_node_list)
        input('Press [Enter] to continue.')
        break

        '''
        # generate DBC files
        try:
            sheetName = sheet_name_list[int(sheet_index)]
            # process data and generate df(data frame)
            df = process_data(sheetName)
            ldf_notes(df)
            break
        except:
            print('\nError! 請確認所選擇的sheet內容是否正確')
            input('Press [Enter] to continue.\n')'''
