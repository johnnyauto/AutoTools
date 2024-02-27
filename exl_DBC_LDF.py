import exlDBC
import exlLDF


while True:
    print('\nThis program can generate DBC and LDF files from excel.\n')
    print('0: DBC file')
    print('1: LDF file')
    generator = input("請輸入數字選擇要產生哪種檔案, 預設值為'0': ")

    if generator == '':
        generator = '0'

    if generator == '0':
        exlDBC.dbc_main()
        break
    elif generator == '1':
        exlLDF.ldf_main()
        break
    else:
        pass

