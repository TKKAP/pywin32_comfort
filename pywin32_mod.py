import os
import win32com.client



TESTPATH = rCUsersT KDesktopTraningexcel_comobjectsampledatatestbook.xlsx



def check_strikethrough(sheet, cell_address)
    Excelのセル内のテキストに対する取り消し線の有無を判別する。
    # sheet.UsageRange.Rows

    cell = sheet.Range(cell_address)
    cell_value = cell.Value
    print(type(cell_value))

    if type(cell_value) == float
        cell_value = str(cell_value)
    elif cell_value == None
        cell_value = 

    temp_cell_char = cell.GetCharacters
    for i in range(1, len(cell_value)+1)
        cell_char = temp_cell_char(i,1)
        
        char_strike = cell_char.Font.Strikethrough
        char_value = cell_char.Text

        print(f文字{char_value},  取り消し線{char_strike})


    # workbook.Close()
    # excel.Quit()



def combine_cells(sheetA, sheetB, dest_sheet, cellA_adress, cellB_adress, dest_cell_adress)
    2つのシートのセルのテキストを結合する。
    print(type(cellA_adress), type(cellB_adress), type(dest_cell_adress))
    try
        cellA = sheetA.Range(cellA_adress)
        cellB = sheetB.Range(cellB_adress)
        dest_cell = dest_sheet.Range(dest_cell_adress)

        if cellA.Value is None
            cellA_value = 
        else
            cellA_value = cellA.Value

        if cellB.Value is None
            cellB_value = 
        else
            cellB_value = cellB.Value

        if dest_cell.Value is None
            dest_cell_value = 
        else
            dest_cell_value = dest_cell.Value
    except Exception as e
        print(fError {e})
        return

    # if type(cellA_value) == float
    #     cellA_value = str(cellA_value)
    # elif cellA_value == None
    #     cellA_value = 

    # if type(cellB_value) == float
    #     cellB_value = str(cellB_value)
    # elif cellB_value == None
    #     cellB_value = 

    # if type(dest_cell_value) == float
    #     dest_cell_value = str(dest_cell_value)
    # elif dest_cell_value == None
    #     dest_cell_value = 

    temp_cellA_char = cellA.GetCharacters
    temp_cellB_char = cellB.GetCharacters
    temp_dest_cell_char = dest_cell.GetCharacters

    start_pos = len(dest_cell_value)
    # start_pos = text_len_dest_cell
    for ia in range(1,len(cellA_value)+1)
        src_cell_char = temp_cellA_char(ia,1)
        temp_dest_cell_char(start_pos,1).Text = src_cell_char.Text
        temp_dest_cell_char(start_pos,1).Font.Color = src_cell_char.Font.Color
        temp_dest_cell_char(start_pos,1).Font.Strikethrough = src_cell_char.Font.Strikethrough
        start_pos += 1

    # 挿入するメッセージ
    temp_msg = 
    vbLf = n
    for im in range(1,len(temp_msg)+1)
        if temp_msg[im] == vbLf
            temp_dest_cell_char(start_pos,1).Text = vbLf
        else
            src_cell_char = temp_msg(im,1)
            temp_dest_cell_char(start_pos,1).Text = src_cell_char.Text
            temp_dest_cell_char(start_pos,1).Font.Color = src_cell_char.Font.Color
            temp_dest_cell_char(start_pos,1).Font.Strikethrough = src_cell_char.Font.Strikethrough
        
        
        start_pos += 1


    for ib in range(1,len(cellB_value)+1)
        src_cell_char = temp_cellB_char(ib,1)
        temp_dest_cell_char(start_pos,1).Text = src_cell_char.Text
        temp_dest_cell_char(start_pos,1).Font.Color = src_cell_char.Font.Color
        temp_dest_cell_char(start_pos,1).Font.Strikethrough = src_cell_char.Font.Strikethrough
        start_pos += 1

    

        



# 使用例
excel_file = TESTPATH
sheet_name = Sheet1
cell_address_A = D2
cell_address_B = E2

excel = win32com.client.Dispatch(Excel.Application)
excel.Visible = True  # Excelを非表示で実行する場合はTrue
workbook = excel.Workbooks.Open(excel_file)
sheet = workbook.Sheets(sheet_name)

check_strikethrough(sheet,cell_address_A)
combine_cells(sheetA=sheet, sheetB=sheet, dest_sheet=sheet, 
              cellA_adress=cell_address_A, cellB_adress=cell_address_B, dest_cell_adress=F2)
check_strikethrough(sheet, F2)