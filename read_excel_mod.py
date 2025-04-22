import os
import sys
import traceback
import win32com.client
from typing import cast



TESTPATH = r"C:/Users/T K/Desktop/Traning/excel_comobject/sampledata/機能ファイル.xlsx"



def check_strikethrough(sheet, cell_address):
    """Excelのセル内のテキストに対する取り消し線の有無を判別する。
    引数:
        sheet: シート
        cell_address: セルアドレス
    """
    # sheet.UsageRange.Rows
    try:
        cell = sheet.Range(cell_address)
        cell_value = cell.Value
        print(type(cell_value))

        if type(cell_value) == float:
            cell_value = str(cell_value)
        elif cell_value == None:
            cell_value = ""

        temp_cell_char = cell.GetCharacters
        for i in range(1, len(cell_value)+1):
            cell_char = temp_cell_char(i,1)
            
            char_strike = cell_char.Font.Strikethrough
            char_value = cell_char.Text

            print(f"文字:{char_value},  取り消し線:{char_strike}")
    except Exception as e:
        print(f"Error: {e}")
        return

    # workbook.Close()
    # excel.Quit()


def append_text(sheet, cell_address, text, strike):
    """Excelのセルにテキストを追加する。
    引数:
        sheet: シート
        cell_address: セルアドレス
        text: 追加するテキスト
        strike: 取り消し線の有無 (True/False)
    """
    try:
        cell = sheet.Range(cell_address)
        cell_value = cell.Value

        if type(cell_value) == float:
            cell_value = str(cell_value)
        elif cell_value == None:
            cell_value = ""

        temp_cell_char = cell.GetCharacters
        start_pos = len(cell_value)+1
        for i in range(0, len(text)):
            add_char = text[i]
            
            temp_cell_char(start_pos,1).Text = add_char
            temp_cell_char(start_pos,1).Font.Strikethrough = strike
            start_pos += 1
            print(f"文字:{add_char},  取り消し線:{strike}")
        
    except Exception as e:
        tb = traceback.format_exc()
        print("⚠️ フルスタックトレース:\n", tb)
        return


def combine_cells(sheetA, sheetB, dest_sheet, cellA_adress, cellB_adress, dest_cell_adress, fillcolor):
    """【動作確認済み】2つのシートのセルのテキストを結合する。
    引数:
        sheetA: シートA
        sheetB: シートB
        dest_sheet: 結合先のシート
        cellA_adress: シートAのセルアドレス
        cellB_adress: シートBのセルアドレス
        dest_cell_adress: 結合先のセルアドレス
        fillcolor: セルの塗りつぶし色
    """
    print(type(cellA_adress), type(cellB_adress), type(dest_cell_adress))
    try:
        cellA = sheetA.Range(cellA_adress)
        cellB = sheetB.Range(cellB_adress)
        dest_cell = dest_sheet.Range(dest_cell_adress)

        if cellA.Value is None:
            cellA_value = ""
        else:
            cellA_value = cellA.Value

        if cellB.Value is None:
            cellB_value = ""
        else:
            cellB_value = cellB.Value

        if dest_cell.Value is None:
            dest_cell_value = ""
        else:
            dest_cell_value = dest_cell.Value
    except Exception as e:
        print(f"Error: {e}")
        return

    temp_cellA_char = cellA.GetCharacters
    temp_cellB_char = cellB.GetCharacters
    temp_dest_cell_char = dest_cell.GetCharacters

    start_pos = len(dest_cell_value)
    # start_pos = text_len_dest_cell
    # cellAのテキスト
    for ia in range(1,len(cellA_value)+1):
        src_cell_char = temp_cellA_char(ia,1)
        temp_dest_cell_char(start_pos,1).Text = src_cell_char.Text
        temp_dest_cell_char(start_pos,1).Font.Color = src_cell_char.Font.Color
        temp_dest_cell_char(start_pos,1).Font.Strikethrough = src_cell_char.Font.Strikethrough
        start_pos += 1

    # cellBのテキスト
    temp_msg = ""
    vbLf = "\n"
    for im in range(1,len(temp_msg)+1):
        if temp_msg[im] == vbLf:
            temp_dest_cell_char(start_pos,1).Text = vbLf
        else:
            src_cell_char = temp_msg(im,1)
            temp_dest_cell_char(start_pos,1).Text = src_cell_char.Text
            temp_dest_cell_char(start_pos,1).Font.Color = src_cell_char.Font.Color
            temp_dest_cell_char(start_pos,1).Font.Strikethrough = src_cell_char.Font.Strikethrough
        start_pos += 1
    
    # 最終的なテキスト
    for ib in range(1,len(cellB_value)+1):
        src_cell_char = temp_cellB_char(ib,1)
        temp_dest_cell_char(start_pos,1).Text = src_cell_char.Text
        temp_dest_cell_char(start_pos,1).Font.Color = src_cell_char.Font.Color
        temp_dest_cell_char(start_pos,1).Font.Strikethrough = src_cell_char.Font.Strikethrough
        start_pos += 1
    
    dest_cell.Interior.Color = fillcolor

    

        



# 使用例
excel_file = TESTPATH
sheet_name = "信号の詳細項目シート"
cell_address_A = "D6"
cell_address_B = "E6"

excel = win32com.client.Dispatch("Excel.Application")
excel = cast("win32com.client.CDispatch", excel)  # 型情報をヒントとして与える
excel.Visible = False  # Excelを非表示で実行する場合はTrue
workbook = excel.Workbooks.Open(excel_file)
sheet = workbook.Sheets(sheet_name)



append_text(sheet, cell_address_A,text="追加取り消し線対象テキスト" ,strike=True)
# check_strikethrough(sheet,cell_address_A)
combine_cells(sheetA=sheet, sheetB=sheet, dest_sheet=sheet, 
             cellA_adress=cell_address_A, cellB_adress=cell_address_B, dest_cell_adress="F6",fillcolor=0xFF0000)
# check_strikethrough(sheet, "F2")
# print(dir(excel))
workbook.Close(SaveChanges=1) # 保存して終了
excel.Quit()  # Excelを終了する
