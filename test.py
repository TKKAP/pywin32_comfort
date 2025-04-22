import win32com.client
import os

SIGNALNAME_COL = 1  # 信号名の列番号

# ===== Excelの初期化とファイル読み込み =====
def open_excel_app():
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        return excel
    except Exception as e:
        print(f"Error initializing Excel: {e}")
        return None
    
def open_workbook(excel, filepath):
    try:
        return excel.Workbooks.Open(os.path.abspath(filepath))
    except Exception as e:
        print(f"Error opening workbook {filepath}: {e}")
        return None


# ===== 色付きセルかどうかを判定 =====
def is_colored(cell):
    """色付きセルかどうかを判定する関数"""
    try:
        return (cell.Interior.Color) != 16777215 and (cell.Interior.Color != 0)  # 白または透明じゃない
    except Exception as e:
        print(f"Error checking cell color: {e}")
        return False


# ===== 信号名で束ねファイルから対応行を探す =====
def find_matching_row(sheet, signal_name, col=SIGNALNAME_COL):
    """信号名で束ねファイルから対応行を探す関数"""
    try:
        for row in range(2, sheet.UsedRange.Rows.Count + 1):
            if sheet.Cells(row, col).Value == signal_name:
                return row
        return None
    except Exception as e:
        print(f"Error finding matching row: {e}")
        return None


# ===== 機能ファイルの行を束ねファイルに挿入 =====
def insert_row_from_func(mg_sheet, func_sheet, func_row, insert_at_row):
    """機能ファイルの行を束ねファイルに挿入する関数"""
    try:
        mg_sheet.Rows(insert_at_row + 1).Insert()
        for col in range(1, func_sheet.UsedRange.Columns.Count + 1):
            mg_sheet.Cells(insert_at_row + 1, col).Value = func_sheet.Cells(func_row, col).Value
            mg_sheet.Cells(insert_at_row + 1, col).Interior.Color = func_sheet.Cells(func_row, col).Interior.Color
    except Exception as e:
        print(f"Error inserting row: {e}")  


# ===== セルにテキストと書式を追記 =====
def append_text_and_format(dest_cell, source_cell):
    """セルにテキストと書式を追記する関数"""
    try:
        existing_text = dest_cell.Value or ""
        new_text = f"{existing_text}\n{source_cell.Value}" if existing_text else source_cell.Value
        dest_cell.Value = new_text
        dest_cell.Interior.Color = source_cell.Interior.Color
        # フォントのコピーなど必要があればここで追加

    except Exception as e:
        print(f"Error appending text and format: {e}")

# ===== かぶり項目抽出シートに記録 =====
def record_conflict(conflict_sheet, func_sheet, func_row):
    """かぶり項目抽出シートに記録する関数"""
    try:
        last_row = conflict_sheet.UsedRange.Rows.Count + 1
        for i in range(1, func_sheet.UsedRange.Columns.Count + 1):
            conflict_sheet.Cells(last_row, i).Interior.Color = func_sheet.Cells(func_row, i).Interior.Color
            conflict_sheet.Cells(last_row, i).Value = func_sheet.Cells(func_row, i).Value
            
    except Exception as e:
        print(f"Error recording conflict: {e}")


# ===== メイン処理フロー =====
def process_files(mg_wb, func_wb):
    mg_sheet = mg_wb.Sheets("信号の詳細項目シート")
    mg_conflict_sheet = mg_wb.Sheets("かぶり項目抽出シート")
    func_sheet = func_wb.Sheets("信号の詳細項目シート")

    i_row = 2
    max_row = func_sheet.UsedRange.Rows.Count
    mg_row_index = 2
    for mg_row_index in range(2, mg_sheet.UsedRange.Rows.Count + 1):
        for mg_column_index in range(1, mg_sheet.UsedRange.Columns.Count + 1):
            
            for func_row_index in range(2, func_sheet.UsedRange.Rows.Count + 1):
                for func_column_index in range(1, func_sheet.UsedRange.Columns.Count + 1):
                    func_cell = func_sheet.Cells(func_row_index, func_column_index)
                    
                    # if func_cell.Value is None:
                    #     break
                    # cell = func_sheet.Cells(mg_row_index, j_col)
                    # if cell.Value is None:
                    #     break

                    if is_colored(func_cell):
                        signal_name = func_sheet.Cells(mg_row_index, 1).Value
                        matched_row = find_matching_row(mg_sheet, signal_name)

                        col_header = func_sheet.Cells(1, func_column_index).Value
                        if col_header == "信号名" and matched_row:
                            # 新規作成
                            insert_row_from_func(mg_sheet, func_sheet, mg_row_index, matched_row)
                            continue  # 次の行へ
                        elif matched_row:
                            # 追記
                            target_mg_cell = mg_sheet.Cells(matched_row, mg_column_index)
                            append_text_and_format(target_mg_cell, func_cell)

                            if is_colored(target_mg_cell):
                                record_conflict(mg_conflict_sheet, func_sheet, func_row_index)
                    else:
                        pass


# ===== 実行部分 =====
def main():
    mg_filepath = r"C:\Users\T K\Desktop\Traning\excel_comobject\sampledata\束ねファイル.xlsx"
    func_filepath = r"C:\Users\T K\Desktop\Traning\excel_comobject\sampledata\機能ファイル.xlsx"

    excel = open_excel_app()
    mg_wb = open_workbook(excel, mg_filepath)
    func_wb = open_workbook(excel, func_filepath)

    process_files(mg_wb, func_wb)

    mg_wb.Save()
    func_wb.Close(False)
    mg_wb.Close(True)
    excel.Quit()


if __name__ == "__main__":
    main()