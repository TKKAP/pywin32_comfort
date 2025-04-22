```mermaid
flowchart TD
    %% create_mgfile[束ねファイル作成]
    %% file_mg[束ねファイル読み込み]
    %% file_func[機能ファイル読み込み]

    %% check_funcfile[機能ファイルに色付きのセルが無いかチェック]
    %% exist_coloredcell{機能ファイルの参照中のセルがなんらかの色で塗りつぶされている?}

    %% %% search_samesignalcell_mgfile[参照中の機能ファイルの信号名と同じ信号名を持つ行を束ねファイルから検索]
    %% condition_exist_samesignalcell{参照中の行の信号名と一致する行が束ねファイル側にもあるか?}
        
    %% append_mgfile_cell[束ねファイルに機能ファイルの色付きセルのテキストを追記]
    %% i_col_plus1[marge_col++]
    %% i_row_plus1[marge_row++]

    %% j_col_plus1[func_col++]
    %% j_row_plus1[func_row++]

    %% funcfile_nextcell[機能ファイルの次のセルに移動]

    %% compare_signalcell[束ねファイルi行目,信号列のセルと機能ファイルのj行目,信号列のセルを比較]

    %% condition_signalname{信号名が一致するか?}

    create_mgfile[束ねファイル作成] --> file_mg[束ねファイル読み込み]
    file_mg --> file_func[機能ファイル読み込み]
    file_func --> exist_coloredcell{機能ファイルの参照中のセルがなんらかの色で塗りつぶされている?}
    
    
    exist_coloredcell -- はい --> search_samesignal_row[参照中の機能ファイルの信号名と同じ信号名を持つ行を束ねファイルから検索]
    search_samesignal_row -- はい --> what_column{機能ファイルの参照中のセルはなんの列か?}
    what_column -- 信号名 --> insert_funcrow2mgrow[束ねファイルに参照中の機能ファイルの行を挿入]
    what_column -- 備考欄 --> append_text[束ねファイルに機能ファイルの色付きセルのテキストと書式情報を追記]
    what_column -- その他 --> append_text[束ねファイルに機能ファイルの色付きセルのテキストと書式情報を追記]
    append_text --> paint_cell[束ねファイルのセルを塗りつぶし]
    paint_cell --> is_already_painted{束ねファイルのセルはすでに塗りつぶされているか?}
    is_already_painted -- はい --> record_paintedrow[参照中の機能ファイルの列を束ねファイルの被り項目抽出シートに追加]

    record_paintedrow --> i_row_plus1[i_row++]
    insert_funcrow2mgrow --> i_row_plus1[i_row++]
    i_row_plus1 --> exist_coloredcell

    exist_coloredcell -- いいえ --> j_col_plus1[j_col++]
    j_col_plus1 --> exist_coloredcell
    %% funcfile_nextcell --> check_funcfile


    %% A[開始] --> B[データ入力]
    %% B --> C{データは有効か?}
    %% C -- はい --> D[処理を実行]
    %% C -- いいえ --> E[エラーを表示]
    %% D --> F[終了]
    %% E --> F
````
