# gph.py

import win32com.client as win32

def change_series_formula(excel_file_path, graph_sheet_name, data_sheet_name, data_range, graph_name, series_index):
    try:
        # Excelアプリケーションを起動します
        excel = win32.Dispatch("Excel.Application")
        workbook = excel.Workbooks.Open(excel_file_path)

        # グラフがあるシートを選択します
        graph_sheet = workbook.Sheets(graph_sheet_name)

        # データの参照先のシートを選択します
        data_sheet = workbook.Sheets(data_sheet_name)

        # 指定されたグラフ名と系列インデックスの系列を指定されたデータの範囲に変更します
        for chart_object in graph_sheet.ChartObjects():
            if chart_object.Name == graph_name:
                series_collection = chart_object.Chart.SeriesCollection()
                if series_index <= series_collection.Count:
                    series = series_collection(series_index)  # 指定された系列インデックスを取得
                    # 新しい式を設定
                    new_formula = f"=SERIES('{data_sheet_name}'!{data_range},{series_index})"
                    series.Formula = new_formula
                    print(f"{graph_name} の{series_index}番目の系列の式が変更されました: {new_formula}")

        # 変更を保存します
        workbook.Save()
        print("Excelファイルが保存されました。")
    except Exception as e:
        print(f"############ error occur : {str(e)}")

    # Excelを終了します
    workbook.Close()
    excel.Quit()
