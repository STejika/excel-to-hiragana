import streamlit as st
import pandas as pd
import pykakasi
import xlrd  # .xls ファイル用
import openpyxl  # .xlsx ファイル用
from io import BytesIO

def convert_to_hiragana(text):
    """テキストをひらがなに変換する関数"""
    if pd.isna(text):
        return ""
    if not isinstance(text, str):
        text = str(text)
    
    kks = pykakasi.kakasi()
    result = kks.convert(text)
    hiragana = ''.join([item['hira'] for item in result])
    return hiragana

def process_excel_file(file):
    """Excelファイルの全シートの全セルの内容をひらがなに変換する"""
    # ファイル拡張子の確認
    file_extension = file.name.split('.')[-1].lower()
    
    # Excelファイルの全シートを読み込む
    excel_file = pd.ExcelFile(file)
    sheet_names = excel_file.sheet_names
    
    # 各シートのDataFrameを保存する辞書
    dfs = {}
    
    # 各シートに対して処理を実行
    for sheet_name in sheet_names:
        # シートを読み込む
        df = pd.read_excel(file, sheet_name=sheet_name)
        
        # 全てのセルに対してひらがな変換を適用
        for column in df.columns:
            df[column] = df[column].apply(convert_to_hiragana)
        
        # 変換済みのDataFrameを辞書に保存
        dfs[sheet_name] = df
    
    return dfs

def main():
    st.title("Excel ひらがな変換アプリ")
    st.write("Excelファイルのすべてのセルをひらがなに変換します")
    
    # ファイルアップローダーの表示
    uploaded_file = st.file_uploader("下のボタンを押してExcelファイルを選択するか，ファイルをドラッグアンドドロップしてください", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        try:
            # ファイルの処理
            converted_dfs = process_excel_file(uploaded_file)
            
            # 変換結果の表示
            st.write("変換結果:")
            for sheet_name, df in converted_dfs.items():
                st.write(f"シート: {sheet_name}")
                st.dataframe(df)
                st.divider()  # シート間の区切り線
            
            # 変換済みファイルのダウンロード用にExcelファイルを作成
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for sheet_name, df in converted_dfs.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # ダウンロードボタンの表示
            st.download_button(
                label="変換済みファイルをダウンロード",
                data=output.getvalue(),
                file_name="converted_hiragana.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"エラーが発生しました: {str(e)}")

if __name__ == "__main__":
    main()

