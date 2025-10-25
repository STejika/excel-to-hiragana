import streamlit as st
import pandas as pd
import pykakasi
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
    """Excelファイルの全セルの内容をひらがなに変換する"""
    # Excelファイルを読み込む
    df = pd.read_excel(file)
    
    # 全てのセルに対してひらがな変換を適用
    for column in df.columns:
        df[column] = df[column].apply(convert_to_hiragana)
    
    return df

def main():
    st.title("Excel ひらがな変換アプリ")
    st.write("Excelファイルのすべてのセルをひらがなに変換します")
    
    # ファイルアップローダーの表示
    uploaded_file = st.file_uploader("Excelファイルを選択してください", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        try:
            # ファイルの処理
            converted_df = process_excel_file(uploaded_file)
            
            # 変換結果の表示
            st.write("変換結果:")
            st.dataframe(converted_df)
            
            # 変換済みファイルのダウンロード用にExcelファイルを作成
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                converted_df.to_excel(writer, index=False)
            
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