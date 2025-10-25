# Excel ひらがな変換アプリ

Excelファイルのすべてのセルの内容をひらがなに変換するStreamlitアプリケーションです。

## 機能

- Excelファイル（.xlsx, .xls）のアップロード
- すべてのセルの内容をひらがなに変換
- 変換結果のプレビュー表示
- 変換済みファイルのダウンロード

## 必要要件

- Python 3.8以上
- 必要なパッケージ：
  - streamlit
  - pandas
  - openpyxl
  - pykakasi

## インストール方法

```bash
# リポジトリのクローン
git clone https://github.com/[ユーザー名]/excel-to-hiragana.git
cd excel-to-hiragana

# 必要なパッケージのインストール
pip install -r requirements.txt
```

## 使用方法

1. 以下のコマンドでアプリケーションを起動します：
```bash
streamlit run excel_to_hiragana.py
```

2. ブラウザで表示されたURL（通常は http://localhost:8501 ）にアクセスします。

3. 「ファイルを選択」ボタンをクリックしてExcelファイルをアップロードします。

4. 変換結果を確認し、必要に応じて「変換済みファイルをダウンロード」ボタンをクリックします。

## ライセンス

MITライセンス