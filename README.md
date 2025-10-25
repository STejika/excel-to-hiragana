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
  - xlrd  # 古い .xls 形式を読み込むために必要

## インストール方法

```bash
# リポジトリのクローン
git clone https://github.com/[ユーザー名]/excel-to-hiragana.git
cd excel-to-hiragana

# 必要なパッケージのインストール
pip install -r requirements.txt
```

推奨：プロジェクトと依存関係は仮想環境で管理してください。PowerShell の場合の例：

```powershell
# 仮想環境作成（まだ作っていない場合）
python -m venv .venv

# 仮想環境をアクティブ化
# 実行ポリシーの制限がある場合は venv の Python を直接使ってインストール/起動する方法もあります（下参照）
.\.venv\Scripts\Activate

# 依存関係をインストール
\.venv\Scripts\python.exe -m pip install -r requirements.txt

# Streamlit を起動（仮想環境の python を利用）
\.venv\Scripts\python.exe -m streamlit run excel_to_hiragana.py
```

実行ポリシーのために `Activate` ができない場合は、仮想環境の Python を直接呼ぶ方法を使ってください（上の `\.venv\Scripts\python.exe` を利用）。

## 使用方法

1. 以下のコマンドでアプリケーションを起動します：
```bash
streamlit run excel_to_hiragana.py
```

2. ブラウザで表示されたURL（通常は http://localhost:8501 ）にアクセスします。

3. 「ファイルを選択」ボタンをクリックしてExcelファイルをアップロードします。

4. 変換結果を確認し、必要に応じて「変換済みファイルをダウンロード」ボタンをクリックします。

注意点:
- `.xlsx` ファイルは内部で `openpyxl` を使ってワークブックを直接編集するため、フォント・セル背景・数式などの書式は原則保持されます。
- `.xls`（Excel 97-2003）形式は `xlrd` を介して読み込み、変換後は新しい `.xlsx` として出力するため、元の書式情報は保持されない可能性があります。可能な限り `.xlsx` を使用してください。
- ダウンロードされるファイル名は「元のファイル名_ひらがな.xlsx」となります（例: `report.xlsx` -> `report_ひらがな.xlsx`）。

## ライセンス

MITライセンス
