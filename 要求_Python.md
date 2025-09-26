## やりたいこと
Excelブックにあるテーブルからデータを取り出し，TSV形式のテキストファイルに出力する．
そのあと，シェルを呼び出してGitのステージング，コミット，プッシュを行う．

## 前提
MacとWindows両方で使いたい．
共通のコードで実行できれば望ましいが，無理であれば別のファイルになっても良い．

## Excelブックの仕様
ブック上に「Tbl_Main」というテーブルがある．
Tbl_Mainには，「入力」「出力」という列があり，さらにいくつかの列がある．

## ファイル・フォルダ構成
```
./ -- Gitリポジトリ
├─ venv/ -- Python仮想環境 (gitignoreの対象)
├─ 設定値.xlsx -- データを取り出したいExcelファイル
├─ main.py -- メインプログラム
├─ MyRomanTable.txt -- 出力先ファイル
└─ Ignore_ExcelBackUp/ -- 設定値.xlsxのバックアップ．gitignoreの対象．

<!-- └─ AZIK-KeyMap_GoogleJPInput/ -- これがGitリポジトリ
    ├─ GoogleJPInput_Settings.md
    ├─ README.md
    └─ RomanTable.txt -- 取り出したデータはここに上書きしたい -->
```

## 取り出したいデータについて
Excelブックの Tbl_Main テーブルから，「入力」「出力」列だけをタブ区切りで取り出す．列見出しは不要．

## 補遺
Excelファイルを読み出す前にバックアップを取る
バックアップは Ignore_ExcelBackUp フォルダに保存し、過去のバックアップは保持します（削除しません）。
