プロジェクト: My_AZIK

このプロジェクトは `設定値.xlsx` の中にあるテーブル（`Tbl_Main`）から `入力` と `出力` の列を抽出し、TSV 形式のファイルに出力します。

ディレクトリ構成:
- `src/` - ソースコード（実行: `python src/main.py`）
- `inputs/` - 入力ファイル（例: `inputs/設定値.xlsx`）
- `outputs/` - 生成された一時的な出力（開発中）
- `docs/` - 要求仕様や設計ドキュメント
- `Ignore_ExcelBackUp/` - Excel のバックアップ保存先（.gitignore 対象）

依存パッケージ: `requirements.txt` を参照してください

実行例:
```bash
# 仮想環境を有効にしてから実行する例
source venv/bin/activate
python src/main.py --no-git

# オプション例: 明示的に入力/出力を指定
python src/main.py --excel inputs/設定値.xlsx --output MyRomanTable.txt

# Git の自動コミット・プッシュを行う（環境に git が必要）
python src/main.py --excel inputs/設定値.xlsx --output MyRomanTable.txt
```

簡単な使い方:
- バックアップと出力を確認したいだけなら `--no-git` を付けて実行します。
- Git に自動で追加・コミット・プッシュしたい場合は `--no-git` を指定せず実行します（環境に git が必要です）。
