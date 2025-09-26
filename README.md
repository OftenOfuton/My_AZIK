プロジェクト: My_AZIK

このプロジェクトは `設定値.xlsx` の中にあるテーブル（`Tbl_Main`）から `入力` と `出力` の列を抽出し、TSV 形式のファイルに出力します。

ディレクトリ構成:
- `src/` - ソースコード（実行: `python src/main.py`）
- `outputs/` - 生成された出力ファイル（TSV 等）
- `docs/` - 要求仕様や設計ドキュメント
- `Ignore_ExcelBackUp/` - Excel のバックアップ保存先（.gitignore 対象）

依存パッケージ: `requirements.txt` を参照してください

実行例:
```bash
python src/main.py --no-git
```

簡単な使い方:
- バックアップと出力を確認したいだけなら `--no-git` を付けて実行します。
- Git に自動で追加・コミット・プッシュしたい場合は `--no-git` を指定せず実行します（環境に git が必要です）。
