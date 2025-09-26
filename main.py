import argparse
import os
import shutil
import subprocess
import sys
from datetime import datetime

import openpyxl
import pandas as pd

# 設定
EXCEL_FILE = '設定値.xlsx'
BACKUP_DIR = 'Ignore_ExcelBackUp'
TSV_FILE = 'MyRomanTable.txt'
TABLE_NAME = 'Tbl_Main'
COLUMNS = ['入力', '出力']

def backup_excel(excel_file: str, backup_dir: str) -> str:
    """バックアップを作成。既存バックアップは削除してから新しいファイルを保存し、保存先パスを返す。"""
    os.makedirs(backup_dir, exist_ok=True)
    # 以前のバックアップは保持する（gitignore 対象のため削除しない）
    # バックアップ作成
    now = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = os.path.join(backup_dir, f'{now}_{os.path.basename(excel_file)}')
    shutil.copy2(excel_file, backup_path)
    print(f'バックアップ作成: {backup_path}')
    return backup_path

# Excelからデータ抽出
def extract_table(excel_file: str, table_name: str, columns: list) -> pd.DataFrame:
    """openpyxl を使ってテーブル名で検出し、見つからなければ列名でフォールバックする。

    戻り値は指定した列のみの DataFrame。
    """
    # まず openpyxl でテーブル名を探す
    wb = openpyxl.load_workbook(excel_file, data_only=True)
    for ws in wb.worksheets:
        # ワークシートのテーブル列挙
        for tbl in getattr(ws, "_tables", []):
            # 一部環境で tbl が文字列など期待しない型になることがあるため安全に扱う
            if not hasattr(tbl, 'name'):
                continue
            if tbl.name == table_name:
                # table.ref は範囲文字列（例: A1:C10）
                ref = tbl.ref
                cells = list(ws[ref])
                if not cells:
                    continue
                # 1行目をヘッダ、残りをデータとして DataFrame を作る
                headers = [str(cell.value) if cell.value is not None else "" for cell in cells[0]]
                df_rows = []
                for row in cells[1:]:
                    df_rows.append([cell.value for cell in row])
                try:
                    df_tbl = pd.DataFrame(df_rows, columns=headers)
                except Exception:
                    # 列数が不一致など
                    continue
                if set(columns).issubset(df_tbl.columns):
                    result = df_tbl[columns].copy()
                    # ヘッダ行がデータとして混入している場合を除外（各列の値が列名と一致する行）
                    try:
                        header_mask = pd.DataFrame({c: result[c].astype(str).fillna('') == c for c in columns}).all(axis=1)
                        result = result.loc[~header_mask]
                    except Exception:
                        pass
                    # 完全に空の行は除去
                    result = result.dropna(how='all')
                    return result

    # フォールバック: pandas で全シート読み込みして列名検索
    sheets = pd.read_excel(excel_file, sheet_name=None, engine='openpyxl')
    for data in sheets.values():
        if set(columns).issubset(data.columns):
            result = data[columns].copy()
            # 同様にヘッダの混入や空行を取り除く
            try:
                header_mask = pd.DataFrame({c: result[c].astype(str).fillna('') == c for c in columns}).all(axis=1)
                result = result.loc[~header_mask]
            except Exception:
                pass
            result = result.dropna(how='all')
            return result

    raise RuntimeError(f"テーブル '{table_name}' または列 {columns} が見つかりませんでした。")

# TSV出力
def write_tsv(df: pd.DataFrame, out_file: str) -> None:
    # Windows の場合は BOM 付き UTF-8 を使うと Excel で扱いやすい
    if os.name == 'nt':
        encoding = 'utf-8-sig'
        lineterm = '\r\n'
    else:
        encoding = 'utf-8'
        lineterm = '\n'

    df.to_csv(out_file, sep='\t', header=False, index=False, encoding=encoding, lineterminator=lineterm)
    print(f'TSV出力: {out_file}')

# Git操作
def git_commit_push(commit_message: str) -> None:
    # git コマンドが使えるか確認
    if shutil.which('git') is None:
        raise RuntimeError('git が見つかりません。PATH を確認してください。')

    try:
        subprocess.run(['git', 'add', '.'], check=True)
        subprocess.run(['git', 'commit', '-m', commit_message], check=True)
        subprocess.run(['git', 'push'], check=True)
        print('Gitへプッシュ完了')
    except subprocess.CalledProcessError as e:
        print('Git操作が失敗しました:', e)
        raise

def parse_args():
    p = argparse.ArgumentParser(description='Excel のテーブルから TSV を作成して Git にコミット・プッシュします')
    p.add_argument('--excel', '-e', default=EXCEL_FILE, help='入力 Excel ファイル')
    p.add_argument('--output', '-o', default=TSV_FILE, help='出力 TSV ファイル')
    p.add_argument('--backup-dir', '-b', default=BACKUP_DIR, help='バックアップ保存先ディレクトリ')
    p.add_argument('--table', '-t', default=TABLE_NAME, help='Excel テーブル名')
    p.add_argument('--no-git', action='store_true', help='Git 操作を行わない')
    return p.parse_args()


def main():
    args = parse_args()

    if not os.path.exists(args.excel):
        print(f'入力ファイルが見つかりません: {args.excel}')
        sys.exit(2)

    try:
        backup_excel(args.excel, args.backup_dir)
    except Exception as e:
        print('バックアップに失敗しました:', e)
        sys.exit(3)

    try:
        df = extract_table(args.excel, args.table, COLUMNS)
    except Exception as e:
        print('データ抽出に失敗しました:', e)
        sys.exit(4)

    try:
        write_tsv(df, args.output)
    except Exception as e:
        print('TSV 出力に失敗しました:', e)
        sys.exit(5)

    if not args.no_git:
        try:
            git_commit_push(f'Update {os.path.basename(args.output)}')
        except Exception:
            print('Git 操作でエラーが発生しました。手動で確認してください。')


if __name__ == '__main__':
    main()
