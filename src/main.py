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
    os.makedirs(backup_dir, exist_ok=True)
    now = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = os.path.join(backup_dir, f'{now}_{os.path.basename(excel_file)}')
    shutil.copy2(excel_file, backup_path)
    print(f'バックアップ作成: {backup_path}')
    return backup_path


def extract_table(excel_file: str, table_name: str, columns: list) -> pd.DataFrame:
    wb = openpyxl.load_workbook(excel_file, data_only=True)
    for ws in wb.worksheets:
        for tbl in getattr(ws, "_tables", []):
            if not hasattr(tbl, 'name'):
                continue
            if tbl.name == table_name:
                ref = tbl.ref
                cells = list(ws[ref])
                if not cells:
                    continue
                headers = [str(cell.value) if cell.value is not None else "" for cell in cells[0]]
                df_rows = []
                for row in cells[1:]:
                    df_rows.append([cell.value for cell in row])
                try:
                    df_tbl = pd.DataFrame(df_rows, columns=headers)
                except Exception:
                    continue
                if set(columns).issubset(df_tbl.columns):
                    result = df_tbl[columns].copy()
                    try:
                        header_mask = pd.DataFrame({c: result[c].astype(str).fillna('') == c for c in columns}).all(axis=1)
                        result = result.loc[~header_mask]
                    except Exception:
                        pass
                    result = result.dropna(how='all')
                    return result

    sheets = pd.read_excel(excel_file, sheet_name=None, engine='openpyxl')
    for data in sheets.values():
        if set(columns).issubset(data.columns):
            result = data[columns].copy()
            try:
                header_mask = pd.DataFrame({c: result[c].astype(str).fillna('') == c for c in columns}).all(axis=1)
                result = result.loc[~header_mask]
            except Exception:
                pass
            result = result.dropna(how='all')
            return result

    raise RuntimeError(f"テーブル '{table_name}' または列 {columns} が見つかりませんでした。")


def write_tsv(df: pd.DataFrame, out_file: str) -> None:
    if os.name == 'nt':
        encoding = 'utf-8-sig'
        lineterm = '\r\n'
    else:
        encoding = 'utf-8'
        lineterm = '\n'

    df.to_csv(out_file, sep='\t', header=False, index=False, encoding=encoding, lineterminator=lineterm)
    print(f'TSV出力: {out_file}')


def git_commit_push(commit_message: str) -> None:
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
