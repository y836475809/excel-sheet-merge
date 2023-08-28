# coding=utf-8
import os
import sys

import util
from merge_excel_sheet import MergeExcelSheet


def main(excel_filepath: str, staged: bool, row_start_not_empty: bool):
    mes = MergeExcelSheet(excel_filepath, staged, row_start_not_empty)
    mes.merge()
    mes.save_merged_file()


if __name__ == "__main__":
    args = sys.argv
    if len(args) < 4:
        print(f"invalid argument")
        print(f"usage: {os.path.basename(args[0])} excel_filepath staged row_start")
        print(f"staged: staged or no-staged")
        print(f"row_start: row1 or row_not_empty")
        sys.exit()

    if not util.git_core_quotepath_is_false():
        print(f"Need to add 'quotepath=false' in core section of .gitconfig")
        sys.exit()

    target_filepath = args[1]
    if not os.path.exists(target_filepath):
        print(f"Not exist {target_filepath}")
        sys.exit()

    staged = True
    staged_arg = args[2]
    if staged_arg == "staged":
        staged = True
    elif staged_arg == "no-staged":
        staged = False
    else:
        print(f"{staged_arg} is not staged or no-staged")
        sys.exit()

    row_start_not_empty = True
    row_start_arg = args[3]
    if row_start_arg == "row1":
        row_start_not_empty = False
    elif row_start_arg == "row_not_empty":
        row_start_not_empty = True
    else:
        print(f"{row_start_arg} is not row1 or row_not_empty")
        sys.exit()

    main(target_filepath, staged, row_start_not_empty)
