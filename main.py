# coding=utf-8
import os
import sys

import util
from merge_excel_sheet import StartRow, MergeExcelSheet


def main(_excel_filepath: str, _staged: bool, _start_row: StartRow):
    mes = MergeExcelSheet(_excel_filepath, _staged, _start_row)
    mes.merge()
    mes.save_merged_file()


if __name__ == "__main__":
    args = sys.argv
    if len(args) < 4:
        print(f"invalid argument")
        print(f"usage: {os.path.basename(args[0])} excel_filepath staged start_row")
        print(f"staged: staged or no-staged")
        print(f"start_row: row_first or row_not_none")
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

    start_row = StartRow.First
    start_row_arg = args[3]
    if start_row_arg == "row_first":
        start_row = StartRow.First
    elif start_row_arg == "row_not_none":
        start_row = StartRow.NotNone
    else:
        print(f"{start_row_arg} is not row_first or row_not_none")
        sys.exit()

    main(target_filepath, staged, start_row)
