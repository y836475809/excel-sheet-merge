# coding=utf-8
import os
import sys

from merge_excel_sheet import MergeExcelSheet


def main(excel_filepath: str, staged: bool):
    mes = MergeExcelSheet(excel_filepath, staged)
    mes.merge()
    mes.save_merged_file()


if __name__ == "__main__":
    args = sys.argv
    if len(args) < 2:
        print(f"invalid argument")
        print(f"usage: {os.path.basename(args[0])} excel_filepath [no-staged]")
        sys.exit()

    if len(args) < 3:
        main(args[1], True)
    elif len(args) < 4:
        main(args[1], False)
