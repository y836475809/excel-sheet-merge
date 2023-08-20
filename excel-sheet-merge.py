# coding=utf-8
import os
import subprocess
import sys
from typing import List

import openpyxl

import em_util


def parse_patch(lines: List[str]):
    commands = []
    sheet_name = lines[0].split("/")[-1].replace(".csv", "")
    diffs = lines[4:]
    for i, d in enumerate(diffs):
        if d.startswith("@@ "):
            info = d.replace("@@ ", "").replace(" @@", "")
            try:
                cmd = parse_line_info(info)
            except:
                raise Exception(f"Parse error: {sheet_name}, {d}")
            cmd["sheetname"] = sheet_name
            if cmd["type"] == "set":
                r = int(cmd["range"])
                ch = [x[1:] for x in diffs[i+1+r:i+1+r*2]]
                cmd["data"] = ch
                commands.append(cmd)
            if cmd["type"] == "add" or cmd["type"] == "delete":
                r = int(cmd["range"])
                ch = [x[1:] for x in diffs[i+1:i+1+r]]
                cmd["data"] = ch
                commands.append(cmd)
    return commands


def parse_line_info(line_info: str):
    diff = line_info.replace("-", "").replace("+", "").split(" ")
    pre = diff[0]
    post = diff[1]
    if pre.find(",") == -1 and post.find(",") == -1:
        # changeval
        assert pre == post
        return {
            "type": "set",
            "line1": pre,
            "line2": post,
            "range": 1,
            "data": [],
            "sheetname": ""
        }
    if pre.find(",") != -1 and post.find(",") != -1:
        # addcol, delcol, changeval2
        r1 = pre.split(",")
        r2 = post.split(",")
        if r1[1] == "0":
            return {
                "type": "add",
                "line1": r1[0],
                "line2": r2[0],
                "range": r2[1],
                "data": [],
                "sheetname": ""
            }
        elif r2[1] == "0":
            return {
                "type": "delete",
                "line1": r1[0],
                "line2": r2[0],
                "range": r1[1],
                "data": [],
                "sheetname": ""
            }
        else:
            # assert pre == post
            # assert r1[1] == r2[1]
            return {
                "type": "set",
                "line1": r1[0],
                "line2": r2[0],
                "range": r1[1],
                "data": [],
                "sheetname": ""
            }
    if pre.find(",") != -1 and post.find(",") == -1:
        # addrow
        r1 = pre.split(",")
        return {
            "type": "add",
            "line1": r1[0],
            "line2": post,
            "range": 1,
            "data": [],
            "sheetname": ""
        }
    if pre.find(",") == -1 and post.find(",") != -1:
        # delrow
        r2 = post.split(",")
        return {
            "type": "delete",
            "line1": pre,
            "line2": r2[0],
            "range": 1,
            "data": [],
            "sheetname": ""
        }
    raise Exception


def sheet_merge(wb, cmds: List):
    cmds.reverse()
    sheetname = cmds[0]["sheetname"]
    if sheetname in wb.sheetnames:
        ws = wb[sheetname]
    else:
        wb.create_sheet(index=0, title=sheetname)
        ws = wb[sheetname]

    row_offset = em_util.get_row_offset(ws)
    for c in cmds:
        if c["type"] == "delete":
            print(f"merge: delete {sheetname}")
            row = int(c["line1"]) + row_offset
            for ri in reversed(range(int(c["range"]))):
                ws.delete_rows(row + ri)
        if c["type"] == "add":
            print(f"merge: add {sheetname}")
            row = int(c["line1"]) + 1 + row_offset
            for ri in range(int(c["range"])):
                ws.insert_rows(row + ri)
            datalist = c["data"]
            for data in datalist:
                values = data.split(",")
                for ci, v in enumerate(values):
                    c1 = ws.cell(row=row, column=ci + 1)
                    if em_util.isint(v):
                        c1.value = int(v)
                    elif em_util.isfloat(v):
                        c1.value = float(v)
                    else:
                        c1.value = v
                row = row + 1
        if c["type"] == "set":
            print(f"merge: set {sheetname}")
            row = int(c["line1"]) + row_offset
            datalist = c["data"]
            for data in datalist:
                for cell in ws[row]:
                    cell.value = None
                values = data.split(",")
                for ci, v in enumerate(values):
                    c1 = ws.cell(row=row, column=ci + 1)
                    if em_util.isint(v):
                        c1.value = int(v)
                    elif em_util.isfloat(v):
                        c1.value = float(v)
                    else:
                        c1.value = v
                row = row + 1


def get_diff_unified(cached: str) -> List:
    cmd1 = f"git diff {cached} --unified=0 --name-only -- *.csv"
    output_names = subprocess.run(cmd1, capture_output=True, text=True).stdout
    if output_names == "":
        return []

    files = filter(lambda x: x != "", output_names.split("\n"))
    patches = []
    for fn in files:
        cmd2 = f"git diff {cached} --unified=0 {fn}"
        output_str = subprocess.run(cmd2, capture_output=True, text=True).stdout
        patches.append(output_str.split("\n"))
    return patches


def main():
    args = sys.argv
    if len(args) < 2:
        print(f"invalid argument")
        print(f"usage: {os.path.basename(args[0])} filepath ['staged']")
        sys.exit()

    excel_filepath = ""
    staged = ""
    if len(args) == 2:
        excel_filepath = args[1]
    if len(args) == 3 and args[2] == "staged":
        excel_filepath = args[1]
        staged = "--cached"

    workbook = openpyxl.load_workbook(excel_filepath, keep_vba=True)
    patches = get_diff_unified(staged)
    if len(patches) == 0:
        print("Not find diff, exit")
        sys.exit()

    for pt in patches:
        sheet_merge(workbook, parse_patch(pt))

    dir_name = os.path.dirname(excel_filepath)
    filename = os.path.splitext(os.path.basename(excel_filepath))[0]
    _, ext = os.path.splitext(excel_filepath)
    merged_filepath = os.path.join(dir_name, f"{filename}-merged{ext}")
    workbook.save(merged_filepath)
    print(f"save merged file: {merged_filepath}")


if __name__ == "__main__":
    main()
