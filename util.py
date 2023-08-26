# coding=utf-8


def isint(s):
    try:
        int(s, 10)
    except ValueError:
        return False
    else:
        return True


def isfloat(s):
    try:
        float(s)
    except ValueError:
        return False
    else:
        return True


def get_row_offset(ws):
    offset = 0
    for row in ws:
        if is_row_empty(row):
            offset = offset + 1
        else:
            return offset
    return offset


def is_row_empty(row):
    for cell in row:
        if cell.value is not None:
            return False
    return True

