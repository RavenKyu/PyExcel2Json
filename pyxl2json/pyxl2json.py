#!/usr/bin/env python
# coding=utf8
"""
====================================
 :mod: PyExcel2Json
====================================
.. moduleauthor:: Raven Lim <hong18s@gmail.com>
.. note::
"""
import openpyxl
import json

def _open_excel_file(filename) -> openpyxl.Workbook:
    return openpyxl.load_workbook(filename)

def _read_sheet(wb:openpyxl.Workbook, sheet:str):
    return wb[sheet]

def _read_head(ws:openpyxl.worksheet.Worksheet, data_range):
    return [x.value for x in ws[data_range] for x in x]

def _read_data(ws:openpyxl.worksheet.Worksheet, data_range):
    data = []
    for row in ws[data_range]:
        cells = []
        for cell in row:
            cells.append(cell.value)
        data.append(cells)
    return data

def _to_dict(head: list, data: list):
    r = []
    for d in data:
        r.append(dict(zip(head, d)))
    return r

def _to_json(data: list):
    return json.dumps(data, indent=4)

def main(argspec):
    wb = _open_excel_file(argspec.excel_filename)
    ws = _read_sheet(wb, argspec.sheet)
    head = _read_head(ws, argspec.head)
    data = _read_data(ws, argspec.data)

    d = _to_dict(head, data)
    print(_to_json(d))
