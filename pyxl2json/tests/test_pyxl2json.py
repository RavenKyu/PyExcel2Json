#!/usr/bin/env python
# coding=utf8

"""
====================================
 :mod: Test case for Config Module
====================================
.. module author:: 임덕규 <hong18s@gmail.com>
.. note:: MIT License
"""

import unittest
import openpyxl
from pyxl2json import parse_args
from pyxl2json.pyxl2json import \
    _open_excel_file, _read_sheet, _read_head, _read_data, _to_dict, _to_json,\
    excel_column_calculate

from pyxl2json import ArgsError


################################################################################
class TestUnit (unittest.TestCase):
    """
    Test Unit for Locate Image
    """
    _INPUT_FILENAME_1 = 'test_excel_data_1.xlsx'
    _INPUT_FILENAME_2 = 'test_excel_data_2.xlsx'
    _SHEET_NAME_1 = "Sheet1"
    _HEAD_RANGE = "A1:D1"
    _DATA_RANGE = "A2:D5"

    # ==========================================================================
    def setUp(self):
        pass

    # ==========================================================================
    def tearDown(self):
        pass

    # ==========================================================================
    def test_010_argument_filename(self):
        # 단일 파일일 경우
        args = [self._INPUT_FILENAME_1]
        argspec = parse_args(args)
        self.assertEqual(argspec.excel_filename, [self._INPUT_FILENAME_1])

        # 다수 파일일 경우
        args = [self._INPUT_FILENAME_1, self._INPUT_FILENAME_2]
        argspec = parse_args(args)
        self.assertEqual(argspec.excel_filename,
                         [self._INPUT_FILENAME_1, self._INPUT_FILENAME_2])

    # ==========================================================================
    def test_011_argument_head(self):
        _COLUMN_RANGE = 'A1:F1'
        _WRONG_COLUMN_RANGE = 'A1F1'

        # --head 옵션이 없을 때의 값은 None
        args = [self._INPUT_FILENAME_1,]
        argspec = parse_args(args)
        self.assertEqual(argspec.head, None)

        args = [self._INPUT_FILENAME_1, '--head', _COLUMN_RANGE]
        argspec = parse_args(args)
        self.assertEqual(argspec.head, _COLUMN_RANGE)

        args = [self._INPUT_FILENAME_1, '--head', _WRONG_COLUMN_RANGE]
        self.assertRaises(ArgsError, lambda: parse_args(args))


    # ==========================================================================
    def test_012_argument_data_range(self):
        _COLUMN_RANGE = 'A1:F1'
        _WRONG_COLUMN_RANGE = 'A1F1'

        args = [self._INPUT_FILENAME_1, '--data', _COLUMN_RANGE]
        argspec = parse_args(args)
        self.assertEqual(_COLUMN_RANGE, argspec.data)

        args = [self._INPUT_FILENAME_1, '--data', _WRONG_COLUMN_RANGE]
        self.assertRaises(ArgsError, lambda: parse_args(args))



    # ==========================================================================
    def test_020_open_excel_file(self):
        args = [self._INPUT_FILENAME_1,]
        argspec = parse_args(args)
        wb = _open_excel_file(argspec.excel_filename[0])
        self.assertIsInstance(wb, openpyxl.Workbook)

    # ==========================================================================
    def test_013_argument_sheet_name(self):
        _WRONG_SHEET_NAME = "Sheet 2"

        args = [self._INPUT_FILENAME_1, '--sheet', self._SHEET_NAME_1]
        argspec = parse_args(args)
        wb = _open_excel_file(argspec.excel_filename[0])
        ws = _read_sheet(wb, argspec.sheet)
        self.assertIsInstance(ws, openpyxl.worksheet.Worksheet)

        args = [self._INPUT_FILENAME_1, '--sheet', _WRONG_SHEET_NAME]
        argspec = parse_args(args)
        wb = _open_excel_file(argspec.excel_filename[0])
        self.assertRaises(
            KeyError, lambda: _read_sheet(wb, argspec.sheet))

    # ==========================================================================
    def test_014_argument_as_file(self):
        args = [self._INPUT_FILENAME_1, '--asfile']
        argspec = parse_args(args)
        self.assertTrue(argspec.asfile)

    # ==========================================================================
    def test_015_argument_verbose(self):
        args = [self._INPUT_FILENAME_1, '--verbose']
        argspec = parse_args(args)
        self.assertTrue(argspec.verbose)

    # ==========================================================================
    def test_014_read_head(self):
        _HEAD_DATA = ["NAME", "VALUE", "COLOR", "DATE"]

        # Head 범위 값을 명시하지 않았을 때
        args = [self._INPUT_FILENAME_1, '--sheet', self._SHEET_NAME_1]
        argspec = parse_args(args)
        wb = _open_excel_file(argspec.excel_filename[0])
        ws = _read_sheet(wb, argspec.sheet)
        data = _read_head(ws, None)
        self.assertEqual(data, _HEAD_DATA)

        # Head 범위 값을 명시하였을 때
        args = [self._INPUT_FILENAME_1, '--sheet', self._SHEET_NAME_1]
        argspec = parse_args(args)
        wb = _open_excel_file(argspec.excel_filename[0])
        ws = _read_sheet(wb, argspec.sheet)
        data = _read_head(ws, self._HEAD_RANGE)
        self.assertEqual(data, _HEAD_DATA)

    # ==========================================================================
    def test_014_read_data(self):

        _DATA = [
            ['Alan', 12, 'blue', 'Sep. 25, 2009'],
            ['Shan', 13, "green	blue", 'Sep. 27, 2009'],
            ['John', 45, 'orange', 'Sep. 29, 2009'],
            ['Minna', 27, 'teal', 'Sep. 30, 2009'],]

        args = [self._INPUT_FILENAME_1, '--sheet', self._SHEET_NAME_1]
        argspec = parse_args(args)
        wb = _open_excel_file(argspec.excel_filename[0])
        ws = _read_sheet(wb, argspec.sheet)

        # 데이터 범위 값을 명시하지 않았을 때
        data = _read_data(ws, None)
        self.assertEqual(data, _DATA)

        # 데이터 범위 값을 명시하였을 때
        data = _read_data(ws, self._DATA_RANGE)
        self.assertEqual(data, _DATA)



    # ==========================================================================
    def test_015_to_dict(self):
        _DATA = [{'NAME': 'Alan', 'VALUE': 12, 'COLOR': 'blue',
          'DATE': 'Sep. 25, 2009'},
         {'NAME': 'Shan', 'VALUE': 13, 'COLOR': 'green\tblue',
          'DATE': 'Sep. 27, 2009'},
         {'NAME': 'John', 'VALUE': 45, 'COLOR': 'orange',
          'DATE': 'Sep. 29, 2009'},
         {'NAME': 'Minna', 'VALUE': 27, 'COLOR': 'teal',
          'DATE': 'Sep. 30, 2009'}]

        args = [self._INPUT_FILENAME_1, '--sheet', self._SHEET_NAME_1]
        argspec = parse_args(args)
        wb = _open_excel_file(argspec.excel_filename[0])
        ws = _read_sheet(wb, argspec.sheet)
        head = _read_head(ws, self._HEAD_RANGE)
        data = _read_data(ws, self._DATA_RANGE)
        d = _to_dict(head, data)
        self.assertEqual(_DATA, d)

    # ==========================================================================
    def test_020_excel_colum_calculate(self):
        _COLUMN_NUMBER_1 = 1
        _COLUMN_NUMBER_2 = 26
        _COLUMN_NUMBER_3 = 27
        _COLUMN_NUMBER_4 = 1000

        _COLUMN_STRING_1 = 'A'
        _COLUMN_STRING_2 = 'Z'
        _COLUMN_STRING_3 = 'AA'
        _COLUMN_STRING_4 = 'ALL'

        self.assertEqual(_COLUMN_STRING_1,
                         excel_column_calculate(_COLUMN_NUMBER_1, []))
        self.assertEqual(_COLUMN_STRING_2,
                         excel_column_calculate(_COLUMN_NUMBER_2, []))
        self.assertEqual(_COLUMN_STRING_3,
                         excel_column_calculate(_COLUMN_NUMBER_3, []))
        self.assertEqual(_COLUMN_STRING_4,
                         excel_column_calculate(_COLUMN_NUMBER_4, []))




    # def test_100_open_excel_file(self):
    #     self.assertIsInstance()


