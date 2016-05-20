# -*- coding: utf-8 -*-
import os
import tempfile
from unittest import TestCase

from os import unlink

import datetime

from bunch import Bunch
from openpyxl import load_workbook


from classy_xlsx.columns import (
    FloatXlsxColumn, IntegerXlsxColumn, PercentXlsxColumn, RatioColumn, WeightedAverageColumn,
    TextXlsxColumn, UnicodeXlsxColumn,
    DateXlsxColumn, DateTimeXlsxColumn
)
from classy_xlsx.regions import XlsxTable
from classy_xlsx.workbook import XlsxWorkbook
from classy_xlsx.worksheet import XlsxSheet


class TestRatioColumn(RatioColumn):
    dividend = 'int_col'
    divisor = 'float_col'


class TestWeightedAverageColumn(WeightedAverageColumn):
    divisor = 'int_col'


class TestXlsxTable(XlsxTable):
    int_col = IntegerXlsxColumn(title=u'')
    float_col = FloatXlsxColumn(title=u'')
    percent_col = PercentXlsxColumn(title=u'')
    ratio_col = TestRatioColumn(title=u'')
    wa_col = TestWeightedAverageColumn(title=u'')
    text_col = TextXlsxColumn(title=u'')
    uni_col = UnicodeXlsxColumn(title=u'')
    date_col = DateXlsxColumn(title=u'')
    datetime_col = DateTimeXlsxColumn(title=u'')

    def get_queryset(self):
        return [
            Bunch(
                int_col=1,
                float_col=10.1,
                percent_col=0.3,
                wa_col=100,
                text_col='some text',
                uni_col=u'какой-то текст',
                date_col=datetime.date(year=2016, month=01, day=01),
                datetime_col=datetime.datetime(year=2016, month=01, day=01, hour=12),
            ),
            Bunch(
                int_col=2,
                float_col=20.2,
                percent_col=0.5,
                wa_col=50,
                text_col='some text',
                uni_col=u'какой-то текст',
                date_col=datetime.date(year=2016, month=01, day=01),
                datetime_col=datetime.datetime(year=2016, month=01, day=01, hour=12),
            ),
        ]


class TestXlsxSheet(XlsxSheet):
    table = TestXlsxTable()


class TestXlsxWorkbook(XlsxWorkbook):
    sheet = TestXlsxSheet()


class ContextTest(TestCase):
    def setUp(self):
        ff, self.out_file_name = tempfile.mkstemp(suffix='.xlsx')
        os.close(ff)
        self.wb = TestXlsxWorkbook(file_name=self.out_file_name)

    def test_all(self):
        self.wb.make_report()
        wb2 = load_workbook(self.wb.file_name)
        self.assertIn('sheet', wb2.get_sheet_names())
        sheet = wb2.get_sheet_by_name('sheet')
        print sheet


    # def tearDown(self):
    #     unlink(self.out_file_name)
