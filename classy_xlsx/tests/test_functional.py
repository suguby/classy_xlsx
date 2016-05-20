# -*- coding: utf-8 -*-
import tempfile
from unittest import TestCase

from os import unlink

from classy_xlsx.columns import (
    FloatXlsxColumn, IntegerXlsxColumn, PercentXlsxColumn, RatioColumn, WeightedAverageColumn,
    TextXlsxColumn, UnicodeXlsxColumn,
    DateXlsxColumn, DateTimeXlsxColumn
)
from classy_xlsx.regions import XlsxTable
from classy_xlsx.utils import Bunch
from classy_xlsx.workbook import XlsxWorkbook
from classy_xlsx.worksheet import XlsxSheet


class TestXlsxTable(XlsxTable):
    int_col = IntegerXlsxColumn()
    float_col = FloatXlsxColumn()
    percent_col = PercentXlsxColumn()
    ratio_col = RatioColumn()
    wa_col = WeightedAverageColumn()
    text_col = TextXlsxColumn()
    uni_col = UnicodeXlsxColumn()
    date_col = DateXlsxColumn()
    datetime_col = DateTimeXlsxColumn()

    def get_queryset(self):
        return [
            Bunch(
                int_col=1,
                float_col=10.1,
                percent_col=0.3,
                text_col = 'some text',
                uni_col = u'какой-то текст',
                date_col = '11-11-2016',
                datetime_col = '11-11-2016 11:11:11',
            ),
            Bunch(
                int_col=2,
                float_col=20.2,
                percent_col=0.5,
                text_col='some text',
                uni_col=u'какой-то текст',
                date_col='11-11-2016',
                datetime_col='11-11-2016 11:11:11',
            ),
        ]


class TestXlsxSheet(XlsxSheet):
    table = TestXlsxTable()


class TestXlsxWorkbook(XlsxWorkbook):
    sheet = TestXlsxSheet()


class ContextTest(TestCase):

    def setUp(self):
        self.out_file = tempfile.NamedTemporaryFile(delete=False)
        self.wb = TestXlsxWorkbook(file_name=self.out_file.name)

    def test_all(self):
        self.wb.make_report()

    def tearDown(self):
        self.out_file.close()
        unlink(self.out_file.name)
