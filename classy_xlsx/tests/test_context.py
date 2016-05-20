# -*- coding: utf-8 -*-
from unittest import TestCase

from classy_xlsx.columns import XlsxColumn
from classy_xlsx.regions import XlsxTable
from classy_xlsx.workbook import XlsxWorkbook
from classy_xlsx.worksheet import XlsxSheet


class TestXlsxColumn(XlsxColumn):
    pass


class TestXlsxTable(XlsxTable):
    col = TestXlsxColumn()


class TestXlsxSheet(XlsxSheet):
    table = TestXlsxTable()


class TestXlsxWorkbook(XlsxWorkbook):
    sheet = TestXlsxSheet()


class ContextTest(TestCase):

    def setUp(self):
        self.wb = TestXlsxWorkbook(context=dict(param1=1))

    def test_inheritance(self):
        # self.wb.make_report()
        self.assertEquals(self.wb.sheet.table.col.context.param1, 1)