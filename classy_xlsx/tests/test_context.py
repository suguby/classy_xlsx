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

    def get_extra_context(self):
        return dict(param3=3)


class TestXlsxSheet(XlsxSheet):
    table = TestXlsxTable()


class TestXlsxWorkbook(XlsxWorkbook):
    sheet = TestXlsxSheet()


class Test2XlsxWorkbook(XlsxWorkbook):
    default_context = dict(param2=2)
    sheet = TestXlsxSheet()


class ContextTest(TestCase):

    def test_inheritance(self):
        self.wb = TestXlsxWorkbook(context=dict(param1=1))
        self.assertEquals(self.wb.sheet.table.col.context.param1, 1)

    def test_extra_context(self):
        self.wb = Test2XlsxWorkbook()
        self.assertEquals(self.wb.sheet.table.col.context.param2, 2)

    def test_get_extra_context(self):
        self.wb = TestXlsxWorkbook()
        self.assertEquals(self.wb.sheet.table.col.context.param3, 3)
