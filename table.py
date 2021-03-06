#-*- coding: utf-8 -*-
from __future__ import print_function
import datetime
from xml.etree.ElementTree import Element, tostring, register_namespace
import copy

__author__ = 'jiho'


class RowMapper:
    def __init__(self, headers, row=None, styles=None):
        self.xmlns = 'urn:schemas-microsoft-com:office:spreadsheet'
        self.columns = {}

        self.cells = None

        # 실 데이터가 넘어오지 않은 경우는 컬럼 메타에 대한 처리만 한다.
        if row:
            self.cells = row.findall("./{%s}Cell" % self.xmlns)

        # 헤더 컬럼 수와 실제 데이터 컬럼수가 같아질때까지 반복한다
        res = self.missing_cell_find()
        self.missing_cell_insert(res)

        for i, header in enumerate(headers):
            header_meta = ColumnMapper(header, styles)
            header_meta.preprocessing()

            if row:
                try:
                    column_meta = ColumnMapper(self.cells[i], styles)
                    column_meta.preprocessing()

                    self.columns[header_meta.getText()] = column_meta
                except IndexError:
                    pass

    def __getitem__(self, column_name):
        if column_name in self.columns:
            return self.columns[column_name]
        else:
            print(self.columns.keys())
            raise Exception("'%s'에 해당하는 컬럼이 없습니다." % column_name)

    def keys(self):
        return self.columns

    def emptyCreateColumn(self):
        cell_column = Element('{{{0}}}Cell'.format(self.xmlns))

        data_column = Element('{{{0}}}Data'.format(self.xmlns))
        data_column.attrib["{{{0}}}Type".format(self.xmlns)] = "String"

        cell_column.append(data_column)

        return cell_column

    def missing_cell_find(self):
        missing_cell_indexs = []
        last_column_no = 0

        for i, cell in enumerate(self.cells):
            cell_index = cell.get("{%s}Index" % self.xmlns)
            if cell_index:
                missing_cell_indexs.append((last_column_no, int(cell_index)))
                last_column_no = int(cell_index)
            else:
                last_column_no += 1

        return missing_cell_indexs

    def missing_cell_insert(self, missing_cell_info):
        # 차라리 빈 Cell을 생성해넣는게 빠르겠다.
        for entry in missing_cell_info:
            for cell_no in range(entry[0], entry[1] - 1):
                self.cells.insert(cell_no, self.emptyCreateColumn())


class ColumnMapper:
    def __init__(self, cell, style_dict):
        self.namespaces = {
            'ns': 'urn:schemas-microsoft-com:office:spreadsheet',
            'ss': 'urn:schemas-microsoft-com:office:spreadsheet'
        }

        # 특정 걸럼이 날짜 데이터 타입인지 저장해둔다.
        self.is_date = False
        self.styles = style_dict
        self.cell = cell
        self.cell_data = None
        self.style_data = None

    def preprocessing(self):
        # styleID 속성값을 가져온다. 없는 경우 None을 반환한다.
        cell_style_id = self.cell.get("{%s}StyleID" % self.namespaces['ss'])

        if cell_style_id and self.styles:
            self.style_data = self.styles[cell_style_id]

        # 실 데이터를 가져온다(String과 DateTime)
        cell_data = list(self.cell)
        cell_data = cell_data[0] if len(cell_data) > 0 else Element('{%s}Data' % self.namespaces['ns'])

        if cell_data.tag:
            # Type 속성값을 가져온다 없는 경우 None을 반환한다
            cell_data_type = cell_data.get("{%s}Type" % self.namespaces['ss'])
            if cell_data_type == "DateTime":
                self.is_date = True

            if self.is_date:
                self.cell_data = datetime.datetime.strptime(cell_data.text, "%Y-%m-%dT%H:%M:%S.%f")
            else:
                self.cell_data = cell_data.text
        else:

            self.cell_data = ""

    def getText(self):
        return self.cell_data

    @property
    def style(self):
        return self.style