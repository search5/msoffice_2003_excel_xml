#-*- coding: utf-8 -*-
__author__ = 'jiho'

class RowMapper:
    def __init__(self, headers, row = None, styles = None):
        self.xmlns = 'urn:schemas-microsoft-com:office:spreadsheet'
        self.columns = {}

        if row != None: # 실 데이터가 넘어오지 않은 경우는 컬럼 메타에 대한 처리만 한다.
            cells = row.findall("./{%s}Cell" % self.xmlns)

        for i, header in enumerate(headers):
            header_meta = ColumnMapper(header, styles)

            if row != None:
                column_meta = ColumnMapper(cells[i], styles)

                self.columns[header_meta.getText()] = column_meta

    def __getitem__(self, column_name):
        if column_name in self.columns:
            return self.columns[column_name]
        else:
            raise Exception("'%s'에 해당하는 컬럼이 없습니다.")

    def keys(self):
        return self.columns

class ColumnMapper:
    def __init__(self, cell, style):
        self.namespaces = {
            'ns': 'urn:schemas-microsoft-com:office:spreadsheet',
            'ss': 'urn:schemas-microsoft-com:office:spreadsheet'
        }

        # styleID 속성값을 가져온다. 없는 경우 None을 반환한다.
        cell_style_id = cell.get("{%s}StyleID" % self.namespaces['ss'])
        if cell_style_id != None and style != None:
            self.style = style[cell_style_id]

        # 실 데이터를 가져온다(문자열)
        cell_data = cell.find("./{%s}Data" % self.namespaces['ns'])
        if cell_data != None:
            self.cell_data = cell_data.text
        else:
            self.cell_data = ""

    def getText(self):
        return self.cell_data

    @property
    def style(self):
        return self.style