#-*- coding: utf-8 -*-
from __future__ import print_function
from __future__ import absolute_import

__author__ = 'jiho'

"""
Excel 2003-2004 XML Data Structure

Workbook/
    Styles/
        Style/
            @ss:ID
            Font/
                @ss:Underline = "Single"
    Worksheet
        @ss:Name
        Table/
            Row/
                Cell/
                    @ss:StyleID
                    Data
"""

import xml.etree.cElementTree as ET
from .style import Style
from .table import RowMapper


"""
# 최종목적은 row[0]['column'].style.getUnderline()과 같은 형태를 원하는 것임..

Examples:
    >>> obj = ExcelLib(u'제5차국내외문서목록(20130917)테스트.xml')
    >>> obj.setWorksheet(u'2011 국내문서-낱장')
    >>> obj.buildStyles()
    >>> rows = obj.buildRows()

    >>> print rows[0]['기록철'].style.getUnderline()
"""


class ExcelLib:
    def __init__(self, parse_xml_name):
        self.xml = ET.parse(parse_xml_name)
        self.xmlns = 'urn:schemas-microsoft-com:office:spreadsheet'
        self.namespaces = {
            'ss': 'urn:schemas-microsoft-com:office:spreadsheet'
        }
        self.styles = {}
        self.rows = []
        self.column_header = None
        self.sheet = None
        self.header_line = 0
        self.content_line = 1

    def set_header_line(self, num=0):
        self.header_line = num

    def set_content_line(self, num=1):
        self.content_line = num

    # Worksheet 선택 메서드
    def set_worksheet(self, sheet_name):
        for worksheet in self.xml.findall("./{%s}Worksheet" % self.xmlns):
            if worksheet.get('{%s}Name' % self.namespaces['ss']) == sheet_name:
                self.sheet = worksheet

        if not self.sheet:
            raise Exception('%s 이름을 가진 워크시트가 없습니다' % sheet_name)

    def get_sheets(self):
        sheets = [worksheet.get('{%s}Name' % self.namespaces['ss'])
                  for worksheet in self.xml.findall("./{%s}Worksheet" % self.xmlns)]
        return sheets

    # Style 구축 메서드
    def build_styles(self):
        styles = self.xml.find("./{%s}Styles" % self.xmlns)

        for style in styles.findall('./{%s}Style' % self.xmlns):
            obj_style = Style(style)

            self.styles[obj_style.styleID] = obj_style

    # 로우 정보 구축하기
    def build_rows(self):
        if self.header_line < -1:
            raise ExcelLibException('헤더 라인 번호가 입력되어 있지 않습니다')

        if not self.content_line:
            raise ExcelLibException('콘텐츠 라인 시작 번호가 입력되어 있지 않습니다')

        table = self.sheet.find("./{%s}Table" % self.xmlns)
        rows = table.findall("./{%s}Row" % self.xmlns)

        build_rows = []

        # 헤더라인을 미리 가져온다(컬럼명 라인 - 사용자가 입력한)
        row_header = rows[self.header_line]

        for i, row in enumerate(rows[self.content_line:]):
            data_row_meta = RowMapper(row_header, row, self.styles)
            build_rows.append(data_row_meta)

        self.column_header = tuple(build_rows[self.header_line].keys())

        return build_rows

    def get_column_header(self):
        return self.column_header.keys()


class ExcelLibException(Exception):
    pass


if __name__ == "__main__":
    obj = ExcelLib(u'제5차국내외문서목록(20130917)테스트.xml')
    obj.set_worksheet(u'2011 국내문서-낱장')
    obj.build_styles()
    obj.build_rows()
    print(obj.get_column_header())
