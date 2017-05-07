#-*- coding: utf-8 -*-
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
from style import Style
from table import RowMapper

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

    # Worksheet 선택 메서드
    def setWorksheet(self, sheet_name):
        self.sheet = None

        for worksheet in self.xml.findall("./{%s}Worksheet" % (self.xmlns)):
            if worksheet.get('{%s}Name' % self.namespaces['ss']) == sheet_name:
                self.sheet = worksheet

        if self.sheet == None:
            raise Exception('%s 이름을 가진 워크시트가 없습니다' % sheet_name)

    # Style 구축 메서드
    def buildStyles(self):
        styles = self.xml.find("./{%s}Styles" % self.xmlns)
        for style in styles.findall('./{%s}Style' % self.xmlns):
            objStyle = Style(style)

            self.styles[objStyle.styleID] = objStyle

    # 로우 정보 구축하기
    def buildRows(self):
        table = self.sheet.find("./{%s}Table" % self.xmlns)
        rows = table.findall("./{%s}Row" % self.xmlns)

        build_rows = []

        # 헤더라인을 미리 가져온다(컬럼명 라인 - 사용자가 입력한)
        row_header = rows[0]

        for i, row in enumerate(rows[2:]):
            data_row_meta = RowMapper(row_header, row, self.styles)
            build_rows.append(data_row_meta)

        self.column_header = build_rows[0].keys()

        return build_rows

    def getColumnHeader(self):
        return self.column_header.keys()


if __name__ == "__main__":
    obj = ExcelLib(u'제5차국내외문서목록(20130917)테스트.xml')
    obj.setWorksheet(u'2011 국내문서-낱장')
    obj.buildStyles()
    obj.buildRows()
    print obj.getColumnHeader()