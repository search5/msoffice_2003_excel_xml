msoffice_2003_excel_xml
=======================

Microsoft Office 2003 Excel XML Reading Library

제가 업무에 쓰기 위해서 간략히 개발한 라이브러리입니다.


Examples:
<pre><code>
    >>> obj = ExcelLib(u'제5차국내외문서목록(20130917)테스트.xml')
    >>> obj.setWorksheet(u'2011 국내문서-낱장')
    >>> obj.buildStyles()
    >>> rows = obj.buildRows()

    >>> print rows[0]['기록철'].style.getUnderline()
</code></pre>


Excel 2003-2004 XML Data Structure
==================================
<pre>
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
</pre>
