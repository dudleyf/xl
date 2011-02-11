Xl
==

A library for reading and writing Excel OOXML files based on [openpyxl][opxl].
It was a quick hack, and I'm not currently using it. I may or may not start working on it again someday.
It's likely buggy, possibly unusable, and certainly incomplete.
Also it has a stupid name.

Usage
-----
    require 'xl'

    wb = Xl::Workbook.new
    ws = wb.create_sheet
    ws.cell('A1').value = 'Hello, World!'
    Xl::Xml.save_workbook(wb, 'hello.xlsx')

[opxl]: http://bitbucket.org/ericgazoni/openpyxl/

