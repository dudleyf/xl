Xl
==

This is a library for reading and writing Excel OOXML files, based on [openpyxl][opxl].
It's a pretty direct port, so much of it is still pythonic. I expect the interface
(and implementation) will change drastically before it's all said and done.
It's brand new, so it's likely buggy, possibly unusable, and certainly incomplete.
Also it has a stupid name.

Usage
-----
    require 'xl'

    wb = Xl::Workbook.new
    ws = wb.create_sheet
    ws.cell('A1').value = 'Hello, World!'
    Xl::Xml.save_workbook(wb, 'hello.xlsx')

[opxl]: http://bitbucket.org/ericgazoni/openpyxl/

