require File.join(File.dirname(__FILE__), '../lib/xl')

wb = Xl::Workbook.new
ws = wb.create_sheet
ws.title = "Hello, World!"
ws.cell('A1').value = 'Hello'
ws.cell('A2').value = 'World'
Xl::Xml.save_workbook(wb, 'hello.xlsx')
