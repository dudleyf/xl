require File.join(File.dirname(__FILE__), '../lib/xl')

wb = Xl::Workbook.new
ws = wb.create_sheet

ws.cell('A1').value = 'Character'
ws.cell('A2').value = 'Shadow'
ws.cell('A3').value = 'Speedy'
ws.cell('A4').value = 'Bashful'
ws.cell('A5').value = 'Pokey'

ws.cell('B1').value = 'Nickname'
ws.cell('B2').value = 'Blinky'
ws.cell('B3').value = 'Pinky'
ws.cell('B4').value = 'Inky'
ws.cell('B5').value = 'Clyde'

ws.cell('C1').value = 'Color'
ws.cell('C2').value = 'Red'
ws.cell('C3').value = 'Pink'
ws.cell('C4').value = 'Cyan'
ws.cell('C5').value = 'Orange'

ws.cell('C2').style.fill.pattern_type = Xl::Fill::FILL_SOLID
ws.cell('C2').style.fill.fg_color = Xl::Color.new(Xl::Color::RED)

ws.cell('C3').style.fill.pattern_type = Xl::Fill::FILL_SOLID
ws.cell('C3').style.fill.fg_color = Xl::Color.new(Xl::Color::PINK)

ws.cell('C4').style.fill.pattern_type = Xl::Fill::FILL_SOLID
ws.cell('C4').style.fill.fg_color = Xl::Color.new(Xl::Color::CYAN)

ws.cell('C5').style.fill.pattern_type = Xl::Fill::FILL_SOLID
ws.cell('C5').style.fill.fg_color = Xl::Color.new(Xl::Color::ORANGE)

%w[A B C].each do |c|
  ws.cell("#{c}1").style.font.size = 14
  ws.cell("#{c}1").style.font.underline = true

  ws.cell("#{c}1").style.borders.top.border_style = Xl::Border::BORDER_THIN
  ws.cell("#{c}5").style.borders.bottom.border_style = Xl::Border::BORDER_THIN
end

1.upto(5) do |i|
  ws.cell("A#{i}").style.borders.left.border_style = Xl::Border::BORDER_THIN
  ws.cell("C#{i}").style.borders.right.border_style = Xl::Border::BORDER_THIN

  ws.cell("B#{i}").style.alignment.horizontal = Xl::Alignment::HORIZONTAL_CENTER
  ws.cell("C#{i}").style.alignment.horizontal = Xl::Alignment::HORIZONTAL_RIGHT
end

Xl::Xml.save_workbook(wb, 'styles.xlsx')
