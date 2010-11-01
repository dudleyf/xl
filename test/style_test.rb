require File.join(File.dirname(__FILE__), "test_helper")

class StyleTest < XlTestCase
  include Xl::Xml

  def test_extract_style_table
    wb = Xl::Workbook.new
    ws = wb.create_sheet
    ws.cell('A1').value = '12.34%'
    now = Time.now
    ws.cell('B4').value = now
    ws.cell('B5').value = now
    ws.cell('C14').value = "This is a test"
    ws.cell('D9').value = '31.31415'
    ws.cell('D9').style.number_format.format_code = Xl::NumberFormat::FORMAT_NUMBER_00

    table = Xl::Xml.extract_style_table(wb)
    assert_equal(3, table.size)
  end

  def test_write_style_table
    wb = Xl::Workbook.new
    ws = wb.create_sheet
    ws.cell('A1').value = '12.34%'
    now = Time.now
    ws.cell('B4').value = now
    ws.cell('B5').value = now
    ws.cell('C14').value = "This is a test"
    ws.cell('D9').value = '31.31415'
    ws.cell('D9').style.number_format.format_code = Xl::NumberFormat::FORMAT_NUMBER_00

    table = Xl::Xml.extract_style_table(wb)
    content = Xl::Xml.write_style_table(table)
    assert_xml_equal(test_data('writer/expected/styles_simple.xml'), content)
  end

  def test_read_style_table
    style_table = Xl::Xml.read_style_table(test_data('reader/simple-styles.xml'))
    assert_equal(4, style_table.size)
    assert_equal(Xl::NumberFormat::BUILTIN_FORMATS[9], style_table[1].number_format.format_code)
    assert_equal('yyyy-mm-dd', style_table[2].number_format.format_code)
  end

  def test_read_cell_style
    style_table = Xl::Xml.read_style_table(test_data('reader/empty-workbook-styles.xml'))
    assert_equal(2, style_table.length)
  end

  def test_styles_equal
    s1 = Xl::Style.new
    s2 = Xl::Style.new

    assert s1 == s2
    assert s1.eql?(s2)
    assert s1.hash == s2.hash
    assert [s1,s2].uniq == [s1]

    s2.font.size = 20

    assert s1 != s2
    assert !s1.eql?(s2)
    assert s1.hash != s2.hash
    assert [s1,s2].uniq == [s1,s2]
  end

  def test_write_style_table_fonts
    wb = Xl::Workbook.new
    ws = wb.create_sheet

    ws.cell('E3').style.font = Xl::Font.new({
      :name => "Times New Roman",
      :size => 33,
      :bold => true,
      :italic => true,
      :superscript => true,
      :outline => true,
      :shadow => true,
      :underline => Xl::Font::UNDERLINE_SINGLE,
      :strikethrough => true,
      :color => Xl::Color.new(Xl::Color::RED)
    })

    table = Xl::Xml.extract_style_table(wb)
    content = Xl::Xml.write_style_table(table)
    assert_xml_equal test_data('writer/expected/styles_font.xml'), content
  end

end
