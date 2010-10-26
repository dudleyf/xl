require File.join(File.dirname(__FILE__), "test_helper")

class StyleTest < XlTestCase
  def setup
    @workbook = Xl::Workbook.new
    @worksheet = @workbook.create_sheet
    @worksheet.cell('A1').value = '12.34%'
    now = Time.now
    @worksheet.cell('B4').value = now
    @worksheet.cell('B5').value = now
    @worksheet.cell('C14').value = "This is a test"
    @worksheet.cell('D9').value = '31.31415'
    @worksheet.cell('D9').style.number_format.format_code = Xl::NumberFormat::FORMAT_NUMBER_00
  end
  
  def test_create_style_table
    table = Xl::Xml.create_style_table(@workbook)
    assert_equal(3, table.size)
  end

  def test_write_style_table
    table = Xl::Xml.create_style_table(@workbook)
    content = Xl::Xml.write_style_table(table)
    assert_xml_equal(test_data('writer/expected/simple-styles.xml'), content)
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
end
