require File.join(File.dirname(__FILE__), "test_helper")

class StringsTest < XlTestCase

  def test_extract_string_table
    wb = Xl::Workbook.new
    ws = wb.create_sheet

    ws.cell('B12').value = 'hello'
    ws.cell('B13').value = 'world'
    ws.cell('D28').value = 'hello'

    table = Xl::Xml.extract_string_table(wb)

    assert_equal({
      'hello' => 0,
      'world' => 1
    }, table)
  end

  def test_read_string_table
    table = Xl::Xml.read_string_table(test_data('reader/sharedStrings.xml'))

    assert_equal({
      0 => 'This is cell A1 in Sheet 1', 
      1 => 'This is cell G5'
    }, table)
  end

  def test_read_formatted_string_table
    table = Xl::Xml.read_string_table(test_data('reader/shared-strings-rich.xml'))

    assert_equal({
      0 => 'Welcome',
      1 => 'to the best shop in town' ,
      2 => "     let's play "
    }, table)
  end
  
end
