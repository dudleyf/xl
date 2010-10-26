require File.join(File.dirname(__FILE__), "test_helper")

class ReadTest < XlTestCase

  def setup
    @tmpdir = Dir.mktmpdir
  end

  def teardown
    FileUtils.rm_rf(@tmpdir) if @tmpdir && File.directory?(@tmpdir)
  end

  def test_read_worksheet
    xml_source = test_data('reader/sheet2.xml')
    ws = Xl::Xml.read_worksheet(xml_source, Xl::Workbook.new, 'Sheet 2', {1 => 'hello'}, {1 => Xl::Style.new})

    assert_kind_of Xl::Worksheet, ws
    assert_equal 'hello', ws.cell('G5').value
    assert_equal 30, ws.cell('D30').value
    assert_equal 0.09, ws.cell('K9').value
  end

  def test_read_workbook
    genuine_wb = test_data_file('genuine/empty.xlsx')
    test_wb = File.join(@tmpdir, 'test.xlsx')
    FileUtils.cp(genuine_wb, test_wb)
    
    wb = Xl::Xml.load_workbook(test_wb)
    assert_kind_of(Xl::Workbook, wb)

    sheet2 = wb.get_sheet_by_name('Sheet2 - Numbers')
    assert_kind_of(Xl::Worksheet, sheet2)
    assert_equal('This is cell G5', sheet2.cell('G5').value)
    assert_equal(18, sheet2.cell('D18').value)    
  end

  def test_read_workbook_no_string_table
    wb = Xl::Xml.load_workbook(test_data_file('genuine/empty-no-string.xlsx'))
    assert_kind_of(Xl::Workbook, wb)
  end

  def test_read_workbook_with_styles
    genuine_wb = test_data_file('genuine/empty-with-styles.xlsx')
    wb = Xl::Xml.load_workbook(genuine_wb)
    ws = wb.get_sheet_by_name('Sheet1')

    assert_equal Xl::NumberFormat::FORMAT_GENERAL, ws.cell('A1').style.number_format.format_code
    assert_equal Xl::NumberFormat::FORMAT_DATE_XLSX14, ws.cell('A2').style.number_format.format_code
    assert_equal Xl::NumberFormat::FORMAT_NUMBER_00, ws.cell('A3').style.number_format.format_code
    assert_equal Xl::NumberFormat::FORMAT_DATE_TIME3, ws.cell('A4').style.number_format.format_code
    assert_equal Xl::NumberFormat::FORMAT_PERCENTAGE_00, ws.cell('A5').style.number_format.format_code
  end
end
