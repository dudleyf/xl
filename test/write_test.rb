require File.join(File.dirname(__FILE__), "test_helper")

class WriteTest < XlTestCase

  def setup
    @tmpdir = Dir.mktmpdir
  end

  def teardown
    FileUtils.rm_rf(@tmpdir) if @tmpdir && File.directory?(@tmpdir)
  end

  def test_write_workbook_rels
    wb = Xl::Workbook.new
    wb.create_sheet
    content = Xl::Xml.write_workbook_rels(wb)
    assert_xml_equal test_data('writer/expected/workbook.xml.rels'), content
  end

  def test_write_workbook
    wb = Xl::Workbook.new
    wb.create_sheet
    content = Xl::Xml.write_workbook(wb)
    assert_xml_equal test_data('writer/expected/workbook.xml'), content
  end

  def test_write_string_table
    table = {
      'hello' => 1,
      'world' => 2,
      'nice' =>  3
    }
    content = Xl::Xml.write_string_table(table)
    assert_xml_equal test_data('writer/expected/sharedStrings.xml'), content
  end

  def test_write_worksheet
    wb = Xl::Workbook.new
    ws = wb.create_sheet
    ws.cell('F42').value = 'hello'
    content = Xl::Xml.write_worksheet(ws, {'hello' => 0}, {})
    assert_xml_equal test_data('writer/expected/sheet1.xml'), content
  end

  def test_write_worksheet_with_formula
    wb = Xl::Workbook.new
    ws = wb.create_sheet
    ws.cell('F1').value = 10
    ws.cell('F2').value = 32
    ws.cell('F3').value = '=F1+F2'
    content = Xl::Xml.write_worksheet(ws, {}, {})
    assert_xml_equal test_data('writer/expected/sheet1_formula.xml'), content
  end

  def test_write_worksheet_with_style
    wb = Xl::Workbook.new
    ws = wb.create_sheet
    ws.cell('F1').value = '13%'

    shared_style_table = Xl::Xml.create_style_table(wb)
    style_id_by_hash = {}.tap do |h|
      shared_style_table.each do |style, id|
        h[style.crc] = id
      end
    end

    content = Xl::Xml.write_worksheet(ws, {}, style_id_by_hash)
    assert_xml_equal test_data('writer/expected/sheet1_style.xml'), content
  end

  def test_write_worksheet_with_height
    wb = Xl::Workbook.new
    ws = wb.create_sheet
    ws.cell('F1').value = 10
    ws.row_dimensions[ws.cell('F1').row].height = 30
    content = Xl::Xml.write_worksheet(ws, {}, {})
    assert_xml_equal test_data('writer/expected/sheet1_height.xml'), content
  end

  def test_write_worksheet_with_hyperlink
    wb = Xl::Workbook.new
    ws = wb.create_sheet
    ws.cell('A1').value = 'test'
    ws.cell('A1').hyperlink = 'http://test.com'
    content = Xl::Xml.write_worksheet(ws, {'test' => 0}, {})
    assert_xml_equal test_data('writer/expected/sheet1_hyperlink.xml'), content
  end

  def test_write_worksheet_with_hyperlink_relationships
    wb = Xl::Workbook.new
    ws = wb.create_sheet

    assert_equal(0, ws.relationships.size)

    ws.cell('A1').value = 'test'
    ws.cell('A1').hyperlink = 'http://test.com/'
    assert_equal(1, ws.relationships.size)

    ws.cell('A2').value = 'test'
    ws.cell('A2').hyperlink = 'http://test2.com/'
    assert_equal(2, ws.relationships.size)

    content = Xl::Xml.write_worksheet_rels(ws)
    assert_xml_equal test_data('writer/expected/sheet1_hyperlink.xml.rels'), content
  end

  def test_hyperlink_value
    wb = Xl::Workbook.new
    ws = wb.create_sheet

    ws.cell('A1').hyperlink = 'http://test.com'
    assert_equal 'http://test.com', ws.cell('A1').value

    content = Xl::Xml.write_worksheet(ws, {'http://test.com' => 0}, {})
    ws.cell('A1').value = "test"
    assert_equal "test", ws.cell('A1').value
  end

  def test_save_empty_workbook
    wb = Xl::Workbook.new
    wb.create_sheet
    destfile = File.join(@tmpdir, 'empty_book.xlsx')
    Xl::Xml.save_workbook(wb, destfile)
    assert File.file?(destfile)
  end

end
