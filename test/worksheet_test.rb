require File.join(File.dirname(__FILE__), "test_helper")

class WorksheetTest < XlTestCase

  def setup
    @wb = Xl::Workbook.new
  end

  def test_new_worksheet
    ws = Xl::Worksheet.new(@wb)
    assert_equal(@wb, ws.parent)
  end

  def test_get_cell
    ws = Xl::Worksheet.new(@wb)
    c = ws.cell('A1')
    assert_equal('A1', c.get_coordinate)
  end

  def test_set_wrong_title
    assert_raises(Xl::SheetTitleError) do
      Xl::Worksheet.new(@wb, 'X' * 50)
    end
  end

  def test_worksheet_dimension
    ws = Xl::Worksheet.new(@wb)
    assert_equal 'A1:A1', ws.calculate_dimension
    ws.cell('B12').value = 'AAA'
    assert_equal 'A1:B12', ws.calculate_dimension
  end

  def test_worksheet_range
    ws = Xl::Worksheet.new(@wb)
    range = ws.range('A1:C4')

    assert_equal(4, range.length)
    assert_equal(3, range.first.length)
  end

  def test_worksheet_range_named_range
    ws = Xl::Worksheet.new(@wb)
    @wb.create_named_range('test_range', ws, 'C5')
    range = ws.range('test_range')

    assert_kind_of(Xl::Cell, range)
    assert_equal(5, range.row)
  end

  def test_cell_offset
    ws = Xl::Worksheet.new(@wb)
    assert_equal 'C17', ws.cell('B15').offset(2,1).get_coordinate
  end

  def test_range_offset
    ws = Xl::Worksheet.new(@wb)
    range = ws.range('A1:C4', 1, 3)

    assert_equal(4, range.length)
    assert_equal(3, range.first.length)
    assert_equal('D2', range.first.first.get_coordinate)
  end

  def test_cell_alternate_coordinates
    ws = Xl::Worksheet.new(@wb)
    c = ws.cell(8,4)
    assert_equal "D8", c.get_coordinate
  end

  def test_cell_range_name
    ws = Xl::Worksheet.new(@wb)
    @wb.create_named_range('test_range_single', ws, 'B12')
    assert_raises(Xl::CellCoordinatesError) do
      ws.cell('test_range_single')
    end
    assert_equal(ws.range('B12'), ws.range('test_range_single'))
    assert_equal(ws.range('B12'), ws.cell('B12'))
  end

  def test_garbage_collect
    ws = Xl::Worksheet.new(@wb)
    ws.cell('A1').value = ''
    ws.cell('B2').value = '0'
    ws.cell('C4').value = 0

    ws.garbage_collect

    assert_equal(ws.get_cell_collection, [ws.cell('B2'), ws.cell('C4')])
  end

  def test_hyperlink_relationships
    ws = Xl::Worksheet.new(@wb)

    assert_equal(0, ws.relationships.length)

    ws.cell('A1').hyperlink = "http://test.com"
    assert_equal(1, ws.relationships.length)
    assert_equal("rId1", ws.cell('A1').hyperlink_rel_id)
    assert_equal("rId1", ws.relationships.first.id)
    assert_equal("http://test.com", ws.relationships.first.target)
    assert_equal("External", ws.relationships.first.target_mode)

    ws.cell('A2').hyperlink = "http://test2.com"
    assert_equal(2, ws.relationships.length)
    assert_equal("rId2", ws.cell('A2').hyperlink_rel_id)
    assert_equal("rId2", ws.relationships[1].id)
    assert_equal("http://test2.com", ws.relationships[1].target)
    assert_equal("External", ws.relationships[1].target_mode)
  end

  def test_sheet_protection_password
    p = Xl::Worksheet::SheetProtection.new
    assert_equal('CBEB', p.hash_password('test'))

    p.password = 'test'
    assert_equal('CBEB', p.password)
  end

end
