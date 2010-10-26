require File.join(File.dirname(__FILE__), "test_helper")

class WorkbookTest < XlTestCase

  def test_new_workbook
    assert_nothing_raised do
      wb = Xl::Workbook.new
    end
  end

  def test_get_active_sheet
    wb = Xl::Workbook.new
    wb.create_sheet

    s = wb.get_active_sheet
    assert_not_nil(s)
    assert_equal(wb.worksheets.first, s)
  end

  def test_create_sheet_with_index
    wb = Xl::Workbook.new

    s = wb.create_sheet(0)
    assert_equal(wb.worksheets.first, s)
  end

  def test_create_sheet_without_index
    wb = Xl::Workbook.new

    s = wb.create_sheet
    assert_equal(wb.worksheets.last, s)
  end

  def test_remove_sheet
    wb = Xl::Workbook.new

    s = wb.create_sheet(0)
    assert(wb.worksheets.include?(s))
    wb.remove_sheet(s)
    assert(!wb.worksheets.include?(s))
  end

  def test_get_sheet_by_name
    wb = Xl::Workbook.new

    s = wb.create_sheet
    title = 'my sheet'
    s.title = title
    assert_equal(s, wb.get_sheet_by_name(title))
  end

  def test_get_index
    wb = Xl::Workbook.new

    s = wb.create_sheet(0)
    idx = wb.get_index(s)
    assert_equal(idx, 0)
  end

  def test_get_sheet_names
    wb = Xl::Workbook.new

    names = %w[Sheet Sheet1 Sheet2 Sheet3 Sheet4 Sheet5]
    6.times do
      wb.create_sheet(0)
    end

    actual = wb.get_sheet_names
    assert_equal(actual.sort, names.sort)
  end

  # TODO: Fix this test
  # def test_get_named_ranges
  #   wb = Xl::Workbook.new
  #
  #   assert(false)
  # end

  def test_get_named_range
    wb = Xl::Workbook.new
    s = wb.create_sheet
    r = Xl::NamedRange.new('test_nr', s, 'A1')
    wb.add_named_range(r)
    assert_equal(r, wb.get_named_range('test_nr'))
  end

  def test_remove_named_range
    wb = Xl::Workbook.new
    s = wb.create_sheet
    r = Xl::NamedRange.new('test_nr', s, 'A1')

    wb.add_named_range(r)
    assert(wb.get_named_ranges.include?(r))
    wb.remove_named_range(r)
    assert(!wb.get_named_ranges.include?(r))
  end
end
