require File.join(File.dirname(__FILE__), "test_helper")

class NamedRangeTest < XlTestCase
  def setup
  end

  def test_split
    assert_equal(['My Sheet', 'D', 8], Xl::NamedRange.split("'My Sheet'!$D$8"))
  end
  
  def test_split_no_quotes
    assert_equal(['HYPOTHESES', 'B', 3], Xl::NamedRange.split("HYPOTHESES!$B$3:$L$3"))
  end

  def test_read_named_ranges
    content = test_data('reader/workbook.xml')
    wb = Xl::Workbook.new
    wb.create_sheet.title = 'My Sheeet'
    named_ranges = Xl::Xml.read_named_ranges(content, wb)
    assert_equal(['My Sheeet!D8'], named_ranges.map(&:to_s))
  end

end
