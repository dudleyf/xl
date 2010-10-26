require File.join(File.dirname(__FILE__), "test_helper")

require 'time'

class CellTest < XlTestCase

  def test_value_types
    c = Xl::Cell.new(nil, 'A', 1)
    assert_equal Xl::Cell::TYPE_NULL, c.data_type

    c.value = 42
    assert_equal Xl::Cell::TYPE_NUMERIC, c.data_type

    c.value = 'hello'
    assert_equal Xl::Cell::TYPE_STRING, c.data_type

    c.value = '=42'
    assert_equal Xl::Cell::TYPE_FORMULA, c.data_type

    c.value = '4.2'
    assert_equal Xl::Cell::TYPE_NUMERIC, c.data_type

    c.value = '-42.00'
    assert_equal Xl::Cell::TYPE_NUMERIC, c.data_type

    c.value = '0'
    assert_equal Xl::Cell::TYPE_NUMERIC, c.data_type

    c.value = 0
    assert_equal Xl::Cell::TYPE_NUMERIC, c.data_type

    c.value = 0.0001
    assert_equal Xl::Cell::TYPE_NUMERIC, c.data_type

    c.value = '0.9999'
    assert_equal Xl::Cell::TYPE_NUMERIC, c.data_type
  end

  def test_type_predicates
    c = Xl::Cell.new(nil, 'A', 1)
    assert c.null?

    c.value = 42
    assert c.numeric?

    c.value = 'hello'
    assert c.string?

    c.value = '=42'
    assert c.formula?

    c.value = true
    assert c.bool?
  end

  # def test_time_value
  #   wb = Xl::Workbook.new
  #   ws = Xl::Worksheet.new(wb)
  #   c = Xl::Cell.new(ws, 'A', 1)
  # 
  #   c.value = DateTime.parse('03:40:16')
  #   assert_equal Xl::Cell::TYPE_NUMERIC, c.data_type
  #   assert_equal(DateTime.parse("03:40:16"), c.value)
  # end

  def test_date_format_applied_to_non_dates
    wb = Xl::Workbook.new
    ws = Xl::Worksheet.new(wb)
    c = Xl::Cell.new(ws, 'A', 1)
  
    c.value = Time.now
    c.value = 'testme'
    assert_equal 'testme', c.value
  end

end
