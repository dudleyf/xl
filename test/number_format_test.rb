require File.join(File.dirname(__FILE__), "test_helper")

class NumberFormatTest < XlTestCase

  def setup
    @workbook = Xl::Workbook.new
    @worksheet = Xl::Worksheet.new(@workbook, 'Test')
  end

  def test_insert_float
    @worksheet.cell('A1').value = 3.14
    assert_equal Xl::Cell::TYPE_NUMERIC, @worksheet.cell('A1').data_type
  end
  
  def test_insert_percentage
    @worksheet.cell('A1').value = "3.14%"
    assert_equal Xl::Cell::TYPE_NUMERIC, @worksheet.cell('A1').data_type
    assert_in_delta 0.0314, @worksheet.cell('A1').value, 0.00001
  end

  def test_insert_date
    @worksheet.cell('A1').value = Time.now
    assert_equal Xl::Cell::TYPE_NUMERIC, @worksheet.cell('A1').data_type
  end

  def test_internal_date
    dt = DateTime.new(2010, 7, 13, 6, 37, 41)
    @worksheet.cell('A3').value = dt
    assert_equal 40372.27616898148, @worksheet.cell('A3').raw_value
  end
  
  def test_date_interpretation
    dt = DateTime.new(2010, 7, 13, 6, 37, 41)
    @worksheet.cell('A3').value = dt
    assert_equal dt, @worksheet.cell('A3').value
  end

  def test_number_format_style
    @worksheet.cell('A1').value = '12.6%'
    assert_equal Xl::NumberFormat::FORMAT_PERCENTAGE, @worksheet.cell('A1').style.number_format.format_code
  end
end
