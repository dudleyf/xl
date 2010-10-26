require File.join(File.dirname(__FILE__), "test_helper")

class DateHelperTest < XlTestCase

  def setup
    @dates_1900 = {
      DateTime.parse("1997-08-11") => 35653,
      DateTime.parse("1997-08-11T01:01:01Z") => 35653.042372685188,
      DateTime.parse("2006-10-15T13:25:58Z") => 39005.559699074074
    }

    @dates_1904 = {
      DateTime.parse("2001-08-12") => 35653,
      DateTime.parse("2001-08-12T01:01:01Z") => 35653.042372685188,
      DateTime.parse("2010-10-16T13:25:58Z") => 39005.559699074074
    }
  end

  def test_ruby_to_excel_1900
    @dates_1900.each do |k,v|
      assert_equal v, Xl::DateHelper.ruby_to_excel(k)
    end
  end

  def test_ruby_to_excel_1904
    @dates_1904.each do |k,v|
      assert_equal v, Xl::DateHelper.ruby_to_excel(k, true)
    end
  end

  def test_excel_to_ruby_1900
    @dates_1900.each do |k,v|
      assert_equal k, Xl::DateHelper.excel_to_ruby(v)
    end    
  end

  def test_excel_to_ruby_1904
    @dates_1904.each do |k,v|
      assert_equal k, Xl::DateHelper.excel_to_ruby(v, true)
    end
  end

end
