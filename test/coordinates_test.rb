require File.join(File.dirname(__FILE__), "test_helper")

class CoordinatesTest < XlTestCase
  include Xl::Coordinates

  def test_coordinatep
    assert(coordinate?('A1'))
    assert(coordinate?('ZZZ34344'))
    assert(!coordinate?(3434))
    assert(!coordinate?('234thoeutho90gc34hb'))
  end

  def test_coordinate_from_string
    col, row = coordinate_from_string("ZF46")
    assert_equal("ZF", col)
    assert_equal(46, row)
  end
  
  def test_invalid_coordinate
    assert_raises Xl::CellCoordinatesError do
      coordinate_from_string("AAA")
    end
  end
  
  def test_absolute_coordinate
    assert_equal('$ZF$51', absolute_coordinate('ZF51'))
  end
  
  def test_column_index_from_string
    assert_equal 10, column_index_from_string('J')
    assert_equal 270, column_index_from_string('JJ')
    assert_equal 7030, column_index_from_string('JJJ')
    assert_raises(Xl::ColumnStringIndexError) do
      column_index_from_string('JJJJ')
    end
    assert_raises(Xl::ColumnStringIndexError) do
      column_index_from_string('')
    end
  end
  
  def test_column_letter
    assert_equal 'ZZZ', get_column_letter(18278)
    assert_equal 'AA', get_column_letter(27)
    assert_equal 'Z', get_column_letter(26)
  end
end
