require File.join(File.dirname(__FILE__), "test_helper")

class ThemeTest < XlTestCase
  def test_write_theme
    content = Xl::Xml.write_theme
    assert_xml_equal(test_data('writer/expected/theme1.xml'), content)
  end
end
