require File.join(File.dirname(__FILE__), "test_helper")

class MetaTest < XlTestCase
  def test_write_content_types
    wb = Xl::Workbook.new
    wb.create_sheet
    wb.create_sheet
    wb.create_sheet
    content = Xl::Xml.write_content_types(wb)
    assert_xml_equal(test_data('writer/expected/[Content_Types].xml'), content)
  end

  def test_write_root_rels
    wb = Xl::Workbook.new
    wb.create_sheet
    content = Xl::Xml.write_root_rels(wb)
    assert_xml_equal(test_data('writer/expected/.rels'), content)
  end
end
