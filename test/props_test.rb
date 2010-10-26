require File.join(File.dirname(__FILE__), "test_helper")

class PropsTest < XlTestCase
  
  def test_read_properties_core
    Xl::Zip.read(test_data_file('genuine/empty.xlsx')) do |archive|
      content = archive.core
      props = Xl::Xml.read_properties_core(content)
      
      assert_equal '*.*', props.creator
      assert_equal '*.*', props.last_modified_by
      
      assert_equal DateTime.new(2010, 4, 9, 20, 43, 12), props.created
      assert_equal DateTime.new(2010, 4, 11, 16, 20, 29), props.modified
    end
  end

  def test_read_sheets_titles
    Xl::Zip.read(test_data_file('genuine/empty.xlsx')) do |archive|
      content = archive.app
      sheet_titles = Xl::Xml.read_sheets_titles(content)
      assert_equal ['Sheet1 - Text', 'Sheet2 - Numbers', 'Sheet3 - Formulas'], sheet_titles
    end
  end

  def test_read_sheets_titles_mixed
    content = test_data('reader/app-multi-titles.xml')
    sheet_titles = Xl::Xml.read_sheets_titles(content)
    assert_equal ['ToC', 'ContractYear', 'ContractTier', 'Demand', 'LinearizedFunction', 'Market', 'Transmission'], sheet_titles
  end

  def test_number_of_parts
    content = Xl::Xml.read_xml(test_data('reader/app-multi-titles.xml'), :default_namespace_prefix => 'ep')
    parts = Xl::Xml.send(:get_number_of_parts, content)
    assert_equal [{'Worksheets' => 7, 'Named Ranges' => 7}, ['Worksheets', 'Named Ranges']], parts
  end

  def test_write_properties_core
    props = Xl::Workbook::DocumentProperties.new
    props.creator = 'TEST_USER'
    props.last_modified_by = 'SOMEBODY'
    props.created = DateTime.new(2010, 4, 1, 20, 30, 0)
    props.modified = DateTime.new(2010, 4, 5, 14, 5, 30)

    content = Xl::Xml.write_properties_core(props)

    assert_xml_equal(test_data('writer/expected/core.xml'), content)
  end

  def test_write_properties_app
    wb = Xl::Workbook.new
    wb.create_sheet
    wb.create_sheet
    content = Xl::Xml.write_properties_app(wb)
    assert_xml_equal(test_data('writer/expected/app.xml'), content)
  end
end