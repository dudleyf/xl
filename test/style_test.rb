require File.join(File.dirname(__FILE__), "test_helper")

class StyleTest < XlTestCase
  include Xl::Xml

  def test_extract_style_table
    wb = Xl::Workbook.new
    ws = wb.create_sheet
    ws.cell('A1').value = '12.34%'
    now = Time.now
    ws.cell('B4').value = now
    ws.cell('B5').value = now
    ws.cell('C14').value = "This is a test"
    ws.cell('D9').value = '31.31415'
    ws.cell('D9').style.number_format.format_code = Xl::NumberFormat::FORMAT_NUMBER_00

    table = Xl::Xml.extract_style_table(wb)
    assert_equal(4, table.size)
  end

  def test_write_style_table
    wb = Xl::Workbook.new
    ws = wb.create_sheet
    ws.cell('A1').value = '12.34%'
    now = Time.now
    ws.cell('B4').value = now
    ws.cell('B5').value = now
    ws.cell('C14').value = "This is a test"
    ws.cell('D9').value = '31.31415'
    ws.cell('D9').style.number_format.format_code = Xl::NumberFormat::FORMAT_NUMBER_00

    table = Xl::Xml.extract_style_table(wb)
    content = Xl::Xml.write_style_table(table)
    assert_xml_equal(test_data('writer/expected/styles_simple.xml'), content)
  end

  def test_read_style_table
    style_table = Xl::Xml.read_style_table(test_data('reader/simple-styles.xml'))
    assert_equal(4, style_table.size)
    assert_equal(Xl::NumberFormat::BUILTIN_FORMATS[9], style_table[1].number_format.format_code)
    assert_equal('yyyy-mm-dd', style_table[2].number_format.format_code)
  end

  def test_read_cell_style
    style_table = Xl::Xml.read_style_table(test_data('reader/empty-workbook-styles.xml'))
    assert_equal(2, style_table.length)
  end

  def test_styles_equal
    s1 = Xl::Style.new
    s2 = Xl::Style.new

    assert s1 == s2
    assert s1.eql?(s2)
    assert s1.hash == s2.hash
    assert [s1,s2].uniq == [s1]

    s2.font.size = 20

    assert s1 != s2
    assert !s1.eql?(s2)
    assert s1.hash != s2.hash
    assert [s1,s2].uniq == [s1,s2]
  end

  def test_write_style_table_fonts
    wb = Xl::Workbook.new
    ws = wb.create_sheet

    ws.cell('E3').style.font = Xl::Font.new({
      :name => "Times New Roman",
      :size => 33,
      :bold => true,
      :italic => true,
      :superscript => true,
      :outline => true,
      :shadow => true,
      :underline => Xl::Font::UNDERLINE_SINGLE,
      :strikethrough => true,
      :color => Xl::Color.new(Xl::Color::RED)
    })

    # ws.cell('E3').style.font = Xl::Font.new({
    #   :name => "Times New Roman",
    #   :size => 33,
    #   :bold => true,
    #   :italic => true,
    #   :superscript => true,
    #   :outline => true,
    #   :shadow => true,
    #   :underline => :single,
    #   :strikethrough => true,
    #   :color => :red
    # })

    table = Xl::Xml.extract_style_table(wb)
    content = Xl::Xml.write_style_table(table)
    assert_xml_equal test_data('writer/expected/styles_font.xml'), content
  end

  def test_write_style_table_borders
    wb = Xl::Workbook.new
    ws = wb.create_sheet

    ws.cell('A1').style.borders = Xl::Borders.new({
      :top => Xl::Border.new(:border_style => Xl::Border::BORDER_THIN),
      :bottom => Xl::Border.new(:border_style => Xl::Border::BORDER_DOUBLE, :color => Xl::Color.new(Xl::Color::RED)),
      :left => Xl::Border.new(:border_style => Xl::Border::BORDER_HAIR),
      :right => Xl::Border.new(:border_style => Xl::Border::BORDER_THICK)
    })

    # ws.cell('A1').style.borders = {
    #   :top => :thin,
    #   :bottom => {:style => :double, :color => :red},
    #   :left => :dashed
    #   :right => :thick
    # })

    table = Xl::Xml.extract_style_table(wb)
    content = Xl::Xml.write_style_table(table)
    assert_xml_equal test_data('writer/expected/styles_border.xml'), content
  end

  def test_border_medium_or_thick
    Xl::Border::MEDIUM_OR_THICK_BORDERS.each do |s|
       assert Xl::Border.new(:border_style => s).medium_or_thick?, s
     end

    [Xl::Border::BORDER_NONE,
     Xl::Border::BORDER_DASHDOT,
     Xl::Border::BORDER_DASHDOTDOT,
     Xl::Border::BORDER_DASHED,
     Xl::Border::BORDER_DOTTED,
     Xl::Border::BORDER_HAIR,
     Xl::Border::BORDER_THIN].each do |s|
       assert !Xl::Border.new(:border_style => s).medium_or_thick?, s
     end
  end
end
