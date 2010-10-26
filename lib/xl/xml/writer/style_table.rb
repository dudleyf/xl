module Xl::Xml::Writer::StyleTable

  def create_style_table(workbook)
    styles_by_crc = {}
    workbook.worksheets.each do |sheet|
      sheet.styles.values.each do |style|
        styles_by_crc[style.crc] = style
      end
    end

    {}.tap do |h|
      styles_by_crc.values.each_with_index do |style, i|
        h[style] = i+1
      end
    end
  end

  def style_table_document(style_table)
    XML::Document.new.tap do |doc|
      styles = style_table.sort_by {|x| x.last}.map {|x| x.first}
      number_format_table = extract_number_formats(styles)

      doc.root = make_node('styleSheet', 'xmlns' => Xl::Xml::NAMESPACES['ns'])

      add_number_formats(doc.root, number_format_table)
      add_fonts(doc.root, styles)
      add_fills(doc.root, styles)
      add_borders(doc.root, styles)
      add_cell_style_xfs(doc.root, styles)
      add_cell_xfs(doc.root, styles, number_format_table)
      add_cell_styles(doc.root, styles)
      add_dxfs(doc.root, styles)
      add_table_styles(doc.root, styles)
    end
  end

  def write_style_table(style_table)
    style_table_document(style_table).to_s
  end

  private

  def add_number_formats(root, format_table)
    user_formats = format_table.reject {|fmt, id| fmt.builtin?}

    make_subnode(root, 'numFmts', 'count' => user_formats.length).tap do |numFmts|
      user_formats.each do |fmt, id|
        make_subnode(numFmts, 'numFmt', 'numFmtId' => id, 'formatCode' => fmt.format_code)
      end
    end
  end

  # @todo actually write real fonts
  def add_fonts(root, styles)
    make_subnode(root, 'fonts', 'count' => 1).tap do |fonts|
      font = make_subnode(fonts, 'font')
      make_subnode(font, 'sz', 'val' => 11)
      make_subnode(font, 'color', 'theme' => 1)
      make_subnode(font, 'name', 'val' => 'Calibri')
      make_subnode(font, 'family', 'val' => 2)
      make_subnode(font, 'scheme', 'val' => 'minor')
    end
  end

  # @todo actually write real fills
  def add_fills(root, styles)
    make_subnode(root, 'fills', 'count' => 2).tap do |fills|

      fill = make_subnode(fills, 'fill')
      make_subnode(fill, 'patternFill', 'patternType' => 'none')

      fill = make_subnode(fills, 'fill')
      make_subnode(fill, 'patternFill', 'patternType' => 'gray125')
    end
  end

  # @todo actually write real borders
  def add_borders(root, styles)
    make_subnode(root, 'borders', 'count' => 1).tap do |borders|
      border = make_subnode(borders, 'border')
      make_subnode(border, 'left')
      make_subnode(border, 'right')
      make_subnode(border, 'top')
      make_subnode(border, 'bottom')
      make_subnode(border, 'diagonal')
    end
  end

  # @todo actually write real cell style xfs
   def add_cell_style_xfs(root, styles)
     make_subnode(root, 'cellStyleXfs', 'count' => 1).tap do |cell_style_xfs|
        make_subnode(cell_style_xfs, 'xf', {
          'numFmtId' => 0,
          'fontId' => 0,
          'fillId' => 0,
          'borderId' => 0
        })
      end
   end

   # @todo actually write real cell styles
   def add_cell_styles(root, styles)
     make_subnode(root, 'cellStyles', 'count' => 1).tap do |cell_styles|
       make_subnode(cell_styles, 'cellStyle', 'name' => 'Normal', 'xfId' => 0, 'builtinId' => 0)
     end
   end

   def add_cell_xfs(root, styles, number_format_table)
     make_subnode(root, 'cellXfs', 'count' => styles.length+1).tap do |cell_xfs|
       make_subnode(cell_xfs, 'xf', 'numFmtId' => 0, 'fontId' => 0, 'fillId' => 0, 'xfId' => 0, 'borderId' => 0)
       styles.each do |style|
         make_subnode(cell_xfs, 'xf', {
           'numFmtId' => number_format_table[style.number_format],
           'applyNumberFormat' => 1,
           'fontId' => 0,
           'fillId' => 0,
           'xfId' => 0,
           'borderId' => 0
         })
       end
     end
   end

   # @todo actually write real dxfs
   def add_dxfs(root, styles)
     make_subnode(root, 'dxfs', 'count' => 0)
   end

   # @todo actually write real table styles
   def add_table_styles(root, styles)
     make_subnode(root, 'tableStyles', {
       'count' => 0,
       'defaultTableStyle' => 'TableStyleMedium9',
       'defaultPivotStyle' => 'PivotStyleLight16'
     })
   end

   def extract_number_formats(styles)
     format_table = {}
     formats = []
     num_fmt_id = 165 # start higher than any builtin
     num_fmt_offset = 0

     styles.each do |style|
       formats << style.number_format unless formats.include?(style.number_format)
     end

     formats.each do |format|
       if format.builtin?
         format_table[format] = format.builtin_format_id(format.format_code)
       else
         format_table[format] = num_fmt_id + num_fmt_offset
         num_fmt_offset += 1
       end
     end

     format_table
   end

end
