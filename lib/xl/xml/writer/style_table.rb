module Xl::Xml::Writer::StyleTable

  def style_table_document(style_table)
    XML::Document.new.tap do |doc|
      styles = style_table.sort_by {|x| x.last}.map {|x| x.first}
      number_format_table = extract_number_formats(styles)
      font_table = extract_style_component(:font, styles)
      border_table = extract_style_component(:borders, styles)

      doc.root = make_node('styleSheet', 'xmlns' => Xl::Xml::NAMESPACES['ns'])
      add_number_formats(doc.root, number_format_table)
      add_fonts(doc.root, font_table)
      add_fills(doc.root, styles)
      add_borders(doc.root, border_table)
      add_cell_style_xfs(doc.root, styles)
      add_cell_xfs(doc.root, styles, number_format_table, font_table, border_table)
      add_cell_styles(doc.root, styles)
      add_dxfs(doc.root, styles)
      add_table_styles(doc.root, styles)
    end
  end

  def write_style_table(style_table)
    style_table_document(style_table).to_s
  end

  def add_number_formats(root, format_table)
    user_formats = format_table.reject {|fmt, id| fmt.builtin?}

    make_subnode(root, 'numFmts', 'count' => user_formats.length).tap do |numFmts|
      user_formats.each do |fmt, id|
        make_subnode(numFmts, 'numFmt', 'numFmtId' => id, 'formatCode' => fmt.format_code)
      end
    end
  end

  def add_fonts(root, font_table)
    fonts = font_table.sort_by {|x| x.last}.map {|x| x.first}
    make_subnode(root, 'fonts', 'count' => fonts.length).tap do |fonts_node|
      fonts.each do |font|
        font_node = make_subnode(fonts_node, 'font')
        make_subnode(font_node, 'b') if font.bold
        make_subnode(font_node, 'strike') if font.strikethrough
        make_subnode(font_node, 'outline') if font.outline
        make_subnode(font_node, 'shadow') if font.shadow
        if font.underline
          if font.underline == Xl::Font::UNDERLINE_SINGLE
            make_subnode(font_node, 'u')
          else
            make_subnode(font_node, 'u', 'val' => font.underline)
          end
        end
        if font.superscript
          make_subnode(font_node, 'vertAlign', 'val' => "superscript")
        elsif font.subscript
          make_subnode(font_node, 'vertAlign', 'val' => "subscript")
        end
        make_subnode(font_node, 'sz', 'val' => font.size) if font.size
        make_subnode(font_node, 'color', 'rgb' => font.color.rgb) if font.color
        make_subnode(font_node, 'name', 'val' => font.name) if font.name
      end
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

  def add_borders(root, border_table)
    borderses = border_table.sort_by {|x| x.last}.map {|x| x.first}
    make_subnode(root, 'borders', 'count' => borderses.length).tap do |borders_node|
      borderses.each do |borders|
        add_border_node(borders_node, borders)
      end
    end
  end

  def add_border_node(root, borders=nil)
    border_node = make_subnode(root, 'border')

    %w[left right top bottom diagonal].each do |d|
      node = make_subnode(border_node, d)
      if borders
        border = borders.send(d)
        if border
          if border.border_style
            node['style'] = border.border_style
            if border.color
              make_subnode(node, 'color', :rgb => border.color.rgb)
            end
          end
        end
      end
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

   def add_cell_xfs(root, styles, number_format_table, font_table, border_table)
     make_subnode(root, 'cellXfs', 'count' => styles.length).tap do |cell_xfs|
       styles.each do |style|
         xf_node = make_subnode(cell_xfs, 'xf', {
           'numFmtId' => 0,
           'fontId' => 0,
           'fillId' => 0,
           'xfId' => 0,
           'borderId' => 0
          })

         num_fmt_id = number_format_table[style.number_format]
         unless num_fmt_id.nil? || num_fmt_id == 0
           xf_node['numFmtId'] = num_fmt_id.to_s
           xf_node['applyNumberFormat'] = '1'
         end

         font_id = font_table[style.font]
         unless font_id.nil? || font_id == 0
           xf_node['fontId'] = font_id.to_s
           xf_node['applyFont'] = '1'
         end

         border_id = border_table[style.borders]
         unless border_id.nil? || border_id == 0
           xf_node['borderId'] = border_id.to_s
           xf_node['applyBorder'] = '1'
         end

         default = Xl::Alignment.new
         if style.alignment != default
           align = style.alignment
           xf_node['applyAlignment'] = '1'

           make_subnode(xf_node, 'alignment').tap do |a|
             a['horizontal'] = align.horizontal.to_s if align.horizontal != default.horizontal
             a['vertical'] = align.vertical.to_s if align.vertical != default.vertical
             a['textRotation'] = align.text_rotation.to_s if align.text_rotation != 0
             a['wrapText'] = '1' if align.wrap_text
             a['shrinkToFit'] = '1' if align.shrink_to_fit
             a['indent'] = align.indent.to_s if align.indent != 0
           end
         end
       end
     end
   end

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

   def extract_style_table(workbook)
     {}.tap do |h|
       workbook.styles.each_with_index {|s, i| h[s] = i}
     end
   end

   def extract_number_formats(styles)
     {}.tap do |table|
       formats = []
       num_fmt_id = 165 # start higher than any builtin

       formats = styles.map {|x| x.number_format }.uniq
       formats.each do |format|
         if format.builtin?
           table[format] = format.builtin_format_id(format.format_code)
         else
           table[format] = num_fmt_id
           num_fmt_id += 1
         end
       end
     end
   end

   def extract_style_component(component, styles)
     {}.tap do |table|
       components = styles.map {|x| x.send(component.to_s)}.uniq
       components.each_with_index do |x,i|
         table[x] = i
       end
     end
   end
end
