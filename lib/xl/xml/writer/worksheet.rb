module Xl::Xml::Writer::Worksheet

  def worksheet_document(worksheet, string_table, style_table)
    XML::Document.new.tap do |doc|
      doc.root = make_node('worksheet', {
        'xml:space' => 'preserve',
        'xmlns' => Xl::Xml::NAMESPACES['ns'],
        'xmlns:r' => Xl::Xml::NAMESPACES['dr']
      })

      add_sheet_pr(doc.root, worksheet)
      add_dimension(doc.root, worksheet)
      add_sheet_views(doc.root, worksheet)
      add_sheet_format_pr(doc.root, worksheet)
      add_worksheet_cols(doc.root, worksheet)
      add_worksheet_data(doc.root, worksheet, string_table, style_table)
      add_worksheet_merged_cells(doc.root, worksheet)
      add_worksheet_hyperlinks(doc.root, worksheet)
    end
  end

  def write_worksheet(worksheet, string_table, style_table)
    worksheet_document(worksheet, string_table, style_table).to_s
  end

  def add_sheet_pr(root, worksheet)
    make_subnode(root, 'sheetPr').tap do |sheet_pr|
      make_subnode(sheet_pr, 'outlinePr', {
        'summaryBelow' => worksheet.show_summary_below? ? "1" : "0",
        'summaryRight' => worksheet.show_summary_right? ? "1" : "0"
      })
    end
  end

  def add_dimension(root, worksheet)
    make_subnode(root, 'dimension', 'ref' => worksheet.calculate_dimension)
  end

  def add_sheet_views(root, worksheet)
    make_subnode(root, 'sheetViews').tap do |sheet_views|
      sheet_view = make_subnode(sheet_views, 'sheetView', 'workbookViewId' => '0')
      make_subnode(sheet_view, 'selection', {
        'activeCell' => worksheet.active_cell,
        'sqref' => worksheet.selected_cell
      })
    end
  end

  def add_sheet_format_pr(root, worksheet)
    make_subnode(root, 'sheetFormatPr', {'defaultRowHeight' => 15})
  end

  def add_worksheet_cols(root, worksheet)
    if worksheet.column_dimensions
      make_subnode(root, 'cols').tap do |cols|
        worksheet.column_dimensions.each do |column_string, dim|
          cidx = column_index_from_string(column_string)
          attrs = {
            'collapsed' => dim.style_index.to_s,
            'min' => cidx.to_s,
            'max' => cidx.to_s
          }
          attrs['customWidth'] = 'true' unless worksheet.default_column_dimension.width == dim.width
          attrs['hidden'] = 'true' unless dim.visible?
          attrs['outlineLevel'] = dim.outline_level if dim.outline_level > 0
          attrs['collapsed'] = 'true' if dim.collapsed?
          attrs['bestFit'] = 'true' if dim.auto_size?
          attrs['width'] = dim.width > 0 ? dim.width.to_s : '9.10'

          make_subnode(cols, 'col', attrs)
        end
      end
    end
  end

  def add_worksheet_data(root, worksheet, string_table, style_table)
    sheet_data = make_subnode(root, 'sheetData')
    max_column = worksheet.get_highest_column
    cells_by_row = {}
    worksheet.get_cell_collection.each do |cell|
      cells_by_row[cell.row] ||= []
      cells_by_row[cell.row] << cell
    end

    cells_by_row.sort_by {|k,v| k}.each do |row_idx, cell|
      row_dimension = worksheet.row_dimensions[row_idx]
      attrs = {'r' => row_idx, 'spans' => "1:#{max_column}"}

      if row_dimension.height > 0
        attrs['ht'] = row_dimension.height
        attrs['customHeight'] = 1
      end

      row = make_subnode(sheet_data, 'row', attrs)
      row_cells = cells_by_row[row_idx]
      sorted_cells = row_cells.sort_by {|x| column_index_from_string(x.column)}
      sorted_cells.each do |cell|
        value = cell.raw_value
        coordinate = cell.get_coordinate
        attrs = {'r' => coordinate, 't' => cell.data_type}
        if worksheet.styles.include?(coordinate)
          attrs['s'] = style_table[worksheet.styles[coordinate]]
        end

        c = make_subnode(row, 'c', attrs)

        unless value.nil?
          if cell.string?
            make_subnode(c, 'v') << string_table[value]
          elsif cell.formula?
            make_subnode(c, 'f') << value[1..-1]
            make_subnode(c, 'v')
          else
            make_subnode(c, 'v') << value
          end
        end
      end
    end

  end

  def add_worksheet_hyperlinks(root, worksheet)
    cells = worksheet.get_cell_collection
    write_hyperlinks = cells.any? {|x| !x.hyperlink_rel_id.nil?}
    if write_hyperlinks
      index = 1
      hyperlinks = make_subnode(root, 'hyperlinks')
      cells.each do |cell|
        unless cell.hyperlink_rel_id.nil?
          make_subnode(hyperlinks, 'hyperlink', {
            'display' => cell.hyperlink,
            'ref' => cell.get_coordinate,
            'r:id' => cell.hyperlink_rel_id
          })
        end
      end
    end
  end

  def add_worksheet_merged_cells(root, worksheet)
    unless worksheet.merged_cells.empty?
      count = worksheet.merged_cells.length
      merge_cells = make_subnode(root, 'mergeCells', :count => count)
      worksheet.merged_cells.each do |ref|
        make_subnode(merge_cells, 'mergeCell', :ref => ref)
      end
    end
  end

  def worksheet_rels_document(worksheet)
    XML::Document.new.tap do |doc|
      doc.root = make_node('Relationships', 'xmlns' => Xl::Xml::NAMESPACES['pr'])
      if worksheet.relationships
        worksheet.relationships.each do |rel, i|
          attrs = {
            'Id' => rel.id,
            'Type' => rel.type,
            'Target' => rel.target
          }
          attrs['TargetMode'] = rel.target_mode if rel.target_mode
          make_subnode(doc.root, 'Relationship', attrs)
        end
      end
    end
  end

  def write_worksheet_rels(worksheet)
    worksheet_rels_document(worksheet).to_s
  end
end
