module Xl::Xml::Writer::StringTable

  def extract_string_table(workbook)
    strings_list = []

    # @todo Extract to Worksheet#strings
    workbook.worksheets.each do |sheet|
      sheet.get_cell_collection.each do |cell|
        if cell.string? && !cell.raw_value.nil?
          strings_list << cell.raw_value
        end
      end
    end

    {}.tap do |h|
      strings_list.uniq.each_with_index {|s, i| h[s] = i}
    end
  end

  def string_table_document(string_table)
    XML::Document.new.tap do |doc|
      strings = string_table.sort_by {|x| x.last}.map {|x| x.first}
      doc.root = make_node('sst', :uniqueCount => strings.length, :xmlns => Xl::Xml::NAMESPACES['ns'])
      strings.each do |str|
        si = make_subnode(doc.root, 'si')
        t = make_subnode(si, 't')
        t['xml:space'] = 'preserve' if str.strip != str
        t << str
      end
    end
  end

  def write_string_table(string_table)
    string_table_document(string_table).to_s
  end

end
