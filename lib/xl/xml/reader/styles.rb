module Xl::Xml::Reader::Styles

  def read_style_table(xml)
    {}.tap do |table|
      doc = read_xml(xml, :default_namespace_prefix => 'ns')

      custom_num_formats = parse_custom_num_formats(doc)
      builtin_formats = Xl::NumberFormat::BUILTIN_FORMATS

      doc.find('//ns:cellXfs/ns:xf').each_with_index do |node, i|
        new_style = Xl::Style.new
        number_format_id = node['numFmtId'].to_i
        if number_format_id < 164
          new_style.number_format.format_code = builtin_formats[number_format_id]
        else
          new_style.number_format.format_code = custom_num_formats[number_format_id]
        end
        table[i] = new_style
      end
    end
  end

  private

    def parse_custom_num_formats(doc)
      {}.tap do |custom_formats|
        doc.find('//ns:numFmts/ns:numFmt').each_with_index do |node, i|
          custom_formats[node['numFmtId'].to_i] = node['formatCode']
        end
      end
    end

end
