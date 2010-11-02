module Xl::Xml::Reader::Worksheet

  class WorksheetParserCallbacks
    include XML::SaxParser::Callbacks

    def initialize(ws, string_table, style_table)
      @ws = ws
      @string_table = string_table
      @style_table = style_table
      @read_value = false
    end

    def on_characters(value)
      if @read_value && !value.nil?
        if @data_type == Xl::Cell::TYPE_STRING
          value = @string_table[value.to_i]
        end
        @ws.cell(@coordinate).value = value
        if !@style_id.nil?
          @ws.cell(@coordinate).style = @style_table[@style_id.to_i]
        end
      end
    end

    def on_start_element_ns(name, attrs, prefix, uri, namespaces)
      case name
        when 'c'
          @coordinate = attrs['r']
          @data_type = attrs.fetch('t', 'n')
          @style_id = attrs['s']
          @read_value = true
        when 'mergeCell'
          @ws.merged_cells << attrs['ref']
      end
    end

    def on_end_element_ns(name, prefix, uri)
      @read_value = false if name == 'c'
    end

    private

      # @example
      #   attrs_to_hash(['a', 1, 'b', 2]) #=> {'a' => 1, 'b' => 2}
      def attrs_to_hash(a)
        {}.tap do |h|
          (0...a.length).step(2) {|i| h[a[i]] = a[i+1]}
        end
      end
  end

  def read_worksheet(xml, parent, preset_title, string_table, style_table)
    Xl::Worksheet.new(parent, preset_title).tap do |ws|
      callbacks = WorksheetParserCallbacks.new(ws, string_table, style_table)
      read_xml(xml, :sax => callbacks)
    end
  end

end
