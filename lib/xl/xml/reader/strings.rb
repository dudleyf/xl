module Xl::Xml::Reader::Strings

  def read_string_table(xml)
    {}.tap do |table|
      doc = read_xml(xml, :default_namespace_prefix => 'ns')
      doc.find('//ns:si').each_with_index do |node, i|
        s = get_string(node)
        table[i] = s
      end
    end
  end

  private

    def get_string(node)
      rich_nodes = node.find('ns:r')
      return get_text(node) if rich_nodes.empty?

      s = ''
      rich_nodes.each {|n| s += get_text(n)}
      s
    end

    def get_text(node)
      t = node.find_first('ns:t')
      cur = t.content
      cur.strip! unless preserve_whitespace?(t)
      cur
    end

    def preserve_whitespace?(node)
      space = node.attributes.get_attribute_ns(Xl::Xml::NAMESPACES['xml'], 'space')
      space && space.value == 'preserve'
    end

end
