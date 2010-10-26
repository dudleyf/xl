module Xl::Xml::Reader::Workbook
  BUGGY_NAMED_RANGES = ['NA()', '#REF!']
  DISCARDED_RANGES = ['Xl_BuiltIn']

  def read_properties_core(xml)
    root = read_xml(xml)

    Xl::Workbook::DocumentProperties.new.tap do |props|
      props.creator = find_creator(root)
      props.last_modified_by = find_last_modified_by(root)
      props.created = find_created(root)
      props.modified = find_modified(root)
    end
  end

  def read_sheets_titles(xml)
    root = read_xml(xml, :default_namespace_prefix => 'ep')

    vector = root.find_first('//ep:TitlesOfParts/vt:vector')
    parts, names = get_number_of_parts(root)
    size = parts[names[0]]
    children = vector.children.map {|x| x.content}
    children[0...size]
  end

  def read_named_ranges(xml, workbook)
    doc = read_xml(xml, :default_namespace_prefix => 'ns')

    [].tap do |named_ranges|
      doc.find('//ns:definedName').each do |node|
        unless discard_named_range?(node)
          sheet_name, column, row = Xl::NamedRange.split(node.content)
          range_name = node['name']
          worksheet = workbook.get_sheet_by_name(sheet_name)
          range = '%s%s' % [column, row]
          named_ranges << Xl::NamedRange.new(range_name, worksheet, range)
        end
      end
    end
  end

  private

    def find_creator(root)
      creator_node = root.find_first('//dc:creator')
      creator_node ? creator_node.content : ''
    end

    def find_last_modified_by(root)
      last_modified_by_node = root.find_first('//cp:lastModifiedBy')
      last_modified_by_node ? last_modified_by_node.content : ''
    end

    def find_created(root)
      created_node = root.find_first('//dcterms:created')
      created_node ? DateTime.parse(created_node.content) : nil
    end

    def find_modified(root)
      modified_node = root.find_first('//dcterms:modified')
      modified_node ? DateTime.parse(modified_node.content) : nil
    end

    def get_number_of_parts(root)
      parts_size = {}
      parts_names = []

      vector = root.find_first('//ep:HeadingPairs/vt:vector')
      children = vector.children

      (0...children.length).step(2) do |child_id|
        part_name = children[child_id].find_first('vt:lpstr').content
        parts_names << part_name unless parts_names.include?(part_name)

        part_size = children[child_id+1].find_first('vt:i4').content
        parts_size[part_name] = part_size.to_i
      end

      [parts_size, parts_names]
    end

    def discard_named_range?(node)
      BUGGY_NAMED_RANGES.include?(node.content) ||
        DISCARDED_RANGES.include?(node['name'])
    end

end
