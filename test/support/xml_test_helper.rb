# Stolen from testunitxml
# Original Copyright © 2006 Henrik Mårtensson

module XmlTestHelper

  def assert_xml_equal(expected, actual, msg=nil)
    expected = parse_xml(expected)
    actual = parse_xml(actual)

    _wrap_assertion do
      full_message = build_message(msg, <<EOT, actual, expected)

<?> expected to be equal to
<?> but was not.
EOT
      assert_block(full_message) do
        documents_equal?(expected, actual)
      end
    end
  end

  def assert_xml_not_equal(expected, actual, msg=nil)
    expected = parse_xml(expected)
    actual = parse_xml(actual)

    _wrap_assertion do
      full_message = build_message(msg, <<EOT, actual, expected)

<?> expected not to be equal to
<?> but it was equal.
EOT
      assert_block(full_message) do
        !documents_equal?(expected, actual)
      end
    end
  end

  private

  def parse_xml(xml)
    options = XML::Parser::Options::NOENT

    case xml
      when IO;     XML::Document.io(xml, :options => options)
      when String; XML::Document.string(xml, :options => options)
      else;        xml
    end
  end

  def documents_equal?(expected, actual)
    iterate_nodes(expected, actual) do |expected, actual|
      return false unless nodes_equal?(expected, actual)
    end
    true
  end

  # Compares two {XML::Node}s representing part of a document.
  # Does not recursively compare the nodes' children.
  def nodes_equal?(expected_node, actual_node)
    !(expected_node.nil? || actual_node.nil?) &&
      actual_node.instance_of?(expected_node.class) &&
      actual_node.element? ?
        elements_equal?(expected_node, actual_node) :
        contents_equal?(expected_node, actual_node)
  end

  def elements_equal?(expected_node, actual_node)
    expected_node.name == actual_node.name &&
      namespaces_equal?(expected_node, actual_node) &&
      attributes_equal?(expected_node, actual_node)
  end

  def namespaces_equal?(expected_node, actual_node)
    expected_ns = expected_node.namespaces.namespace
    actual_ns = actual_node.namespaces.namespace

    (expected_ns == actual_ns) ||
      (expected_ns.respond_to?(:href) && actual_ns.respond_to?(:href) && expected_ns.href == actual_ns.href)
  end

  def contents_equal?(expected_node, actual_node)
    expected_node.content == actual_node.content
  end

  def attributes_equal?(expected_node, actual_node)
    expected_attributes = expected_node.attributes
    actual_attributes = actual_node.attributes

    return false unless expected_attributes.length == actual_attributes.length
    expected_attributes.each do |expected_attribute|
      expected_name = expected_attribute.name
      if expected_attribute.ns?
        expected_ns_uri = expected_attribute.ns.href
        actual_attribute = actual_attributes.get_attribute_ns(expected_ns_uri, expected_name)
      else
        actual_attribute = actual_attributes.get_attribute(expected_name)
      end
      return false unless actual_attribute
      return false if expected_attribute.value != actual_attribute.value
    end
    true
  end

  def iterate_nodes(expected, actual, &block)
    expected_iter = NodeIterator.new(expected)
    actual_iter = NodeIterator.new(actual)
    while expected_iter.next?
      yield expected_iter.next, actual_iter.next
    end
  end

  class NodeIterator

    def initialize(node)
      @next_node = node.kind_of?(XML::Document) ? node.root : node
    end

    def find_next_node(node)
      next_node = if node.child?
        node.first
      elsif node.next?
        node.next
      elsif node.parent && node.parent.next?
        node.parent.next
      end

      return next_node if next_node.nil? || accept_node(next_node)
      find_next_node(next_node)
    end

    def accept_node(node)
      case
        when node.text?;     node.content !~ /^\s*$/
        when node.entity?;   false
        when node.notation?; false
        else;                true
      end
    end

    def next?
      @next_node
    end

    def next
      node = @next_node
      @next_node = find_next_node(node)
      node
    end

  end
end
