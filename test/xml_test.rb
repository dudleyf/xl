require File.join(File.dirname(__FILE__), "test_helper")

class XmlTest < XlTestCase
  include Xl::Xml

  def test_read_xml_file
    file = test_data_file('writer/expected/core.xml')
    doc = read_xml(File.new(file))

    assert_xml_equal(File.new(file), doc)
  end

  def test_read_xml_string
    file = test_data_file('writer/expected/core.xml')
    doc = read_xml(File.read(file))

    assert_xml_equal(File.new(file), doc)
  end

  def test_node_creates_a_node
    node = make_node('foo')

    assert_kind_of(XML::Node, node)
    assert_equal 'foo', node.name
  end

  def test_node_attributes
    node = make_node('foo', :bar => 1, 'baz' => 2)

    assert_equal "1", node.attributes['bar']
    assert_equal "2", node.attributes['baz']
  end

  def test_node_namespace_definition
    node = make_node('foo', 'xmlns:bar' => 'http://bar.com')

    assert_nil node.attributes['xmlns:bar']
    assert_nil node.namespaces.namespace
    assert_equal 'http://bar.com', node.namespaces.find_by_prefix('bar').href
  end

  def test_node_default_namespace
    node = make_node('foo', 'xmlns' => 'http://bar.com')

    assert_nil node.attributes['xmlns']
    assert_equal 'http://bar.com', node.namespaces.default.href
  end

  def test_node_namespace_in_node_name
    node = make_node('bar:foo', 'xmlns:bar' => 'http://bar.com')

    assert node.name == 'foo'
    assert node.namespaces.namespace.prefix == 'bar'
  end

  def test_subnode
    node = make_node('foo')
    sub = make_subnode(node, 'bar')

    assert_equal(sub, node.children.first)
  end

end
