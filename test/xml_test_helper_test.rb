# Stolen from testunitxml
# Original Copyright © 2006 Henrik Mårtensson

require File.join(File.dirname(__FILE__), "test_helper")

class XmlTestHelperTest < Test::Unit::TestCase
  include TestHelper
  include XmlTestHelper

  def assert_xml_equal_fails(expected, actual)
    assert_raise(Test::Unit::AssertionFailedError) do
      assert_xml_equal(expected, actual)
    end
  end

  def setup
    @file = test_data_file('xml_test_helper_test.xml')
  end

  def test_assert_xml_equal_document
    doc1 = XML::Document.file(@file)
    doc2 = XML::Document.file(@file)
    assert_xml_equal(doc1, doc2)
  end

  def test_assert_xml_equal_io
    io1 = File.new(@file)
    io2 = File.new(@file)
    assert_xml_equal(io1, io2)
  end

  def test_assert_xml_equal_string
    str1 = File.read(@file)
    str2 = File.read(@file)
    assert_xml_equal(str1, str2)
  end

  def test_assert_xml_equal_string_and_doc
    str1 = File.read(@file)
    doc1 = XML::Document.file(@file)
    assert_xml_equal(str1, doc1)
  end

  def test_assert_xml_equal_node
    node1 = %(<t:root xmlns:t="urn:x-hm:test" xmlns:x="urn:x-hm:test2" id="a" t:type="test1"/>)
    assert_xml_equal(node1, node1)

    assert_xml_equal_fails(node1, %(<root xmlns:t="urn:x-hm:test" xmlns:x="urn:x-hm:test2" id="a" t:type="test1"/>))
    assert_xml_equal_fails(node1, %(<t:root xmlns:t="urn:x-hm:other" xmlns:x="urn:x-hm:test2" id="a" t:type="test1"/>))
    assert_xml_equal_fails(node1, %(<t:root xmlns:t="urn:x-hm:test" xmlns:x="urn:x-hm:test2" id="a" x:type="test1"/>))
    assert_xml_equal_fails(node1, %(<t:root xmlns:t="urn:x-hm:test" xmlns:x="urn:x-hm:test2" id="a" t:type="test2"/>))
    assert_xml_equal(node1, %(<t:root  id="a" xmlns:t="urn:x-hm:test" xmlns:x="urn:x-hm:test2" t:type="test1"/>))
    assert_xml_equal(node1, %(<s:root xmlns:s="urn:x-hm:test" xmlns:x="urn:x-hm:test" id="a" s:type="test1"/>))
    assert_xml_equal(node1, %(<t:root xmlns:t="urn:x-hm:test" id="a" t:type="test1"/>))
  end

  def test_xml_not_equal
    n = "<a></a>"
    assert_xml_not_equal(n, "<b></b>")
    assert_xml_not_equal(n, "<c></c>")
  end

  def test_assert_xml_equal_text
    text1 = XML::Node.new_text(' Q')
    assert_xml_equal(text1, text1)
#    assert_xml_equal(text1, XML::Node.new_text(' &#81;'))
    assert_xml_equal(text1, XML::Node.new_text(' Q'))
    assert_xml_equal_fails(text1, XML::Node.new_text('  Q '))
  end

  def test_assert_xml_equal_cdata
    cdata1 = XML::Node.new_cdata('Test text')
    cdata2 = XML::Node.new_cdata('Test \ntext')
    assert_xml_equal(cdata1, cdata1)
    assert_xml_equal_fails(cdata1, cdata2)
  end

  def test_assert_xml_equal_comment
    comment1 = XML::Node.new_comment("This is a comment")
    assert_xml_equal(comment1, comment1)
    assert_xml_equal_fails(comment1, XML::Node.new_comment("This is another comment"))
  end

  def test_assert_xml_equal_whitespace
    whitespace1 = XML::Document.string(%Q{<r><a>Some <b>text</b>.</a></r>})
    whitespace2 = XML::Document.string(%Q{<r> <a>Some <b>text</b>.</a>\n  \n  </r>})
    whitespace3 = XML::Document.string(%Q{<r><a>Some <b> text</b>.</a></r>})
    assert_xml_equal(whitespace1, whitespace1)
    assert_xml_equal(whitespace1, whitespace2)
    assert_xml_equal_fails(whitespace1, whitespace3)
  end

end

