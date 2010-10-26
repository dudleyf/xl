TEST_ROOT = File.expand_path(File.dirname(__FILE__))
TEST_SUPPORT_DIR = File.join(TEST_ROOT, 'support')
TEST_DATA_DIR = File.join(TEST_SUPPORT_DIR, 'data')
TEST_SCHEMA_DIR = File.join(TEST_SUPPORT_DIR, 'schemas')

$LOAD_PATH.unshift TEST_SUPPORT_DIR unless $LOAD_PATH.include?(TEST_SUPPORT_DIR)
require File.join(TEST_ROOT, '../lib/xl')

require 'test/unit'
require 'diff/lcs'
require 'redgreen'
require 'zip/zip'

require 'xml_test_helper'

module TestHelper

  def test_data_file(*path)
    File.join(TEST_DATA_DIR, *path)
  end

  def test_data(path)
    File.read(test_data_file(path))
  end

  def validate(xml, schema_name)
    xml = parse_xml(xml)
    schema_file = File.open(File.join(TEST_SCHEMA_DIR, schema_filename))
    base_uri = "file://#{File.dirname(schema_file)}"
    schema_doc = XML::Document.io(schema_file, :base_uri => base_uri)
    schema = XML::Schema.document(schema_doc)

    xml.validate_schema(schema)
  end
end

class XlTestCase < Test::Unit::TestCase
  include TestHelper
  include XmlTestHelper
end

