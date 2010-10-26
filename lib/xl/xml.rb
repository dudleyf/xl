require 'xml'

module Xl::Xml
  extend self

  NAMESPACES = {
    'ns'       => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'cp'       => 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    'ep'       => 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    'pr'       => 'http://schemas.openxmlformats.org/package/2006/relationships',
    'dr'       => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'dc'       => 'http://purl.org/dc/elements/1.1/',
    'dcterms'  => 'http://purl.org/dc/terms/',
    'dcmitype' => 'http://purl.org/dc/dcmitype/',
    'xsi'      => 'http://www.w3.org/2001/XMLSchema-instance',
    'vt'       => 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
    'xml'      => 'http://www.w3.org/XML/1998/namespace'
  }

  CONTENT_TYPES = {
    :theme => 'application/vnd.openxmlformats-officedocument.theme+xml',
    :styles => 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
    :rels => 'application/vnd.openxmlformats-package.relationships+xml',
    :workbook => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
    :extprops => 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
    :coreprops => 'application/vnd.openxmlformats-package.core-properties+xml',
    :strings => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
    :worksheet => 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
  }

  def read_xml(source, opts={})
    context = case source
      when String; XML::Parser::Context.string(source)
      when IO; XML::Parser::Context.io(source)
      when XML::Document; XML::Parser::Context.document(source)
    end

    context.options = XML::Parser::Options::NOENT | XML::Parser::Options::NOBLANKS

    if opts[:sax]
      parser = XML::SaxParser.new(context)
      parser.callbacks = opts[:sax]
    else
      parser = XML::Parser.new(context)
    end

    parser.parse.tap do |doc|
      if opts[:default_namespace_prefix]
        doc.root.namespaces.default_prefix = opts[:default_namespace_prefix]
      end
    end
  end

  def create_document(root)
    XML::Document.new.tap do |doc|
      doc.root = root
    end
  end

  def get_document_content(root)
    create_document(root).to_s
  end

  def make_node(name, attributes={})
    ns_prefix, node_name = split_node_name(name)

    XML::Node.new(node_name).tap do |node|
      extract_namespaces!(node, attributes)
      set_namespace(node, ns_prefix) if ns_prefix
      add_attributes(node, attributes)
    end
  end

  def make_subnode(parent, name, attributes={})
    ns_prefix, node_name = split_node_name(name)
    make_node(node_name, attributes).tap do |node|
      parent << node
      set_namespace(node, ns_prefix) if ns_prefix
    end
  end

  private

    def split_node_name(name)
      name.include?(':') ? name.split(':', 2) : [nil, name]
    end

    def add_attributes(node, attributes)
      attributes.each do |k,v|
        XML::Attr.new(node, k.to_s, v.to_s)
      end
    end

    def extract_namespaces!(node, attributes)
      attributes.each do |k,v|
        match = /xmlns(:(.*))?/.match(k.to_s)
        if match
          prefix = match[2]
          XML::Namespace.new(node, prefix, attributes.delete(k))
        end
      end
    end

    def set_namespace(node, prefix)

      ns = node.namespaces.find_by_prefix(prefix)
      node.namespaces.namespace = ns
    end

end

require 'xl/xml/reader'
require 'xl/xml/writer'

module Xl::Xml

  class << self
    include Reader
    include Writer
  end

end

