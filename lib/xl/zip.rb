require 'zip/zip'

class Xl::Zip
  PACKAGE_PROPS = 'docProps'
  PACKAGE_XL = 'xl'
  PACKAGE_RELS = '_rels'
  PACKAGE_THEME = PACKAGE_XL + '/' + 'theme'
  PACKAGE_WORKSHEETS = PACKAGE_XL + '/' + 'worksheets'

  ARC_CONTENT_TYPES = '[Content_Types].xml'
  ARC_ROOT_RELS = PACKAGE_RELS + '/.rels'
  ARC_WORKBOOK_RELS = PACKAGE_XL + '/' + PACKAGE_RELS + '/workbook.xml.rels'
  ARC_CORE = PACKAGE_PROPS + '/core.xml'
  ARC_APP = PACKAGE_PROPS + '/app.xml'
  ARC_WORKBOOK = PACKAGE_XL + '/workbook.xml'
  ARC_STYLE = PACKAGE_XL + '/styles.xml'
  ARC_THEME = PACKAGE_THEME + '/theme1.xml'
  ARC_SHARED_STRINGS = PACKAGE_XL + '/sharedStrings.xml'

  class << self
    def read(filename, &block)
      new(filename, false, &block)
    end

    def write(filename, &block)
      new(filename, true, &block)
    end
  end

  def initialize(filename, write=false, &block)
    @filename = filename
    @zipfile = write ?
      Zip::ZipFile.new(@filename, Zip::ZipFile::CREATE) :
      Zip::ZipFile.new(@filename)

    if block_given?
      begin
        yield self
      ensure
        close
      end
    end
  end

  def worksheet_path(path)
    File.join(PACKAGE_WORKSHEETS, path + '.xml')
  end

  def worksheet(path)
    get_from_name(worksheet_path(path))
  end

  def add_worksheet(path, content)
    add_from_string(worksheet_path(path), content)
  end

  def worksheet_rels(path)
    get_from_name(File.join(PACKAGE_WORKSHEETS, '_rels', path))
  end

  def add_worksheet_rels(path, content)
    add_from_string(File.join(PACKAGE_WORKSHEETS, '_rels', path), content)
  end

  %w[content_types root_rels workbook_rels core app workbook style theme shared_strings].each do |meth|
    class_eval <<-RUBY, __FILE__, __LINE__ + 1
      def #{meth}
        get_from_name(ARC_#{meth.upcase})
      end

      def add_#{meth}(content)
        add_from_string(ARC_#{meth.upcase}, content)
      end
    RUBY
  end

  def get_from_name(arc_name)
    @zipfile.read(arc_name)
  end

  def has_shared_strings?
    is_in_archive?(ARC_SHARED_STRINGS)
  end

  def is_in_archive?(arc_name)
    @zipfile.find_entry(arc_name)
  end

  def directory_is_in_archive?(dir)
    entry = @zipfile.find_entry(dir)
    entry && entry.directory?
  end

  def add_from_string(arc_name, content)
    mkdir_p(File.dirname(arc_name))
    @zipfile.get_output_stream(arc_name) do |out|
      out << content
    end
  end

  def add_from_file(arc_name, content)
    mkdir_p(File.dirname(arc_name))
    @zipfile.add(arc_name, content)
  end

  def close
    @zipfile.close
  end

  def mkdir(dir)
    @zipfile.mkdir(dir) unless directory_is_in_archive?(dir)
  end

  def mkdir_p(dirname)
    dirs = dirname.split('/')
    path = dirs.shift
    mkdir(path)
    dirs.each do |dir|
      path = File.join(path, dir)
      mkdir(path)
    end
  end
end
