require File.join(File.dirname(__FILE__), "test_helper")

class ZipTest < XlTestCase

  def setup
    @tmpdir = Dir.mktmpdir
  end

  def teardown
    FileUtils.rm_rf(@tmpdir) if @tmpdir && File.directory?(@tmpdir)
  end

  def test_archive_accessors
    filename = File.join(@tmpdir, 'test.zip')

    Xl::Zip.write(filename) do |z|

      z.add_from_string(Xl::Zip::ARC_APP, "foo")
      assert_equal z.get_from_name(Xl::Zip::ARC_APP), z.app

      z.add_from_string(Xl::Zip::ARC_CONTENT_TYPES, "foo")
      assert_equal z.get_from_name(Xl::Zip::ARC_CONTENT_TYPES), z.content_types

      z.add_from_string(Xl::Zip::ARC_ROOT_RELS, "foo")
      assert_equal z.get_from_name(Xl::Zip::ARC_ROOT_RELS), z.root_rels

      z.add_from_string(Xl::Zip::ARC_WORKBOOK_RELS, "foo")
      assert_equal z.get_from_name(Xl::Zip::ARC_WORKBOOK_RELS), z.workbook_rels

      z.add_from_string(Xl::Zip::ARC_CORE, "foo")
      assert_equal z.get_from_name(Xl::Zip::ARC_CORE), z.core

      z.add_from_string(Xl::Zip::ARC_WORKBOOK, "foo")
      assert_equal z.get_from_name(Xl::Zip::ARC_WORKBOOK), z.workbook

      z.add_from_string(Xl::Zip::ARC_STYLE, "foo")
      assert_equal z.get_from_name(Xl::Zip::ARC_STYLE), z.style

      z.add_from_string(Xl::Zip::ARC_THEME, "foo")
      assert_equal z.get_from_name(Xl::Zip::ARC_THEME), z.theme

      z.add_from_string(Xl::Zip::ARC_SHARED_STRINGS, "foo")
      assert_equal z.get_from_name(Xl::Zip::ARC_SHARED_STRINGS), z.shared_strings
    end
  end

  def test_worksheet
    filename = File.join(@tmpdir, 'test.zip')
    Xl::Zip.write(filename) do |archive|
      name = Xl::Zip::PACKAGE_WORKSHEETS + "/sheet1.xml"
      archive.add_from_string(name, "foo")
      assert_equal archive.get_from_name(name), archive.worksheet('sheet1')
    end
  end

  def test_write_zip
    filename = File.join(@tmpdir, 'test.zip')
    inner_filename = 'file.a'
    inner_content = 'here is the content'
    Xl::Zip.write(filename) do |archive|
      archive.add_from_string(inner_filename, inner_content)
    end

    Zip::ZipFile.open(filename) do |zip|
      assert zip.find_entry(inner_filename)
      assert_equal inner_content, zip.read(inner_filename)
    end
  end

  def test_read_zip
    filename = File.join(@tmpdir, 'test.zip')
    inner_filename = 'file.a'
    inner_content = 'here is the content'

    Zip::ZipFile.open(filename, Zip::ZipFile::CREATE) do |zip|
      zip.get_output_stream(inner_filename) do |out|
        out << inner_content
      end
    end

    read_content = ''
    Xl::Zip.new(filename) do |archive|
      read_content = archive.get_from_name(inner_filename)
    end

    assert_equal inner_content, read_content
  end
end
