module Xl::Xml::Writer
end

require 'xl/xml/writer/string_table'
require 'xl/xml/writer/style_table'
require 'xl/xml/writer/theme'
require 'xl/xml/writer/worksheet'
require 'xl/xml/writer/workbook'

module Xl::Xml::Writer
  include StringTable
  include StyleTable
  include Theme
  include Worksheet
  include Workbook

  def save_workbook(workbook, file)
    Xl::Zip.write(file) do |archive|
      workbook.worksheets.each {|x| x.garbage_collect}

      shared_string_table = create_string_table(workbook)
      shared_style_table = create_style_table(workbook)

      archive.add_shared_strings(write_string_table(shared_string_table))
      archive.add_content_types(write_content_types(workbook))
      archive.add_root_rels(write_root_rels(workbook))
      archive.add_workbook_rels(write_workbook_rels(workbook))
      archive.add_app(write_properties_app(workbook))
      archive.add_core(write_properties_core(workbook.properties))
      archive.add_theme(write_theme)
      archive.add_style(write_style_table(shared_style_table))
      archive.add_workbook(write_workbook(workbook))

      workbook.worksheets.each_with_index do |sheet, i|
        archive.add_worksheet("sheet#{i+1}", write_worksheet(sheet, shared_string_table, shared_style_table))
        if sheet.relationships
          archive.add_worksheet_rels("sheet#{i+1}.xml.rels", write_worksheet_rels(sheet))
        end
      end
    end
  end
end
