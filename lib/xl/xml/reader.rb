module Xl::Xml::Reader
end

require 'xl/xml/reader/strings'
require 'xl/xml/reader/styles'
require 'xl/xml/reader/worksheet'
require 'xl/xml/reader/workbook'

module Xl::Xml::Reader
  include Strings
  include Styles
  include Worksheet
  include Workbook

  # Read a workbook from the given filename.
  #
  # @param [String] filename the path to open
  # @return [Workbook] the workbook
  def load_workbook(filename)
    wb = Xl::Workbook.new

    Xl::Zip.new(filename) do |archive|
      wb.properties = read_properties_core(archive.core)
      wb.worksheets = []
      sheet_names = read_sheets_titles(archive.app)
      string_table = archive.has_shared_strings? ? read_string_table(archive.shared_strings) : {}
      style_table = read_style_table(archive.style)
      sheet_names.each_with_index do |sheet_name, i|
        ws = read_worksheet(archive.worksheet("sheet#{i+1}"), wb, sheet_name, string_table, style_table)
        wb.add_sheet(ws, i)
      end
      wb.named_ranges = read_named_ranges(archive.workbook, wb)
    end

    wb
  end

end
