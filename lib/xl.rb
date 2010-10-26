dir = File.expand_path(File.dirname(__FILE__))
$LOAD_PATH.unshift(dir) unless $LOAD_PATH.include?(dir)

require 'rubygems'
require 'bundler/setup'

module Xl
  class XlError < StandardError; end
  class IncompleteRangeError < XlError; end
  class DataTypeError < XlError; end
  class RangeError < XlError; end
  class SheetTitleError < XlError; end
  class CellCoordinatesError < XlError; end
  class InsufficientCoordinatesError < XlError; end
  class NamedRangeError < XlError; end
  class ColumnStringIndexError < XlError; end
end

require 'xl/date_helper'
require 'xl/zip'
require 'xl/style'
require 'xl/named_range'
require 'xl/coordinates'
require 'xl/cell'
require 'xl/worksheet'
require 'xl/workbook'
require 'xl/xml'
