# Represents a worksheet.
class Xl::Worksheet
  include Xl::Coordinates

  class Relationship
    TYPES = {
      'hyperlink' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
      #        'worksheet': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
      #        'sharedStrings': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
      #        'styles': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
      #        'theme': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
    }

    attr_accessor :type, :target, :target_mode, :id

    def initialize(reltype)
      raise RuntimeError, "Invalid relationship type #{reltype}" unless TYPES.include?(reltype)
      @type = TYPES[reltype]
      @target = ''
      @target_mode = ''
      @id = ''
    end
  end

  # @private
  class PageSetup; end

  # @private
  class HeaderFooter; end

  # @private
  class SheetView; end

  # @private
  class RowDimension
    attr_accessor :row_index
    attr_accessor :height
    attr_accessor :outline_level
    attr_accessor :style_index

    attr_accessor :visible
    alias_method :visible?, :visible

    attr_accessor :collapsed
    alias_method :collapsed?, :collapsed

    def initialize(index=0)
      @row_index = index
      @height = -1
      @visible = true
      @outline_level = 0
      @collapsed = false
      @style_index = nil
    end
  end

  # @private
  class ColumnDimension
    attr_accessor :column_index
    attr_accessor :width
    attr_accessor :outline_level
    attr_accessor :style_index

    attr_accessor :visible
    alias_method :visible?, :visible

    attr_accessor :collapsed
    alias_method :collapsed?, :collapsed

    attr_accessor :auto_size
    alias_method :auto_size?, :auto_size

    def initialize(index='A')
      @column_index = index
      @width = -1
      @auto_size = false
      @visible = true
      @outline_level = 0
      @collapsed = false
      @style_index = 0
    end
  end

  # @private
  class PageMargins
    attr_accessor :left
    attr_accessor :right
    attr_accessor :top
    attr_accessor :bottom
    attr_accessor :header
    attr_accessor :footer

    def initialize
      @left = @right = 0.7
      @top = @bottom = 0.75
      @header = @footer = 0.3
    end
  end

  # @private
  class SheetProtection
    attr_accessor :sheet
    attr_accessor :objects
    attr_accessor :scenarios
    attr_accessor :format_cells
    attr_accessor :format_columns
    attr_accessor :format_rows
    attr_accessor :insert_columns
    attr_accessor :insert_rows
    attr_accessor :insert_hyperlinks
    attr_accessor :delete_columns
    attr_accessor :delete_rows
    attr_accessor :select_locked_cells
    attr_accessor :sort
    attr_accessor :auto_filter
    attr_accessor :pivot_tables
    attr_accessor :select_unlocked_cells

    def initialize
      @password = ''
    end

    # Create a password hash from a given string.
    #
    # This method is based on the algorithm provided by
    # Daniel Rentz of OpenOffice and the PEAR package
    # Spreadsheet_Excel_Writer by Xavier Noguer <xnoguer@rezebra.com>.
    def hash_password(plain_password)
      password = 0
      shift = 1

      0.upto(plain_password.length-1) do |i|
        val = plain_password[i].ord << shift
        rotated_bits = val >> 15
        val &= 0x7fff
        password ^= (val | rotated_bits)
        shift += 1
      end

      password ^= plain_password.length
      password ^= 0xCE4B
      password.to_s(16).upcase
    end

    def set_password(value='', already_hashed=false)
      @password = already_hashed ? value : hash_password(value)
    end

    def password=(value)
      set_password(value, false)
    end

    def password
      @password
    end
  end

  BREAK_NONE = 0
  BREAK_ROW = 1
  BREAK_COLUMN = 2

  SHEETSTATE_VISIBLE = 'visible'
  SHEETSTATE_HIDDEN = 'hidden'
  SHEETSTATE_VERYHIDDEN = 'veryHidden'

  attr_accessor :parent
  attr_accessor :row_dimensions
  attr_accessor :column_dimensions
  attr_accessor :cells
  attr_accessor :merged_cells
  attr_accessor :styles
  attr_accessor :relationships
  attr_accessor :active_cell
  attr_accessor :selected_cell
  attr_accessor :sheet_state
  attr_accessor :page_setup
  attr_accessor :page_margins
  attr_accessor :header_footer
  attr_accessor :sheet_view
  attr_accessor :protection
  attr_accessor :default_row_dimension
  attr_accessor :default_column_dimension

  attr_accessor :show_gridlines
  alias_method :show_gridlines?, :show_gridlines

  attr_accessor :print_gridlines
  alias_method :print_gridlines?, :print_gridlines

  attr_accessor :show_summary_below
  alias_method :show_summary_below?, :show_summary_below

  attr_accessor :show_summary_right
  alias_method :show_summary_right?, :show_summary_right

  def initialize(parent_workbook, title='Sheet')
    @parent = parent_workbook

    @row_dimensions = {}
    @column_dimensions = {}

    unless title
      self.title = 'Sheet%d' % [1 + @parent.worksheets.length]
    else
      self.title = title
    end

    @cells = {}
    @merged_cells = []
    @styles = {}
    @relationships = []
    @selected_cell = 'A1'
    @active_cell = 'A1'
    @sheet_state = SHEETSTATE_VISIBLE
    @page_setup = PageSetup.new
    @page_margins = PageMargins.new
    @header_footer = HeaderFooter.new
    @sheet_view = SheetView.new
    @protection = SheetProtection.new
    @default_row_dimension = RowDimension.new
    @default_column_dimension = ColumnDimension.new
    @show_gridlines = true
    @print_gridlines = false
    @show_summary_below = true
    @show_summary_right = true
  end

  def inspect
    "<Worksheet #{@title}>"
  end

  # Remove empty cells from the cell collection.
  def garbage_collect
    delete_list = []
    @cells.each do |coordinate, cell|
      delete_list << coordinate if cell.value == ''
    end

    delete_list.each {|x| @cells.delete(x)}
  end

  def get_cell_collection
    @cells.values
  end

  # Set the title of the worksheet. Limited to 31 characters, no special characters.
  #
  # @param [String] new_title the new title of the worksheet
  # @raise [SheetTitleError] if the +new_title+ is longer than 31 characters
  #   or contains punctuation characters
  def title=(new_title)
    if /(\*|\:|\/|\\|\?|\[|\])/.match(new_title)
      raise Xl::SheetTitleError, "Invalid character found in sheet title"
    end

    if new_title.length > 31
      raise Xl::SheetTitleError, "Maximum 31 characters allowed in sheet title"
    end

    if @parent.get_sheet_by_name(new_title)
      # if there's already a sheet with this name, append the next lowest integer
      i = 1
      i += 1 while @parent.get_sheet_by_name('%s%d' % [new_title, i])
      @title = '%s%d' % [new_title, i]
    else
      @title = new_title
    end
  end

  # Get the title of the worksheet.
  #
  # @return the title of the worksheet
  def title
    @title
  end

  # Returns a cell object based on the given coordinates.
  #
  # Cells are kept in a dictionary which is empty at the worksheet creation.
  # Calling +cell+ creates the cell in memory when they are first accessed to
  # reduce memory usage.
  #
  # @example Return the cell at row 15, column 0
  #   sheet.cell('A15') # or
  #   sheet.cell(15, 0)
  #
  # @param [String, Fixnum] coordinate_or_row if this is a String, it's
  #   interpreted as a coordinate in 'A13' notation. If it's a Fixnum,
  #   it's the row number and +column+ should also be provided.
  # @param [Fixnum] column the column number of the cell if using row
  #   and column numbers.
  # @raise [InsufficientCoordinatesException] when coordinate or row and column
  #   are not given.
  # @return [Cell] the cell at the given coordinates
  def cell(coordinate_or_row, column=nil)
    if coordinate_or_row.kind_of?(String)
      coordinate = coordinate_or_row
    else
      raise InsufficientCoordinatesError if column.nil?
      coordinate = '%s%s' % [get_column_letter(column), coordinate_or_row]
    end

    unless @cells.has_key?(coordinate)
      column, row = coordinate_from_string(coordinate)
      @cells[coordinate] = Xl::Cell.new(self, column, row)

      unless @column_dimensions.has_key?(column)
        @column_dimensions[column] = ColumnDimension.new(column)
      end

      unless @row_dimensions.has_key?(row)
        @row_dimensions[row] = RowDimension.new(row)
      end
    end

    @cells[coordinate]
  end

  # @return the highest numbered row index in the worksheet
  def get_highest_row
    max = 1
    @row_dimensions.values.each do |x|
      max = x.row_index if x.row_index > max
    end
    max
  end

  # @return the highest numbered column index in the worksheet
  def get_highest_column
    max = 1
    @column_dimensions.values.each do |x|
      c = column_index_from_string(x.column_index)
      max = c if c > max
    end
    max
  end

  #
  def calculate_dimension
    'A1:%s%d' % [get_column_letter(get_highest_column), get_highest_row]
  end

  # Find a 2D array of cells representing the given range.
  def range(*args)
    if args.length == 4
      range_from_rows_and_columns(*args)
    else
      range_string, row_offset, col_offset = args
      if coordinate?(range_string)
        cell(range_string)
      elsif range_string.include?(':') # R1C1 notation
        range_from_coordinates(range_string, row_offset, col_offset)
      else
        range_from_named_range(range_string, row_offset, col_offset)
      end
    end
  end

  # @private
  def range_from_rows_and_columns(min_row, min_col, max_row, max_col)
    (min_row..max_row).map do |r|
      (min_col..max_col).map {|c| cell(r, c)}
    end
  end

  # @private
  def range_from_coordinates(range_string, row_offset=nil, column_offset=nil)
    min_range, max_range = range_string.split(':')

    min_col, min_row = coordinate_from_string(min_range)
    max_col, max_row = coordinate_from_string(max_range)

    min_col = column_index_from_string(min_col)
    max_col = column_index_from_string(max_col)

    if column_offset
      min_col += column_offset
      max_col += column_offset
    end

    if row_offset
      min_row += row_offset
      max_row += row_offset
    end

    range_from_rows_and_columns(min_row, min_col, max_row, max_col)
  end

  # @private
  def range_from_named_range(range_string, row_offset, column_offset)
    named_range = @parent.get_named_range(range_string)

    if named_range.nil?
      raise RuntimeError, "#{range_string} is not a valid range name"
    end

    if named_range.worksheet != self
      raise RuntimeError, "Range #{range_string} is not defined on worksheet #{@title}"
    end

    cell(named_range.range)
  end

  # Get the {Style} from the cell at the given coordinate, creating
  # a new {Style} if necessary.
  #
  # @param [String] coordinate the coordinate of the cell to retrieve the style from
  # @return [Style] the style of the cell at +coordinate+
  def get_style(coordinate)
    unless @styles.has_key?(coordinate)
      @styles[coordinate] = Xl::Style.new
    end

    @styles[coordinate]
  end

  # Create a relationship.
  #
  # @param [Fixnum] rel_type the relationship type to create
  # @return [Relationship]
  def create_relationship(rel_type)
    rel = Relationship.new(rel_type)
    @relationships << rel
    rel_id = @relationships.index(rel)
    rel.id = 'rId' + (rel_id + 1).to_s
    @relationships[rel_id]
  end

  # Merge cells.
  #
  # @param (see Xl::Worksheet#range)
  def merge(*args)
    range = range(*args)
    topleft = range.first.first.get_coordinate
    bottomright = range.last.last.get_coordinate

    self.merged_cells << "#{topleft}:#{bottomright}"
  end
end
