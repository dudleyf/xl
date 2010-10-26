# The main workbook object
class Xl::Workbook

  class DocumentProperties
    attr_accessor :creator,
      :last_modified_by,
      :created,
      :modified,
      :title,
      :subject,
      :description,
      :keywords,
      :category,
      :company

    def initialize(opts={})
      @creator = opts[:creator] || 'Unknown'
      @last_modified_by = opts[:last_modified_by] || @creator
      @created = opts[:created] ||  Time.now
      @modified = opts[:modified] ||  Time.now
      @title = opts[:title] || 'Untitled'
      @subject = opts[:subject] || ''
      @description = opts[:description] || ''
      @keywords = opts[:keywords] || ''
      @category = opts[:category] || ''
      @company = opts[:company] || 'Microsoft Corporation'
    end
  end

  class DocumentSecurity

    attr_accessor :lock_revision,
      :lock_structure,
      :lock_windows,
      :revision_password,
      :workbook_password

    def initialize(opts={})
      @lock_revision = opts[:lock_revision] || false
      @lock_structure = opts[:lock_structure] || false
      @lock_windows = opts[:lock_windows] || false
      @revision_password = opts[:revision_password] || ''
      @workbook_password = opts[:workbook_password] || ''
    end

  end

  attr_accessor :worksheets
  attr_accessor :active_sheet_index
  attr_accessor :named_ranges
  attr_accessor :properties
  attr_accessor :style
  attr_accessor :security

  def initialize(opts={})
    @worksheets = []
#    @worksheets << Xl::Worksheet.new(self)
    @active_sheet_index = 0
    @named_ranges = []
    @properties = DocumentProperties.new
    @style = Xl::Style.new
    @security = DocumentSecurity.new
  end

  # The current active worksheet.
  #
  # @return [Worksheet] the current active worksheet
  def get_active_sheet
    @worksheets[@active_sheet_index]
  end

  # Creates a worksheet at an optional index.
  #
  # @param [optional, Fixnum] index position at which the sheet will be inserted
  def create_sheet(index=nil)
    Xl::Worksheet.new(self).tap do |new_sheet|
      add_sheet(new_sheet, index)
    end
  end

  # Add a sheet to the workbook at the given +index+, or at the end of
  # the sheets if no +index+ is given.
  #
  # @param [Worksheet] sheet the {Worksheet} to add
  # @param [optional, Fixnum] index the position at which the sheet will be inserted
  def add_sheet(sheet, index=nil)
    index ||= @worksheets.length
    @worksheets.insert(index, sheet)
  end

  # Remove a sheet from the workbook.
  #
  # @param [Worksheet] sheet the {Worksheet} to remove
  def remove_sheet(sheet)
    @worksheets.delete(sheet)
  end

  # Find a worksheet by its name.
  #
  # @param [String] name the name of the worksheet to look for
  # @return [Worksheet] if the workbook contains a sheet with the given +name+,
  #   returns that worksheet
  # @return [nil] if the workbook doesn't contain a sheet with the given +name+,
  #   returns nil
  def get_sheet_by_name(name)
    @worksheets.find {|x| x.title == name}
  end

  # Get the index of a worksheet.
  #
  # @param [Worksheet] the sheet whose index will be returned
  # @return [Fixnum] the index of the given +sheet+
  # @return [nil] if the sheet isn't found in the workbook
  def get_index(sheet)
    @worksheets.index(sheet)
  end

  # Get the names of the workbook's sheets.
  #
  # @return [Array<String>] the names of the worksheets in the workbook,
  #   in the same order as the worksheets
  def get_sheet_names
    @worksheets.map {|x| x.title}
  end

  # Add a new {NamedRange} to the workbook.
  #
  # @param [String] name the name of the new range
  # @param [Worksheet] sheet the worksheet on which to create the range
  # @param [String] range a string representing the range
  # @return [NamedRange] a new {NamedRange} object representing the given +range+
  def create_named_range(name, sheet, range)
    add_named_range(Xl::NamedRange.new(name, sheet, range))
  end

  # Get all of the workbook's {NamedRange}s.
  #
  # @return [Array<NamedRange>] this workbook's named ranges
  def get_named_ranges
    @named_ranges
  end

  # Add an existing {NamedRange} to the workbook.
  #
  # @param [NamedRange] range the named range to add to the workbook
  def add_named_range(range)
    @named_ranges << range
  end

  # Find a {NamedRange} by name.
  #
  # @param [String] name
  # @return [NamedRange] if the workbook has a named range with the given +name+,
  #   return it.
  # @return [nil] if the workbook doesn't have a range with the given +name+,
  #   return nil
  def get_named_range(name)
    @named_ranges.find {|x| x.name == name}
  end

  # Remove the given +range+ from the workbook.
  #
  # @param [NamedRange] range the range to remove
  def remove_named_range(range)
    @named_ranges.delete(range)
  end

end
