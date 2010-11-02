# Describes cell associated properties (style, type, value, address, ...)
class Xl::Cell
  include Xl::Coordinates

  attr_accessor :column,
    :row,
    :value,
    :data_type,
    :parent,
    :xf_index,
    :hyperlink_rel,
    :style

  ERROR_CODES = {
    '#NULL!'  => 0,
    '#DIV/0!' => 1,
    '#VALUE!' => 2,
    '#REF!'   => 3,
    '#NAME?'  => 4,
    '#NUM!'   => 5,
    '#N/A'    => 6
  }

  TYPE_STRING = 's'
  TYPE_FORMULA = 'f'
  TYPE_NUMERIC = 'n'
  TYPE_BOOL = 'b'
  TYPE_NULL = 's'
  TYPE_INLINE = 'inlineStr'
  TYPE_ERROR = 'e'

  PERCENTAGE_RE = /^\-?[0-9]*\.?[0-9]*\s?\%$/
  TIME_RE       = /^(\d|[0-1]\d|2[0-3]):[0-5]\d(:[0-5]\d)?$/
  NUMERIC_RE    = /^\-?([0-9]+\.?[0-9]*|[0-9]*\.?[0-9]+)$/

  def initialize(worksheet, column, row, value=nil)
    @parent = worksheet
    @column = column.upcase
    @row = row
    @value = nil
    @hyperlink_rel = nil
    @data_type = TYPE_NULL

    self.value = value unless value.nil?
  end

  def string?;  data_type == TYPE_STRING; end
  def formula?; data_type == TYPE_FORMULA; end
  def numeric?; data_type == TYPE_NUMERIC; end
  def bool?;    data_type == TYPE_BOOL; end
  def null?;    data_type == TYPE_NULL; end
  def inline?;  data_type == TYPE_INLINE; end
  def error?;   data_type == TYPE_ERROR; end

  # Set the hyperlink held in the cell. Automatically sets the +value+ of the cell
  # with link text, but you can modify it afterwards by setting the +value+ property
  # and the hyperlink will remain.
  def hyperlink=(val)
    @hyperlink_rel ||= @parent.create_relationship("hyperlink")
    @hyperlink_rel.target = val
    @hyperlink_rel.target_mode = "External"
    self.value = val if @value.nil?
  end

  # Get the hyperlink held in the cell
  #
  # @return [String] the hyperlink target or ''
  def hyperlink
    @hyperlink_rel ? @hyperlink_rel.target : ''
  end

  # Get the id of the hyperlink held in the cell.
  def hyperlink_rel_id
    @hyperlink_rel ? @hyperlink_rel.id : nil
  end

  def inspect
    "<Cell %s.%s>" % [@parent.title, get_coordinate]
  end

  # Get the value held in the cell
  def value
    val = @value
    if format_as_date?(val)
      val = Xl::DateHelper.excel_to_ruby(val)
    end
    val
  end

  def raw_value
    @value
  end

  # Set the value held in the cell
  def value=(val)
    data_type = @data_type = data_type_for_value(val)
    if data_type == TYPE_STRING
      # percentage detection
      percentage_search = PERCENTAGE_RE.match(val)
      if percentage_search && val.strip != '%'
        val = val.gsub('%', '').to_f / 100.0
        set_value_explicit(val, TYPE_NUMERIC)
        style.number_format.format_code = Xl::NumberFormat::FORMAT_PERCENTAGE
        return true
      end

      # time detection
      # time_search = TIME_RE.match(val)
      # if time_search
      #   sep_count = val.count(':')
      #   if sep_count == 1
      #     h, m = val.split(':').map(&:to_i)
      #     s = 0
      #   elsif sep_count == 2
      #     h, m, s = val.split(':').map(&:to_i)
      #   end
      #   days = (h / 24.0) + (m / 1440.0) + (s / 86400.0)
      #   set_value_explicit(days, TYPE_NUMERIC)
      #   style.number_format.format_code = Xl::NumberFormat::FORMAT_DATE_TIME3
      #   return true
      # end
    elsif data_type == TYPE_NUMERIC
      # date detection
      if (Xl::DateHelper.datelike?(val))
        val = Xl::DateHelper.ruby_to_excel(val)
        set_value_explicit(val, TYPE_NUMERIC)
        style.number_format.format_code = Xl::NumberFormat::FORMAT_DATE_YYYYMMDD2
        return true
      end
    end
    set_value_explicit(val, data_type)
  end

  def has_style?
    !@style.nil?
  end

  def style
    @style ||= Xl::Style.new
  end

  def data_type_for_value(val)
    if val.nil?
      TYPE_NULL
    elsif val == true || val == false
      TYPE_BOOL
    elsif val.kind_of?(Fixnum) || val.kind_of?(Float)
      TYPE_NUMERIC
    elsif !val
      TYPE_STRING
    elsif Xl::DateHelper.datelike?(val)
      TYPE_NUMERIC
    elsif val.kind_of?(String) && /^=.*/.match(val)
      TYPE_FORMULA
    elsif NUMERIC_RE.match(val)
      TYPE_NUMERIC
    elsif ERROR_CODES.include?(val.strip)
      TYPE_ERROR
    else
      TYPE_STRING
    end
  end

  def set_value_explicit(val=nil, data_type=TYPE_STRING)
    case data_type
      when TYPE_INLINE, TYPE_STRING
        @value = check_string(val)
      when TYPE_FORMULA
        @value = val
      when TYPE_BOOL
        @value = !!val
      when TYPE_NUMERIC
        if (val.kind_of?(Fixnum) || val.kind_of?(Float))
          @value = val
        else
          begin
            @value = Integer(val)
          rescue
            @value = Float(val)
          end
        end
      else
        raise Xl::DataTypeError, "Invalid data type #{data_type}"
    end
    @data_type = data_type
  end

  def check_string(s)
    s = s[0..32767] # string must never be longer than 32,767 characters
    s.gsub!(/\r\n/, '\n') # newline must be \n, not \r\n
    s
  end

  # @return [String] the coordinate string for this cell
  def get_coordinate
    return '%s%s' % [column, row]
  end

  # @param [Fixnum] row
  # @param [Fixnum] column
  # @return [Cell] the cell at this Cell's address, offset by +row+ and +column+
  def offset(row=0, column=0)
    offset_column = get_column_letter(column_index_from_string(self.column) + column)
    offset_row = self.row + row

    @parent.cell('%s%s' % [offset_column, offset_row])
  end

  private

    def format_as_date?(val)
      has_style? &&
        style.number_format.date_format? &&
        (val.kind_of?(Fixnum) || val.kind_of?(Float))
    end


end
