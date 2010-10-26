require 'digest/md5'

module Xl

  module Hashable
    include Comparable

    def crc
      Digest::MD5.hexdigest(self.to_yaml)
    end

    def <=>(other)
      self.crc <=> other.crc
    end

    def ==(other)
      self.crc == other.crc
    end
  end

  class Color
    include Hashable

    BLACK = 'FF000000'
    WHITE = 'FFFFFFFF'
    RED = 'FFFF0000'
    DARKRED = 'FF800000'
    BLUE = 'FF0000FF'
    DARKBLUE = 'FF000080'
    GREEN = 'FF00FF00'
    DARKGREEN = 'FF008000'
    YELLOW = 'FFFFFF00'
    DARKYELLOW = 'FF808000'

    attr_accessor :index

    def initialize(index)
      @index = index
    end
  end

  class Font
    include Hashable

    UNDERLINE_NONE = 'none'
    UNDERLINE_DOUBLE = 'double'
    UNDERLINE_DOUBLE_ACCOUNTING = 'doubleAccounting'
    UNDERLINE_SINGLE = 'single'
    UNDERLINE_SINGLE_ACCOUNTING = 'singleAccounting'

    attr_accessor :name, :size, :bold, :italic, :superscript, :subscript, :underline, :strikethrough, :color

    def initialize(opts={})
      @name = opts[:name] || 'Calibri'
      @size = opts[:size] || 11
      @bold = opts[:bold] || false
      @italic = opts[:italic] || false
      @superscript = opts[:superscript] || false
      @subscript = opts[:subscript] || false
      @underline = opts[:underline] || UNDERLINE_NONE
      @strikethrough = opts[:strikethrough] || false
      @color = opts[:color] || Color.new(Color::BLACK)
    end
  end

  class Fill
    include Hashable

    FILL_NONE = 'none'
    FILL_SOLID = 'solid'
    FILL_GRADIENT_LINEAR = 'linear'
    FILL_GRADIENT_PATH = 'path'
    FILL_PATTERN_DARKDOWN = 'darkDown'
    FILL_PATTERN_DARKGRAY = 'darkGray'
    FILL_PATTERN_DARKGRID = 'darkGrid'
    FILL_PATTERN_DARKHORIZONTAL = 'darkHorizontal'
    FILL_PATTERN_DARKTRELLIS = 'darkTrellis'
    FILL_PATTERN_DARKUP = 'darkUp'
    FILL_PATTERN_DARKVERTICAL = 'darkVertical'
    FILL_PATTERN_GRAY0625 = 'gray0625'
    FILL_PATTERN_GRAY125 = 'gray125'
    FILL_PATTERN_LIGHTDOWN = 'lightDown'
    FILL_PATTERN_LIGHTGRAY = 'lightGray'
    FILL_PATTERN_LIGHTGRID = 'lightGrid'
    FILL_PATTERN_LIGHTHORIZONTAL = 'lightHorizontal'
    FILL_PATTERN_LIGHTTRELLIS = 'lightTrellis'
    FILL_PATTERN_LIGHTUP = 'lightUp'
    FILL_PATTERN_LIGHTVERTICAL = 'lightVertical'
    FILL_PATTERN_MEDIUMGRAY = 'mediumGray'

    attr_accessor :fill_type, :rotation, :start_color, :end_color

    def initialize(opts={})
      @fill_type = opts[:fill_type] || FILL_NONE
      @rotation = opts[:rotation] || 0
      @start_color = opts[:start_color] || Color.new(Color::WHITE)
      @end_color = opts[:end_color] || Color.new(Color::BLACK)
    end
  end

  class Border
    include Hashable

    BORDER_NONE = 'none'
    BORDER_DASHDOT = 'dashDot'
    BORDER_DASHDOTDOT = 'dashDotDot'
    BORDER_DASHED = 'dashed'
    BORDER_DOTTED = 'dotted'
    BORDER_DOUBLE = 'double'
    BORDER_HAIR = 'hair'
    BORDER_MEDIUM = 'medium'
    BORDER_MEDIUMDASHDOT = 'mediumDashDot'
    BORDER_MEDIUMDASHDOTDOT = 'mediumDashDotDot'
    BORDER_MEDIUMDASHED = 'mediumDashed'
    BORDER_SLANTDASHDOT = 'slantDashDot'
    BORDER_THICK = 'thick'
    BORDER_THIN = 'thin'

    attr_accessor :border_style, :color

    def initialize(opts={})
      @border_style = opts[:border_style] || BORDER_NONE
      @color = opts[:color] || Color.new(Color::BLACK)
    end
  end

  class Borders
    include Hashable

    DIAGONAL_NONE = 0
    DIAGONAL_UP = 1
    DIAGONAL_DOWN = 2
    DIAGONAL_BOTH = 3

    attr_accessor :left, :right, :top, :bottom, :diagonal, :diagonal_direction, :all_borders, :outline, :inside, :vertical, :horizontal

    def initialize(opts={})
      @left = opts[:left] || Border.new
      @right = opts[:right] || Border.new
      @top = opts[:top] || Border.new
      @bottom = opts[:bottom] || Border.new
      @diagonal = opts[:diagonal] || Border.new
      @diagonal_direction = opts[:diagonal_direction] || DIAGONAL_NONE

      @all_borders = opts[:all_borders] || Border.new
      @outline = opts[:outline] || Border.new
      @inside = opts[:inside] || Border.new
      @vertical = opts[:vertical] || Border.new
      @horizontal = opts[:horizontal] || Border.new
    end
  end

  class Alignment
    include Hashable

    HORIZONTAL_GENERAL = 'general'
    HORIZONTAL_LEFT = 'left'
    HORIZONTAL_RIGHT = 'right'
    HORIZONTAL_CENTER = 'center'
    HORIZONTAL_CENTER_CONTINUOUS = 'centerContinuous'
    HORIZONTAL_JUSTIFY = 'justify'

    VERTICAL_BOTTOM = 'bottom'
    VERTICAL_TOP = 'top'
    VERTICAL_CENTER = 'center'
    VERTICAL_JUSTIFY = 'justify'

    attr_accessor :horizontal, :vertical, :text_rotation, :wrap_text, :shrink_to_fit, :indent

    def initialize(opts={})
      @horizontal = opts[:horizontal] || HORIZONTAL_GENERAL
      @vertical = opts[:vertical] || VERTICAL_BOTTOM
      @text_rotation = opts[:text_rotation] || 0
      @wrap_text = opts[:wrap_text] || false
      @shrink_to_fit = opts[:shrink_to_fit] || false
      @indent = opts[:indent] || 0
    end
  end

  class NumberFormat
    include Hashable

    FORMAT_GENERAL = 'General'

    FORMAT_TEXT = '@'

    FORMAT_NUMBER = '0'
    FORMAT_NUMBER_00 = '0.00'
    FORMAT_NUMBER_COMMA_SEPARATED1 = '#,##0.00'
    FORMAT_NUMBER_COMMA_SEPARATED2 = '#,##0.00_-'

    FORMAT_PERCENTAGE = '0%'
    FORMAT_PERCENTAGE_00 = '0.00%'

    FORMAT_DATE_YYYYMMDD2 = 'yyyy-mm-dd'
    FORMAT_DATE_YYYYMMDD = 'yy-mm-dd'
    FORMAT_DATE_DDMMYYYY = 'dd/mm/yy'
    FORMAT_DATE_DMYSLASH = 'd/m/y'
    FORMAT_DATE_DMYMINUS = 'd-m-y'
    FORMAT_DATE_DMMINUS = 'd-m'
    FORMAT_DATE_MYMINUS = 'm-y'
    FORMAT_DATE_XLSX14 = 'mm-dd-yy'
    FORMAT_DATE_XLSX15 = 'd-mmm-yy'
    FORMAT_DATE_XLSX16 = 'd-mmm'
    FORMAT_DATE_XLSX17 = 'mmm-yy'
    FORMAT_DATE_XLSX22 = 'm/d/yy h:mm'
    FORMAT_DATE_DATETIME = 'd/m/y h:mm'
    FORMAT_DATE_TIME1 = 'h:mm AM/PM'
    FORMAT_DATE_TIME2 = 'h:mm:ss AM/PM'
    FORMAT_DATE_TIME3 = 'h:mm'
    FORMAT_DATE_TIME4 = 'h:mm:ss'
    FORMAT_DATE_TIME5 = 'mm:ss'
    FORMAT_DATE_TIME6 = 'h:mm:ss'
    FORMAT_DATE_TIME7 = 'i:s.S'
    FORMAT_DATE_TIME8 = 'h:mm:ss@'
    FORMAT_DATE_YYYYMMDDSLASH = 'yy/mm/dd@'

    DATE_PATTERNS = [
      Regexp.compile(FORMAT_DATE_YYYYMMDD2) ,
      Regexp.compile(FORMAT_DATE_YYYYMMDD) ,
      Regexp.compile(FORMAT_DATE_DDMMYYYY) ,
      Regexp.compile(FORMAT_DATE_DMYSLASH) ,
      Regexp.compile(FORMAT_DATE_DMYMINUS) ,
      Regexp.compile(FORMAT_DATE_DMMINUS) ,
      Regexp.compile(FORMAT_DATE_MYMINUS) ,
      Regexp.compile(FORMAT_DATE_XLSX14) ,
      Regexp.compile(FORMAT_DATE_XLSX15) ,
      Regexp.compile(FORMAT_DATE_XLSX16) ,
      Regexp.compile(FORMAT_DATE_XLSX17) ,
      Regexp.compile(FORMAT_DATE_XLSX22) ,
      Regexp.compile(FORMAT_DATE_DATETIME) ,
      Regexp.compile(FORMAT_DATE_TIME1) ,
      Regexp.compile(FORMAT_DATE_TIME2) ,
      Regexp.compile(FORMAT_DATE_TIME3) ,
      Regexp.compile(FORMAT_DATE_TIME4) ,
      Regexp.compile(FORMAT_DATE_TIME5) ,
      Regexp.compile(FORMAT_DATE_TIME6) ,
      Regexp.compile(FORMAT_DATE_TIME7) ,
      Regexp.compile(FORMAT_DATE_TIME8) ,
      Regexp.compile(FORMAT_DATE_YYYYMMDDSLASH)
    ]

    FORMAT_CURRENCY_USD_SIMPLE = '"$"#,##0.00_-'
    FORMAT_CURRENCY_USD = '$#,##0_-'
    FORMAT_CURRENCY_EUR_SIMPLE = '[$EUR ]#,##0.00_-'

    BUILTIN_FORMATS = {
      0  => 'General',
      1  => '0',
      2  => '0.00',
      3  => '#,##0',
      4  => '#,##0.00',

      9  => '0%',
      10 => '0.00%',
      11 => '0.00E+00',
      12 => '# ?/?',
      13 => '# ??/??',
      14 => 'mm-dd-yy',
      15 => 'd-mmm-yy',
      16 => 'd-mmm',
      17 => 'mmm-yy',
      18 => 'h:mm AM/PM',
      19 => 'h:mm:ss AM/PM',
      20 => 'h:mm',
      21 => 'h:mm:ss',
      22 => 'm/d/yy h:mm',

      37 => '#,##0 (#,##0)',
      38 => '#,##0 [Red](#,##0)',
      39 => '#,##0.00(#,##0.00)',
      40 => '#,##0.00[Red](#,##0.00)',

      44 => '_("$"* #,##0.00_)_("$"* \(#,##0.00\)_("$"* "-"??_)_(@_)',
      45 => 'mm:ss',
      46 => '[h]:mm:ss',
      47 => 'mmss.0',
      48 => '##0.0E+0',
      49 => '@'
    }

    BUILTIN_FORMATS_REVERSE = {}.tap do |h|
      BUILTIN_FORMATS.each {|k,v| h[v] = k }
    end

    def initialize(opts={})
      @format_code = opts[:format_code] || FORMAT_GENERAL
      @format_index = opts[:format_index] || 0
    end

    def format_code=(code)
      @format_code = code
      @format_index = builtin_format_id(format_code)
    end

    def format_code
      @format_code
    end

    def builtin_format_code(index)
      BUILTIN_FORMATS[index]
    end

    def builtin_format_id(format)
      BUILTIN_FORMATS_REVERSE[format]
    end

    def builtin?(format=@format_code)
      BUILTIN_FORMATS.values.include?(format)
    end

    def date_format?(format=@format_code)
      DATE_PATTERNS.find {|x| x.match(format)}
    end
  end

  class Protection
    include Hashable

    PROTECTION_INHERIT = 'inherit'
    PROTECTION_PROTECTED = 'protected'
    PROTECTION_UNPROTECTED = 'unprotected'

    attr_accessor :locked, :hidden

    def initialize(opts={})
      @locked = opts[:locked] || PROTECTION_INHERIT
      @hidden = opts[:hidden] || PROTECTION_INHERIT
    end
  end

  class Style
    include Hashable

    attr_accessor :font, :fill, :borders, :alignment, :number_format, :protection

    def initialize(opts={})
      @font = opts[:font] || Font.new
      @fill = opts[:fill] || Fill.new
      @borders = opts[:borders] || Borders.new
      @alignment = opts[:alignment] || Alignment.new
      @number_format = opts[:number_format] || NumberFormat.new
      @protection = opts[:protection] || Protection.new
    end
  end
end
