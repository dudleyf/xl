require 'digest/md5'

module Xl

  module Attributes

    def self.included(base)
      base.extend(ClassMethods)
    end

    def attributes
      @attributes
    end

    def write_attribute(attr_name, value)
      @attributes ||= {}
      @attributes[attr_name.to_s] = value
    end

    def read_attribute(attr_name)
      @attributes ||= {}
      @attributes[attr_name.to_s]
    end

    def ==(other)
      @attributes == other.attributes
    end

    def eql?(other)
      self == other
    end

    def hash
      @attributes.hash
    end

    module ClassMethods

      def attribute(attr_name)
        attr_name = attr_name.to_s
        class_eval <<-RUBY, __FILE__, __LINE__
          def #{attr_name}
            read_attribute("#{attr_name}")
          end

          def #{attr_name}=(value)
            write_attribute("#{attr_name}", value)
          end

          def #{attr_name}?
            !!read_attribute("#{attr_name}")
          end
        RUBY
      end

    end

  end

  class Color
    include Attributes

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
    BROWN = 'FFA52A2A'
    CYAN = 'FF00FFFF'
    GRAY = 'FF808080'
    MAGENTA = 'FFFF00FF'
    ORANGE = 'FFFFA500'
    PINK = 'FFFFC0CB'
    PURPLE = 'FF800080'

    attribute :rgb

    def initialize(rgb)
      self.rgb = rgb
    end
  end

  class Font
    include Attributes

    UNDERLINE_NONE = 'none'
    UNDERLINE_DOUBLE = 'double'
    UNDERLINE_DOUBLE_ACCOUNTING = 'doubleAccounting'
    UNDERLINE_SINGLE = 'single'
    UNDERLINE_SINGLE_ACCOUNTING = 'singleAccounting'

    attribute :name
    attribute :size
    attribute :bold
    attribute :italic
    attribute :superscript
    attribute :subscript
    attribute :underline
    attribute :strikethrough
    attribute :outline
    attribute :shadow
    attribute :color

    def initialize(opts={})
      self.name = opts[:name] || 'Calibri'
      self.size = opts[:size] || 11
      self.bold = opts[:bold] || false
      self.italic = opts[:italic] || false
      self.superscript = opts[:superscript] || false
      self.subscript = opts[:subscript] || false
      self.underline = opts[:underline] || nil
      self.strikethrough = opts[:strikethrough] || false
      self.outline = opts[:outline] || false
      self.shadow = opts[:shadow] || false
      self.color = opts[:color] || Color.new(Color::BLACK)
    end
  end

  class Fill
    include Attributes

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

    attribute :fill_type
    attribute :rotation
    attribute :start_color
    attribute :end_color

    def initialize(opts={})
      self.fill_type = opts[:fill_type] || FILL_NONE
      self.rotation = opts[:rotation] || 0
      self.start_color = opts[:start_color] || Color.new(Color::WHITE)
      self.end_color = opts[:end_color] || Color.new(Color::BLACK)
    end
  end

  class Border
    include Attributes

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

    attribute :border_style
    attribute :color

    def initialize(opts={})
      self.border_style = opts[:border_style] || BORDER_NONE
      self.color = opts[:color] || Color.new(Color::BLACK)
    end
  end

  class Borders
    include Attributes

    DIAGONAL_NONE = 0
    DIAGONAL_UP = 1
    DIAGONAL_DOWN = 2
    DIAGONAL_BOTH = 3

    attribute :left
    attribute :right
    attribute :top
    attribute :bottom
    attribute :diagonal
    attribute :diagonal_direction
    attribute :all_borders
    attribute :outline
    attribute :inside
    attribute :vertical
    attribute :horizontal

    def initialize(opts={})
      self.left = opts[:left] || Border.new
      self.right = opts[:right] || Border.new
      self.top = opts[:top] || Border.new
      self.bottom = opts[:bottom] || Border.new
      self.diagonal = opts[:diagonal] || Border.new
      self.diagonal_direction = opts[:diagonal_direction] || DIAGONAL_NONE

      self.all_borders = opts[:all_borders] || Border.new
      self.outline = opts[:outline] || Border.new
      self.inside = opts[:inside] || Border.new
      self.vertical = opts[:vertical] || Border.new
      self.horizontal = opts[:horizontal] || Border.new
    end
  end

  class Alignment
    include Attributes

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

    attribute :horizontal
    attribute :vertical
    attribute :text_rotation
    attribute :wrap_text
    attribute :shrink_to_fit
    attribute :indent

    def initialize(opts={})
      self.horizontal = opts[:horizontal] || HORIZONTAL_GENERAL
      self.vertical = opts[:vertical] || VERTICAL_BOTTOM
      self.text_rotation = opts[:text_rotation] || 0
      self.wrap_text = opts[:wrap_text] || false
      self.shrink_to_fit = opts[:shrink_to_fit] || false
      self.indent = opts[:indent] || 0
    end
  end

  class NumberFormat
    include Attributes

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

    attribute :format_code
    attribute :format_index

    def initialize(opts={})
      self.format_code = opts[:format_code] || FORMAT_GENERAL
      self.format_index = opts[:format_index] || 0
    end

    def format_code=(code)
      write_attribute('format_code', code)
      write_attribute('format_index', builtin_format_id(format_code))
    end

    def builtin_format_code(index)
      BUILTIN_FORMATS[index]
    end

    def builtin_format_id(format)
      BUILTIN_FORMATS_REVERSE[format]
    end

    def builtin?(format=self.format_code)
      BUILTIN_FORMATS.values.include?(format)
    end

    def date_format?(format=self.format_code)
      DATE_PATTERNS.find {|x| x.match(format)}
    end
  end

  class Protection
    include Attributes

    PROTECTION_INHERIT = 'inherit'
    PROTECTION_PROTECTED = 'protected'
    PROTECTION_UNPROTECTED = 'unprotected'

    attribute :locked
    attribute :hidden

    def initialize(opts={})
      self.locked = opts[:locked] || PROTECTION_INHERIT
      self.hidden = opts[:hidden] || PROTECTION_INHERIT
    end
  end

  class Style
    include Attributes

    def self.default_font
      @default_font = Font.new({
        :name => 'Calibri',
        :size => 11
      })
    end

    attribute :font
    attribute :fill
    attribute :borders
    attribute :alignment
    attribute :number_format
    attribute :protection

    def initialize(opts={})
      self.font = opts[:font] || Font.new
      self.fill = opts[:fill] || Fill.new
      self.borders = opts[:borders] || Borders.new
      self.alignment = opts[:alignment] || Alignment.new
      self.number_format = opts[:number_format] || NumberFormat.new
      self.protection = opts[:protection] || Protection.new
    end
  end
end
