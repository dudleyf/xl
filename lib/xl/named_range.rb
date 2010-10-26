class Xl::NamedRange

  attr_accessor :name, :worksheet, :range, :local_only

  def initialize(name, worksheet, range)
    @name = name
    @worksheet = worksheet
    @range = range
    @local_only = false
  end

  def to_s
    "#{@worksheet.title}!#{@range}"
  end

  class << self
    def split(str)
      matches = /'?([^']*)'?!\$([A-Za-z]+)\$([0-9]+)/.match(str)

      raise Xl::XlError, 'Invalid named range string' if matches.nil?

      sheet_name, column, row = matches.captures
      [sheet_name, column, row.to_i]
    end
  end
end
