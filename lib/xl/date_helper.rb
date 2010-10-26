require 'time'
require 'date'


class Date

  def to_w3cdtf
    strftime("%Y-%m-%d")
  end

  def to_date
    self
  end

  def to_datetime
    DateTime.new(year, month, day, 0, 0, 0, 0)
  end

  def to_time(form=:local)
    Time.send(form, year, month, day)
  end

end

class DateTime

  def to_w3cdtf
    to_time(:gm).to_w3cdtf
  end

  def to_date
    Date.new(year, month, day)
  end

  def to_datetime
    self
  end

  def to_time(form=:local)
    dest = form == :local ?
      new_offset(DateTime.now.offset-offset) :
      new_offset

    #Convert a fraction of a day to a number of microseconds
    usec = (dest.sec_fraction * 60 * 60 * 24 * (10**6)).to_i
    Time.send(form, dest.year, dest.month, dest.day, dest.hour, dest.min, dest.sec, usec)
  end

end


class Time

  def to_w3cdtf
    iso8601
  end

  def to_date
    Date.new(year, month, day)
  end

  def to_time
    self
  end

  def to_datetime
    # Convert seconds + microseconds into a fractional number of seconds
    seconds = sec + Rational(usec, 10**6)

    # Convert a UTC offset measured in minutes to one measured in a
    # fraction of a day.
    offset = Rational(utc_offset, 60 * 60 * 24)
    DateTime.new(year, month, day, hour, min, seconds, offset)
  end
end

# Excel date conversions. See http://www.cpearson.com/excel/datetime.htm
# if you'd like more information on the elegant wonder that is Excel's date
# system.
module Xl::DateHelper
  extend self

  BASE_DATE = Date.new(1899, 12, 31)
  BASE_DATE_1904 = Date.new(1904, 1, 1)
  NOT_A_LEAP_YEAR = Date.new(1900, 2, 28)

  def datelike?(d)
    d.kind_of?(DateTime) || d.kind_of?(Date) || d.kind_of?(Time)
  end

  # Convert a ruby Date, DateTime, or Time into an Excel serial date.
  #
  # @param [Date,DateTime,Time] date the datelike object to convert
  # @return [Float] a float representing the given date.
  def ruby_to_excel(date, use_1904=false)
    date = date.to_datetime
    base = base_date(use_1904)
    date_float = (date - base).to_f
    date_float += 1 if adjust_for_stupid_leap_year?(date, use_1904)
    date_float
  end

  # Convert a floating-point Excel serial date into a ruby DateTime.
  #
  # @param [Float] val
  #   A floating-point value representing a date in Excel. The integral part
  #   is the number of days since the base date (which is either Jan 1, 1900 or
  #   Jan 1, 1904 depending on a setting in the workbook) and the fractional
  #   part is the time as a fraction of a 24-hour day.
  # @param [Boolean] use_1904
  #   Whether to calculate dates starting in 1900 or in 1904.
  def excel_to_ruby(val=0.0, use_1904=false)
    val = val.to_f

    base = base_date(use_1904)
    date = base + val
    date -= 1 if adjust_for_stupid_leap_year?(date, use_1904)

    hour = (val % 1)  * 24
    min =  (hour % 1) * 60
    sec =  (min % 1)  * 60

    hour = hour.floor
    min = min.floor
    sec = sec.round

    DateTime.new(date.year, date.month, date.day, hour, min, sec)
  end

  private

    def base_date(use_1904=false)
      use_1904 ? BASE_DATE_1904 : BASE_DATE
    end

    def adjust_for_stupid_leap_year?(date, use_1904)
      (base_date(use_1904) < NOT_A_LEAP_YEAR) && (date > NOT_A_LEAP_YEAR)
    end
end
