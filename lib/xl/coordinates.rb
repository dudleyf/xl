module Xl::Coordinates
  COORDINATE_RE = /^\$?([A-Za-z]+)\$?(\d+)$/

  # Checks if a string looks like a coordinate reference.
  #
  # @param [#to_s] s the string to test
  # @return [MatchData, nil]
  def coordinate?(s)
    COORDINATE_RE.match(s.to_s)
  end

  # Convert a coordinate string like 'A3' to a pair ['A', 12].
  #
  # @param [String] s the string to convert
  # @return [List(String, Fixnum)]
  # @raise [CellCoordinatesException] if the string is an invalid cell reference
  def coordinate_from_string(s)
    matches = coordinate?(s)
    if matches
      column, row = matches.captures
      return [column, row.to_i]
    else
      raise Xl::CellCoordinatesError, "Invalid cell coordinates #{s}"
    end
  end

  # Convert a coordinate string to an absolute coordinate string
  #
  # @example
  #   absolute_coordinate("B12") #=> "$B$12"
  #
  # @param [String] s the coordinate string to convert
  # @return [String] the converted string
  def absolute_coordinate(s)
    '$%s$%d' % coordinate_from_string(s)
  end

  # Convert a column letter into a column number.
  #
  # @example
  #   column_index_from_string("B") #=> 2
  #
  # @param [String] s the column letter to convert
  # @return [Fixnum] the index of the column
  #
  # @todo Rename this and +get_column_letter+ to reflect that they're complements
  def column_index_from_string(s)
    s = s.upcase
    len = s.length
    case
      when len == 1
        (s[0].ord - 64)
      when len == 2
        ((1 + (s[0].ord - 65)) * 26) + (s[1].ord - 64)
      when len == 3
        ((1 + (s[0].ord - 65)) * 676) + ((1 + (s[1].ord - 65)) * 26) + (s[2].ord - 64)
      when len > 3
        raise Xl::ColumnStringIndexError, "Column string index can't be longer than 3 characters"
      else
        raise Xl::ColumnStringIndexError, "Column string index can't be empty"
    end
  end

  # Convert a column number into column letter.
  #
  # @example
  #   get_column_letter(2) #=> "B"
  #
  # @param [Fixnum] n the number of the column to convert
  # @return [String] the column letter
  #
  # @todo Rename this and +column_index_from_string+ to reflect that they're complements
  def get_column_letter(n)
    if n.kind_of?(String)
      n.upcase
    else
      col_name = ""
      quotient = n
      while n > 26
        quotient = n / 26
        rest = n % 26
        if rest > 0
          col_name = (64 + rest).chr + col_name
        else
          col_name = 'Z' + col_name
          quotient -= 1
        end
        n = quotient
      end
      col_name = (64 + quotient).chr + col_name
    end
  end

end
