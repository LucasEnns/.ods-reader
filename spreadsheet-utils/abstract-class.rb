module ThisEnnsHere
  class AbstractSpreadsheet
    def list
      @sheets_csv.keys
    end

    def has?(name)
      @sheets_csv.keys.include? name
    end

    def [](name, excel_reference = nil)
      name = list[name] if name.is_a? Integer
      raise SheetNonExistantError unless has?(name)

      if excel_reference.is_a? String
        if excel_reference.match(/\d/)
          row = row_to_index(excel_reference)
          raise OutOfRangeError if @sheets_csv[name].length < row
          return @sheets_csv[name][row] unless excel_reference.match(/[A-Z]/i)
        end

        column = column_to_index(excel_reference)
        raise OutOfRangeError if @sheets_csv[name][row || 0].length < column
        return get_column(name, column) unless row

        return @sheets_csv[name][row][column]
      end
      return @sheets_csv[name]
    end

    private

    def get_column(sheet, column)
      @sheets_csv[sheet].map { |v| v[column] }
    end

    def column_to_index(column)
      upper = column.upcase.gsub(/\d+/, '')
      if upper.empty?
        raise MissingReferenceError.new('Need a alpha letter to access column')
      end
      index = upper[-1].ord - 65
      index += (upper[-2].ord - 64) * 26 if upper[-2]
      index += (upper[-3].ord - 64) * 26**2 if upper[-3]
      index
    end

    def row_to_index(row)
      unless row.match(/\d/)
        raise MissingReferenceError.new('Need a number to access row')
      end
      row.gsub(/[A-Z]+/i, '').to_i - 1
    end

    class OutOfRangeError < StandardError
      def initialize(
        msg = 'The references or index are out of range for the array'
      )
      end
    end

    class SheetNonExistantError < StandardError
      def initialize(msg = "The name for this sheet doesn't exist")
      end
    end

    class MissingReferenceError < StandardError
    end

    class FileTypeError < StandardError
    end
  end
end
