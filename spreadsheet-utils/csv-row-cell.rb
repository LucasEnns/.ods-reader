module ThisEnnsHere
  class CSV < Array
    ROW_END = /\R/i
    COLUMN_END = /\,/i
    attr_accessor :value

    def initialize(array, first_row = 0)
      @value = array[first_row..-1].map { |row| Row.new(row) }
      super(@value)
    end

    def +(other)
      @value + other.value
    end

    def [](index)
      if index.kind_of? Range
        r_begin = index.begin
        r_end = index.end == -1 ? @value.length - 1 : index.end
        return CSV.new((r_begin..r_end).map { |v| @value[v] })
      end
      @value[index]
    end

    def []=(index, row)
      @value[index] = Row.new(row)
      super(index, @value[index])
    end

    def self.get(url, redirects = 1)
      loop do
        uri = URI.parse(url)
        result = Net::HTTP.get_response(uri)
        return self.deserialize(result.body) if result.is_a? Net::HTTPOK
        raise FileOpeningError.new('Too many redirects') if redirects > 10
        redirects += 1
        url = result['location']
      end
    rescue StandardError => error
      SU.warning('Problème de réseau')
      raise ThisEnnsHere::Errors::ExternalContentError.new error
    end

    def self.read(file_path)
      self.deserialize(self.open(file_path))
    end

    def self.write(file_path, array, mode: 'a')
      File.write(file_path, self.serialize(array), mode: mode)
    end

    def self.open(file_path)
      IO.binread(file_path)
    rescue StandardError
      SU.warning("Fermer le fichier #{File.basename(file_location)}")
      raise ThisEnnsHere::Errors::FileOpeningError.new error
    end

    def self.deserialize(csv, first_row = 0)
      CSV.new(
        csv
          .force_encoding('UTF-8')
          .split(ROW_END)
          .map { |column| "#{column}, ".split(COLUMN_END) },
        first_row,
      )
    end

    def self.serialize(array)
      array.map { |v| v.join(',') }.join("\n")
    end
  end

  class Row < Array
    COLUMNS = [*'A'..'IW'].map.with_index { |v, i| [v, i] }.to_h
    COLUMNS.update((0..256).map.with_index { |v, i| [v.to_s, i] }.to_h)

    def initialize(array)
      @value = array.map { |cell| Cell.new(cell) }
      super(@value)
    end

    def copy
      Row.new(@value)
    end

    def filter(*args)
      Row.new(@value.select { |v| v.send(*args) })
    end

    def any?
      filter(:to_b).to_a.any?
    end

    def with_value
      filter(:to_value)
    end

    def convert_values(method)
      @value.map { |v| v.send(method) }
    end

    def to_in
      convert_values(:to_in)
    end

    def join(separator)
      Array.new(@value).join(separator)
    end

    def []=(index, value)
      column_index = COLUMNS[index.to_s.upcase]
      @value[column_index] = Cell.new(value)
      super(column_index, @value[column_index])
    end

    def [](index)
      if index.kind_of? Range
        first = COLUMNS[index.first.to_s.upcase]
        last = COLUMNS[index.last.to_s.upcase]
        return Row.new(@value[first..last]) if first && last
      end
      return @value[COLUMNS[index.to_s.upcase]] if COLUMNS[index.to_s.upcase]

      raise "#{index} has to be an integer (0..256) or letter reference ('A'..'IW') to a column location"
    end
  end

  class Cell < String
    # include ThisEnnsHere::Types
    def initialize(input)
      @value = input.to_s.strip
      super(@value)
    end

    def upcase
      Cell.new(String.new(@value).upcase)
    end

    def +(string)
      Cell.new(String.new(@value) + string)
    end

    # def to_in
    #   Inches.new(@value)
    # end
    # alias_method :to_d, :to_in

    # def to_mm
    #   Inches.new(@value, 'mm')
    # end
    # alias_method :mm, :to_mm

    # def to_ft
    #   @value.to_f
    # end

    # def to_p
    #   Price.new(@value)
    # end

    # def to_q
    #   Quantity.new(@value)
    # end

    # def to_ratio
    #   @value.to_f / 100
    # end

    # def to_minutes
    #   Minutes.new(@value)
    # end

    # def to_hours
    #   Hours.new(@value)
    # end

    # def to_s
    #   self.strip.gsub(/\-?$/, '')
    # end

    def to_b
      ![
        '',
        '0',
        '0.0',
        'false',
        'FALSE',
        'faux',
        'FAUX',
        'non',
        'no',
        'NO',
        'non',
        'NON',
      ].include? @value
    end

    def to_value
      ![
        '',
        'false',
        'FALSE',
        'faux',
        'FAUX',
        'non',
        'no',
        'NO',
        'non',
        'NON',
      ].include? @value
    end

    def !
      !to_b
    end
  end
end
