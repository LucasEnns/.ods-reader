require 'rexml/document'
require_relative 'spreadsheet-utils/abstract-class'
require_relative 'zip/zipfilesystem'

module ThisEnnsHere
  class ODStoCSVarray < AbstractSpreadsheet
    Encoding.default_external = 'UTF-8'
    def initialize(
      file_location,
      max_columns: 300,
      remove_empty_rows: false,
      empty_values: [],
      empty_test_column: 10
    )
      # max_columns limits the width of the array to save time in parsing the XML
      # no need to return an array of 16000 empty cells ;)
      @max_columns = max_columns
      # remove_empty_rows allows for removing rows with no values
      # but the reference to the row will no longer follow the spreadsheet
      # not a significant time savings but the
      @remove_empty_rows = remove_empty_rows
      # empty values is an array to test if the row is empty
      # example ['', '', '', '', '', '', '', ''] - [''] = []
      # sometimes a template sheet has predefined values but mean nothing on their own
      @empty_values = [nil, '', '0', *empty_values]
      # set the point to test if the row is empty instead of parsing the whole row
      # usually there would be some valuable data in the first few columns
      @empty_test_column = empty_test_column
      open(file_location)
      convert
    end

    private

    def open(file_location)
      @sheets = {}

      Zip::ZipFile.open(file_location) do |zipfile|
        content = REXML::Document.new zipfile.file.read('content.xml')
        spreadSheet =
          content.elements[
            '/office:document-content/office:body/office:spreadsheet'
          ]
        spreadSheet
          .elements
          .each('table:table') do |table|
            title = table.attributes['table:name']
            # add error handling
            @sheets[title] = table
          end
      end
    end

    def convert
      @sheets_csv = {}
      @sheets.each_pair do |name, contents|
        csv = []
        contents
          .elements
          .each('table:table-row') do |row|
            values = row_to_array(row)
            csv << values if values
          end
        @sheets_csv[name] = csv
      end
    end

    def row_to_array(row)
      array = []
      verified = false
      row
        .elements
        .each('table:table-cell') do |cell|
          if !verified
            if array.length >= @empty_test_column &&
                 (array - @empty_values).empty?
              return @remove_empty_rows ? nil : Array.new(@max_columns, '')
            end
            verified = true if array.length >= 8
          end
          repetition = cell.attributes['table:number-columns-repeated'].to_i
          next array += Array.new(@max_columns, '') if repetition > @max_columns
          next array += Array.new(repetition, '') if repetition > 0

          text_value = cell.elements['text:p']
          next array << '' if !text_value
          array << text_value.text || ''
        end
      return nil if @remove_empty_rows && (array - @empty_values).empty?
      array[0...@max_columns]
    end
  end
end
