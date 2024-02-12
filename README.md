# .ods-reader

An easy to use ruby library to convert an open document spreadsheet file to a hash with keys of each sheet name and stores the values as a 2D array in string format like a CSV. No Gems to install.

## Why

This library was created as a way to import a spreadsheet into SketchUp 2017 with the ruby API.

- Installing Gems in SketchUp requires extra intructions for the end user. If a Gem has a C binding it has to be compiled for the users os.
- To simplify this I created a simple class (inspired by the Rods gem) to parse the spreadsheet with compatibily with older ruby versions (tested with 2.2)
- Using my reworked version of the Zip Gem to prevent bleeding out of the namespace and potentially causing issues with other SketchUp extentions

#### For use on small files or when speed is not critical

This library uses "rexml" to read XML files which is extremely slow compared to other libraries like Nokogiri
A 75kb file takes about 5.2s to read and convert
A 9kb file takes about 0.4s to read and convert

## Usage

```ruby
require_relative 'path/to/ods-reader/ods-reader'

module ThisEnnsHere
  spreadsheet = ODStoCSVarray.new('/path/to/file.ods')

  spreadsheet.list # => ['Sheet1', 'Sheet2', ... ]
  spreadsheet.has?('Sheet1') # => true
  spreadsheet['Sheet1'] # => [['row1 col1','row1 col2', ... ], ['row2 col1','row2 col2', ... ], ... ]
end
```
