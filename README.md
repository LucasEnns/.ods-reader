# .ods-reader
A ruby library to convert an open document spreadsheet file to a hash with keys of each sheet name and stores the values as a 2D array in string format like a CSV

## Why 
This library was created as a way to import a spreadsheet into SketchUp 2017 with the ruby API.
- Installing Gems in SketchUp requires extra intructions for the end user. If a Gem has a C binding it has to be compiled for the users os.
- To simplify this I created a simple class (inspired by the Rods gem) to parse the spreadsheet with compatibily with older ruby versions (tested with 2.2)
- Encapsluated the Zip Gem in the library namespace, and reworked the monkey patching to prevent bleeding out of the namespace and potentially causing issues with other SketchUp extentions

## Usage
Copy the repository

  ```ruby
  require_relative 'ods-reader'

  spreadsheet = ODStoCSVarray.new('/path/to/file.ods')

  spreadsheet.list # => ['Sheet1', 'Sheet2', ... ]
  spreadsheet['Sheet1'] # => [['row1 col1','row1 col1', ... ], ['row2 col1','row2 col1', ... ], ... ]
  ```
