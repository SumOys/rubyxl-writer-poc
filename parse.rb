#!/usr/bin/env ruby

require 'rubyXL'
require 'rubyXL/convenience_methods/cell'

require 'rubyXL/convenience_methods/font'
require 'rubyXL/convenience_methods/workbook'

# If a documented function doesn't work, you probably
# need one of these helpers. 
# require 'rubyXL/convenience_methods/worksheet'
# require 'rubyXL/convenience_methods/color'

# or you can load all of them
# require 'rubyXL/convenience_methods'

REPLACEMENT_KEY_REGEX = /({{([A-Z\d._]+)}})/

# RubyXL can take a file, but a byte-stream will be
# what it encounters in the Rail app
xls_file = File.open('simple_template.xlsx').read
workbook = RubyXL::Parser.parse_buffer(xls_file)
workbook.calc_pr.full_calc_on_load = true

s2 = workbook.add_worksheet('Sheet2')
s2.sheet_name = "Bob1"
s2.add_cell(0, 0, "Foo")

replace = {
  "NAME": "Sumit",
  "WEBSITE": "https://example1.com"
}

workbook.worksheets.each_with_index { |worksheet, sheet_num|
  worksheet.each { |row|
     row && row.cells.each { |cell|
       val = cell && cell.value
       cap = REPLACEMENT_KEY_REGEX.match(val)
       unless cap.nil?
        key = cap.captures.last.to_sym
        if replace[key] =~ /^https?:\/\// then
          link = %Q{HYPERLINK("#{replace[key]}", "Download Document")}
          cell.change_contents('', link)
          cell.change_font_color("0000FF")
        else
          cell.change_contents(replace[key])
        end
       end
     }
  }
}

workbook.write('/tmp/output.xlsx')