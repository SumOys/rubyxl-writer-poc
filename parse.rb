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
  "FIRST_NAME": "Sarah",
  "LAST_NAME": "Falsely",
  "WEBSITE_1": "https://example.com",
  "WEBSITE_2": "https://example.org",
  "STREET_NO": "12345",
  "STREET": "Saturn Lane"
}

def render_cell(cell, replace)
  val = cell&.value.to_s
  links = []
  caps = val.scan(REPLACEMENT_KEY_REGEX).to_h
  unless caps.size == 0
    caps.each do |pattern, name|
      key = name.to_sym
      if replace[key].match?(/^https?:\/\//)
        links << %{HYPERLINK("#{replace[key]}", "Download Document")}
      else
        val.gsub!(pattern, replace[key])
      end
    end
  end

  if links.size == 1
    cell.change_contents(val, links.first)
    cell.change_font_color("0000FF")
  elsif links.size > 1
    cell.change_contents("Invalid: Multiple Attachments in a single cell")
  else
    cell.change_contents(val)
  end
      
end

workbook.worksheets.each { |worksheet|
  worksheet.each { |row|
    row&.cells&.each { |cell|
      render_cell(cell, replace)
    }
  }
}

workbook.write('/tmp/output.xlsx')