require 'roo'
require 'sablon'

def sanitize(filename)
  # Bad as defined by wikipedia: https://en.wikipedia.org/wiki/Filename#Reserved_characters_and_words
  # Also have to escape the backslash
  bad_chars = [ '/', '\\', '?', '%', '*', ':', '|', '"', '<', '>', '.', ' ' ]
  bad_chars.each do |bad_char|
    filename.gsub!(bad_char, '_')
  end
  filename
end

def fill(data, template, suffix, template_file)
  output_dir = File.dirname(template_file)
  base = File.basename(template_file, '.docx')
  template.render_to_file(File.join(output_dir, base + suffix + '.docx'), data)
end

def fill_all(datasheet, template_file, suffix_key = nil)
  template = Sablon.template(File.expand_path(template_file))
  xlsx = Roo::Spreadsheet.open(datasheet)
  data_table = xlsx.sheet(0).parse(headers: true, clean: true)
  data_table.shift
  data_table.each_with_index do |d, i|
    if(suffix_key && d.to_h.has_key?(suffix_key))
      sx = "-#{sanitize(d[suffix_key])}"
    else
      sx = "-#{i+1}"
    end
    fill(d.to_h, template, sx, template_file)
  end
end

