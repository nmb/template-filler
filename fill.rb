require 'roo'
require 'sablon'

def fill(data, template, suffix, template_file)
  output_dir = File.dirname(template_file)
  base = File.basename(template_file, '.docx')
  template.render_to_file(File.join(output_dir, base + suffix + '.docx'), data)
end

def fill_all(datasheet, template_file)
  template = Sablon.template(File.expand_path(template_file))
  xlsx = Roo::Spreadsheet.open(datasheet)
    data_table = xlsx.sheet(0).parse(headers: true, clean: true)
    data_table.shift
  data_table.each_with_index do |d, i|
    fill(d.to_h, template, "-#{i+1}", template_file)
  end
end

