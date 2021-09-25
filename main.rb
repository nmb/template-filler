#!/usr/bin/env ruby
require 'libui'
require './fill.rb'

UI = LibUI

def main
@template_file = ""
@data_file = ""
@suffixes = []
@suffix = ""
@cbox

help_text = %{
This program takes a template (word file) and a data source (excel file), and
creates a new word file for each data row.

The template must use MailMerge fields as markers. The data file should contain
column headers on the first row, corresponding to the MailMerge field names.

After a data file has been loaded, it is possible to select a header to use as
suffix for the resulting files.  

The files are created in the same folder as the template, with file names
[template name]-[suffix].
}
about_text = %{
Source code available at https://github.com/nmb/template-filler .
}

UI.init

# Add help menu
menu = UI.new_menu('Help')
help_menu_item = UI.menu_append_item(menu, 'Instructions')
UI.menu_item_on_clicked(help_menu_item) do
    UI.msg_box(@main_window, 'Instructions', help_text)
end
about_menu_item = UI.menu_append_item(menu, 'About')
UI.menu_item_on_clicked(about_menu_item) do
    UI.msg_box(@main_window, 'About', about_text)
end
UI.menu_append_quit_item(menu)
UI.on_should_quit(proc {
  UI.control_destroy(@main_window)
  UI.quit
  0
})

# Set up main window
@main_window = UI.new_window('Fill Template', 200, 100, 1)
UI.window_set_margined(@main_window, 1)
vbox = UI.new_vertical_box
UI.window_set_child(@main_window, vbox)
hbox = UI.new_horizontal_box
UI.box_set_padded(vbox, 1)
UI.box_set_padded(hbox, 1)
UI.box_append(vbox, hbox, 1)

template_button = UI.new_button('Template')
data_button = UI.new_button('Data')
go_button = UI.new_button('Run')

group = UI.new_group('Filename suffix')
UI.group_set_margined(group, 1)
UI.box_append(vbox, group, 0)
inner = UI.new_vertical_box
UI.box_set_padded(inner, 1)
UI.group_set_child(group, inner)

UI.button_on_clicked(template_button) do
  str = UI.open_file(@main_window)
  @template_file = str.to_s unless str.null?

end

UI.button_on_clicked(data_button) do
  str = UI.open_file(@main_window)
  unless str.null?
    @data_file = str.to_s
    @suffix = ""
    UI.box_delete(inner, 0)
    inner = UI.new_vertical_box
    UI.box_set_padded(inner, 1)
    UI.group_set_child(group, inner)
    @cbox = UI.new_combobox
    headers = Roo::Spreadsheet.open(@data_file).sheet(0).parse(headers: true, clean: true).shift
    @suffixes = headers.keys
    @suffixes.each do |h|
      UI.combobox_append(@cbox, h)
    end
    UI.box_append(inner, @cbox, 0)
    UI.combobox_on_selected(@cbox) do |ptr|
      no = UI.combobox_selected(ptr)
      @suffix = @suffixes[no]
end

  end
end

UI.button_on_clicked(go_button) do
  if(File.exist?(@template_file) && File.exist?(@data_file))
    fill_all(@data_file, @template_file, @suffix)
    UI.msg_box(@main_window, 'Information', "Template: #{@template_file}.\nData: #{@data_file}.\nDone!")
  end
end

UI.window_on_closing(@main_window) do
  UI.control_destroy(@main_window)
  UI.quit
  0
end

@cbox = UI.new_combobox
UI.box_append(inner, @cbox, 0)
UI.box_append(hbox, template_button, 0)
UI.box_append(hbox, data_button, 0)
UI.box_append(hbox, go_button, 0)
UI.control_show(@main_window)

UI.main
UI.quit

end

begin
  main unless defined?(Ocra)
rescue Exception => e
    f = Tempfile.new(['template-filler-', '.log']) 
    f.puts e.inspect
    f.puts e.backtrace
    f.close
end

