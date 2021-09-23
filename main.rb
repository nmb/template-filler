#!/usr/bin/env ruby
require 'libui'
require './fill.rb'

@template_file = ""
@data_file = ""

UI = LibUI
UI.init

main_window = UI.new_window('Fill Template', 200, 100, 1)
UI.window_set_margined(main_window, 1)
vbox = UI.new_vertical_box
UI.window_set_child(main_window, vbox)
hbox = UI.new_horizontal_box
UI.box_set_padded(vbox, 1)
UI.box_set_padded(hbox, 1)
UI.box_append(vbox, hbox, 1)

template_button = UI.new_button('Template')
data_button = UI.new_button('Data')
go_button = UI.new_button('Run')

UI.button_on_clicked(template_button) do
  str = UI.open_file(main_window)
  @template_file = str.to_s unless str.null?
end

UI.button_on_clicked(data_button) do
  str = UI.open_file(main_window)
  @data_file = str.to_s unless str.null?
end

UI.button_on_clicked(go_button) do
  if(File.exist?(@template_file) && File.exist?(@data_file))
    print "Running\n"
    fill_all(@data_file, @template_file)
    print "Done.\n"
  end
end

UI.window_on_closing(main_window) do
  puts 'Exit'
  UI.control_destroy(main_window)
  UI.quit
  0
end

#UI.window_set_child(main_window, template_button)
#UI.window_set_child(main_window, data_button)
UI.box_append(hbox, template_button, 0)
UI.box_append(hbox, data_button, 0)
UI.box_append(hbox, go_button, 0)
UI.control_show(main_window)

UI.main
UI.quit
