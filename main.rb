#!/usr/bin/env ruby
require 'glimmer-dsl-libui'
require_relative 'fill.rb'

include Glimmer

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
  suffix for the resulting files, otherwise a counter is used.
  
  The files are created in the same folder as the template, with file names
  [template name]-[suffix].
  }
  about_text = %{
  Source code available at:
  https://github.com/nmb/template-filler .
  }
  if(File.exist?(File.expand_path('version.rb', File.dirname(__FILE__))))
    require_relative 'version.rb'
    about_text = @version + about_text
  end
  
  # Add help menu
  menu('Help') {
    menu_item('Instructions') {
      on_clicked do
        msg_box('Instructions', help_text)
      end
    }
    
    about_menu_item {
      on_clicked do
        msg_box('About', about_text)
      end
    }
    
    quit_menu_item
  }
  
  window('Fill Template', 200, 100) {
    margined true
    
    vertical_box {
      horizontal_box {
        stretchy false
        
        button('Template') {
          stretchy false
          
          on_clicked do
            str = open_file
            @template_file = str.to_s unless str.nil?
          end
        }
        
        button('Data') {
          stretchy false
          
          on_clicked do
            str = open_file
            unless str.nil?
              @data_file = str.to_s
              @suffix = ""
              @inner.delete(0)
              @group.content {
                @inner = vertical_box {
                  @cbox = combobox {
                    headers = Roo::Spreadsheet.open(@data_file).sheet(0).parse(headers: true, clean: true).shift
                    items headers.keys
                  }
                }
              }
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
        }
        
        button('Run') {
          stretchy false
          
          on_clicked do
            if(File.exist?(@template_file) && File.exist?(@data_file))
              fill_all(@data_file, @template_file, @suffix)
              msg_box('Information', "Template: #{@template_file}.\nData: #{@data_file}.\nDone!")
            end
          end
        }
      }
      
      group('Filename suffix') {
        vertical_box {
          @cbox = combobox {
            stretchy false
          }
        }
      }
    }
  }.show
end

begin
  main unless defined?(Ocra)
rescue Exception => e
  f = Tempfile.new(['template-filler-', '.log'])
  f.puts e.inspect
  f.puts e.backtrace
  f.close
end
