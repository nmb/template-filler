# Template Filler

Simple utility for using data from an excel file to fill out word templates.

Select template and data files and press 'Go'. For each row in the data file,
each [MailMerge field](https://en.wikipedia.org/wiki/Mail_merge) in the
template is replaced with the corresponding data point, and saved as a new
file.

The first row of the spreadsheet should contain headers.

For information on how to add MailMerge fields in a word document, c.f. [Sablon
documentation](https://github.com/senny/sablon/blob/master/misc/TEMPLATE.md).

## Instructions
Install prerequisites:

```
bundle install
```
Run program:
```
ruby main.rb
```

## Build binary on windows
Install ocra:

```
gem install ocra

```

Create executable:

```
ocra --dll ruby_builtin_dlls/libssp-0.dll --dll ruby_builtin_dlls/libgmp-10.dll --dll ruby_builtin_dlls/libffi-7.dll --dll ruby_builtin_dlls\zlib1.dll --dll ruby_builtin_dlls\libssl-1_1-x64.dll --dll ruby_builtin_dlls/libcrypto-1_1-x64.dll --gem-full=openssl --gem-all=fiddle --gem-all=libui --chdir-first --add-all-core --output template-filler.exe --gem-all main.rb
```

