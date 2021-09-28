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
Install [ruby](https://www.ruby-lang.org/en/), and then prerequisites:

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
rake build
```

