# -*- encoding: utf-8 -*-

Gem::Specification.new do |s|
  s.name = %q{xl}
  s.version = "0.0.1"

  s.required_rubygems_version = Gem::Requirement.new(">= 0") if s.respond_to? :required_rubygems_version=
  s.authors = ["dudleyf"]
  s.date = %q{2010-10-26}
  s.description = %q{A library for reading and writing Excel .xlsx files.}
  s.email = %q{dudley@steambone.org}
  s.files = ["Rakefile", "lib/xl", "lib/xl/cell.rb", "lib/xl/coordinates.rb", "lib/xl/date_helper.rb", "lib/xl/named_range.rb", "lib/xl/style.rb", "lib/xl/workbook.rb", "lib/xl/worksheet.rb", "lib/xl/xml", "lib/xl/xml/reader", "lib/xl/xml/reader/strings.rb", "lib/xl/xml/reader/styles.rb", "lib/xl/xml/reader/workbook.rb", "lib/xl/xml/reader/worksheet.rb", "lib/xl/xml/reader.rb", "lib/xl/xml/writer", "lib/xl/xml/writer/string_table.rb", "lib/xl/xml/writer/style_table.rb", "lib/xl/xml/writer/theme.rb", "lib/xl/xml/writer/theme1.xml", "lib/xl/xml/writer/workbook.rb", "lib/xl/xml/writer/worksheet.rb", "lib/xl/xml/writer.rb", "lib/xl/xml.rb", "lib/xl/zip.rb", "lib/xl.rb"]
  s.homepage = %q{http://github.org/dudleyf/xl}
  s.require_paths = ["lib"]
  s.rubyforge_project = %q{}
  s.rubygems_version = %q{1.3.7}
  s.summary = %q{A library for reading and writing Excel .xlsx files.}

  if s.respond_to? :specification_version then
    current_version = Gem::Specification::CURRENT_SPECIFICATION_VERSION
    s.specification_version = 3

    if Gem::Version.new(Gem::VERSION) >= Gem::Version.new('1.2.0') then
      s.add_runtime_dependency(%q<libxml-ruby>, [">= 0"])
      s.add_runtime_dependency(%q<rubyzip>, [">= 0"])
    else
      s.add_dependency(%q<libxml-ruby>, [">= 0"])
      s.add_dependency(%q<rubyzip>, [">= 0"])
    end
  else
    s.add_dependency(%q<libxml-ruby>, [">= 0"])
    s.add_dependency(%q<rubyzip>, [">= 0"])
  end
end
