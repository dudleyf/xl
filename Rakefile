$LOAD_PATH.unshift File.join(File.dirname(__FILE__), 'lib')

require 'rubygems'
require 'bundler/setup'

require 'rubygems/specification'
require 'rake/gempackagetask'
require 'rake/testtask'

spec = Gem::Specification.new do |s|
  s.name = 'xl'
  s.version = '0.0.1'
  s.platform = Gem::Platform::RUBY
  s.executables = Dir.glob("bin/**/*").map { |x| x.gsub('bin/', '') }
  s.files = %w(Rakefile) + Dir.glob("bin/**/*") + Dir.glob("lib/**/*")
  s.author = 'dudleyf'
  s.summary = 'A library for reading and writing Excel .xlsx files.'
  s.description = s.summary
  s.email = 'dudley@steambone.org'
  s.homepage = 'http://github.org/dudleyf/xl'
  s.rubyforge_project = ''
  
  s.add_dependency('libxml-ruby')
  s.add_dependency('rubyzip')
end

spec_file = "#{spec.name}.gemspec"
desc "Create #{spec_file}"
file spec_file => "Rakefile" do
  File.open(spec_file, "w") do |file|
    file.puts spec.to_ruby
  end
end

Rake::GemPackageTask.new(spec) do |pkg|
  pkg.gem_spec = spec
end
task :gem => spec_file

Rake::TestTask.new do |t|
  t.libs << "xl"
  t.test_files = FileList['test/*_test.rb']
  t.verbose = true
end
