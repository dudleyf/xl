source :gemcutter

# libxml has a bug (#28671 on rubyforge) that affects whitespace handling.
# It's fixed in my fork, but not in the real repo.
gem "libxml-ruby", :git => "git://github.com/dudleyf/libxml-ruby.git"
gem "rubyzip"

group :test do
  gem "redgreen"
  gem "diff-lcs"
end

group :development do
  gem "rake"
  gem "ruby-debug"
end
