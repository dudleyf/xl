module Xl::Xml::Writer::Theme

  def write_theme
    theme_file = File.join(File.expand_path(File.dirname(__FILE__)), "theme1.xml")
    theme = Xl::Xml.read_xml(File.read(theme_file))
    theme.to_s
  end

end
