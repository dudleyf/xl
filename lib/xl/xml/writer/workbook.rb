module Xl::Xml::Writer::Workbook
  include Xl::Coordinates

  def properties_core_document(props)
    XML::Document.new.tap do |doc|
      doc.root = make_node('cp:coreProperties', {
        'xmlns:cp' => Xl::Xml::NAMESPACES['cp'],
        'xmlns:dc' => Xl::Xml::NAMESPACES['dc'],
        'xmlns:dcterms' => Xl::Xml::NAMESPACES['dcterms'],
        'xmlns:dcmitype' => Xl::Xml::NAMESPACES['dcmitype'],
        'xmlns:xsi' => Xl::Xml::NAMESPACES['xsi']
      })
      make_subnode(doc.root, 'dc:creator') << props.creator
      make_subnode(doc.root, 'cp:lastModifiedBy') << props.last_modified_by
      make_subnode(doc.root, 'dcterms:created', 'xsi:type' => 'dcterms:W3CDTF') << props.created.to_w3cdtf
      make_subnode(doc.root, 'dcterms:modified', 'xsi:type' => 'dcterms:W3CDTF') << props.modified.to_w3cdtf
    end
  end

  def write_properties_core(props)
    properties_core_document(props).to_s
  end

  def content_types_document(workbook)
    XML::Document.new.tap do |doc|
      doc.root = make_node('Types', 'xmlns' => 'http://schemas.openxmlformats.org/package/2006/content-types')
      make_subnode(doc.root, 'Override', {
        'PartName' => '/' + Xl::Zip::ARC_THEME,
        'ContentType' => Xl::Xml::CONTENT_TYPES[:theme]
      })
      make_subnode(doc.root, 'Override', {
        'PartName' => '/' + Xl::Zip::ARC_STYLE,
        'ContentType' => Xl::Xml::CONTENT_TYPES[:styles]
      })
      make_subnode(doc.root, 'Default', {
        'Extension' => 'rels',
        'ContentType' => Xl::Xml::CONTENT_TYPES[:rels]
      })
      make_subnode(doc.root, 'Default', {
        'Extension' => 'xml',
        'ContentType' => 'application/xml'
      })
      make_subnode(doc.root, 'Override', {
        'PartName' => '/' + Xl::Zip::ARC_WORKBOOK,
        'ContentType' => Xl::Xml::CONTENT_TYPES[:workbook]
      })
      make_subnode(doc.root, 'Override', {
        'PartName' => '/' + Xl::Zip::ARC_APP,
        'ContentType' => Xl::Xml::CONTENT_TYPES[:extprops]
      })
      make_subnode(doc.root, 'Override', {
        'PartName' => '/' + Xl::Zip::ARC_CORE,
        'ContentType' => Xl::Xml::CONTENT_TYPES[:coreprops]
      })
      make_subnode(doc.root, 'Override', {
        'PartName' => '/' + Xl::Zip::ARC_SHARED_STRINGS,
        'ContentType' => Xl::Xml::CONTENT_TYPES[:strings]
      })

      workbook.worksheets.each_index do |sheet_id|
        make_subnode(doc.root, 'Override', {
          'PartName' => "/xl/worksheets/sheet#{sheet_id+1}.xml",
          'ContentType' => Xl::Xml::CONTENT_TYPES[:worksheet]
        })
      end
    end
  end

  def write_content_types(workbook)
    content_types_document(workbook).to_s
  end

  def properties_app_document(workbook)
    XML::Document.new.tap do |doc|
      worksheets_count = workbook.worksheets.length
      doc.root = make_node('Properties', {
        'xmlns' => Xl::Xml::NAMESPACES['ep'],
        'xmlns:vt' => Xl::Xml::NAMESPACES['vt']
      })
      make_subnode(doc.root, 'Application') << 'Microsoft Excel'
      make_subnode(doc.root, 'DocSecurity') << '0'
      make_subnode(doc.root, 'ScaleCrop') << 'false'
      make_subnode(doc.root, 'Company')

      make_subnode(doc.root, 'LinksUpToDate') << 'false'
      make_subnode(doc.root, 'SharedDoc') << 'false'
      make_subnode(doc.root, 'HyperlinksChanged') << 'false'
      make_subnode(doc.root, 'AppVersion') << '12.0000'

      heading_pairs = make_subnode(doc.root, 'HeadingPairs')
      vector = make_subnode(heading_pairs, 'vt:vector', :size => 2, :baseType => 'variant')

      variant = make_subnode(vector, 'vt:variant')
      make_subnode(variant, 'vt:lpstr') << 'Worksheets'

      variant = make_subnode(vector, 'vt:variant')
      make_subnode(variant, 'vt:i4') << worksheets_count

      title_of_parts = make_subnode(doc.root, 'TitlesOfParts')
      vector = make_subnode(title_of_parts, 'vt:vector', :size => worksheets_count, :baseType => 'lpstr')
      workbook.worksheets.each do |ws|
        make_subnode(vector, 'vt:lpstr') << ws.title
      end
    end
  end

  def write_properties_app(workbook)
    properties_app_document(workbook).to_s
  end

  def root_rels_document(workbook)
    XML::Document.new.tap do |doc|
      doc.root = make_node('Relationships', 'xmlns' => Xl::Xml::NAMESPACES['pr'])
      make_subnode(doc.root, 'Relationship', {'Id' => 'rId1', 'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',      'Target' => Xl::Zip::ARC_WORKBOOK})
      make_subnode(doc.root, 'Relationship', {'Id' => 'rId2', 'Type' => 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',   'Target' => Xl::Zip::ARC_CORE})
      make_subnode(doc.root, 'Relationship', {'Id' => 'rId3', 'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties', 'Target' => Xl::Zip::ARC_APP})
    end
  end

  def write_root_rels(workbook)
    root_rels_document(workbook).to_s
  end

  def workbook_document(workbook)
    XML::Document.new.tap do |doc|
      doc.root = make_node('workbook', {
        'xmlns' => Xl::Xml::NAMESPACES['ns'],
#        'xml:space' => 'preserve',
        'xmlns:r' => Xl::Xml::NAMESPACES['dr']
      })

      doc.root.space_preserve = true

      add_file_version(doc.root, workbook)
      add_workbook_pr(doc.root, workbook)
      add_book_views(doc.root, workbook)
      add_sheets(doc.root, workbook)
      add_named_ranges(doc.root, workbook)
      add_calc_pr(doc.root, workbook)
    end
  end

  def write_workbook(workbook)
    workbook_document(workbook).to_s
  end

  def workbook_rels_document(workbook)
    XML::Document.new.tap do |doc|
      doc.root = make_node('Relationships', 'xmlns' => Xl::Xml::NAMESPACES['pr'])

      workbook.worksheets.each_with_index do |sheet, i|
        make_subnode(doc.root, 'Relationship', {
          'Id' => "rId#{i+1}",
          'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
          'Target' => "worksheets/sheet#{i+1}.xml"
        })
      end

      rid = workbook.worksheets.length + 1

      make_subnode(doc.root, 'Relationship', {
        'Id' => "rId#{rid}",
        'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
        'Target' => 'sharedStrings.xml'
      })

      make_subnode(doc.root, 'Relationship', {
        'Id' => "rId#{rid+1}",
        'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
        'Target' => 'styles.xml'
      })

      make_subnode(doc.root, 'Relationship', {
        'Id' => "rId#{rid+2}",
        'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
        'Target' => 'theme/theme1.xml'
      })
    end
  end

  def write_workbook_rels(workbook)
    workbook_rels_document(workbook).to_s
  end

  private

  def add_file_version(root, workbook)
    make_subnode(root, 'fileVersion', {
      'appName' => 'xl',
      'lastEdited' => '4',
      'lowestEdited' => '4',
      'rupBuild' => '4505'
    })
  end

  def add_workbook_pr(root, workbook)
    make_subnode(root, 'workbookPr', {
      'defaultThemeVersion' => '124226',
      'codeName' => 'ThisWorkbook'
    })
  end

  def add_book_views(root, workbook)
    book_views = make_subnode(root, 'bookViews')
    make_subnode(book_views, 'workbookView', {
      'activeTab' => workbook.get_index(workbook.get_active_sheet),
      'autoFilterDateGrouping' => 1,
      'firstSheet' => 0,
      'minimized' => 0,
      'showHorizontalScroll' => 1,
      'showSheetTabs' => 1,
      'showVerticalScroll' => 1,
      'tabRatio' => 600,
      'visibility' => 'visible'
    })
  end

  def add_sheets(root, workbook)
    sheets = make_subnode(root, 'sheets')
    workbook.worksheets.each_with_index do |sheet, i|
      nd_sheet = make_subnode(sheets, 'sheet', {
        'name' => sheet.title,
        'sheetId' => i + 1,
        'r:id' => "rId#{i+1}"
      })

      if sheet.sheet_state != Xl::Worksheet::SHEETSTATE_VISIBLE
        nd_sheet['state'] = sheet.sheet_state
      end
    end
  end

  def add_named_ranges(root, workbook)
    defined_names = make_subnode(root, 'definedNames')
    workbook.get_named_ranges.each do |named_range|
      name = make_subnode(defined_names, 'definedName', 'name' => named_range.name)
      if named_range.local_only
        name['localSheetId'] = workbook.get_index(named_range.worksheet)
      end
      name << "'%s'!%s" % [
        named_range.worksheet.title.replace("'", "''"),
        absolute_coordinate(named_range.range)
      ]
    end
  end

  def add_calc_pr(root, workbook)
    make_subnode(root, 'calcPr', {
      'calcId' => '124519',
      'calcMode' => 'auto',
      'fullCalcOnLoad' => '1'
    })
  end

end
