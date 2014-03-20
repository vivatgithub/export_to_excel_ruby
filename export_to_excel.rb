require 'cgi'

header = ["Number Type", "String Type", "Special Chars"]
contents = [
  ["1.1", "ONE", "<b>one</b>" ],
  ["2.1", "TWO", "<b>two</b>" ],
  ["3.1", "THREE", "<b>three</b>" ]
]

File.open("/tmp/out.xls", "w") {|f|
  f.write <<HEADER
<?xml version="1.0"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:x="urn:schemas-microsoft-com:office:excel"
  xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
  xmlns:html="http://www.w3.org/TR/REC-html40">
  <Worksheet ss:Name="Sheet1">
    <Table>
HEADER
  
  f.write "<Row>"
  header.each {|h|
    f.write %Q[\t<Cell><Data ss:Type="String">#{h}</Data></Cell>\n]
  }
  f.write "</Row>"
  
  contents.each {|row|
    f.write "<Row>"
    row.each {|c|
      c = "" if c.nil?
      c = c.to_s.strip

      if c.match(/^[0-9\.]+$/)
        #number type
        f.write %Q[\t<Cell><Data ss:Type="Number">#{c}</Data></Cell>\n]
      else
        #string type
        f.write %Q[\t<Cell><Data ss:Type="String">#{CGI.escapeHTML(c)}</Data></Cell>\n]
      end
      
    }
    f.write "</Row>"
  }
  
  f.write <<FOOTER
    </Table>
  </Worksheet>
</Workbook>
FOOTER
}