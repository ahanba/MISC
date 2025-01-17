#This script converts HTML tables to an XLS file.
#It also works for tables with merged columns and rows.

#coding: utf-8
require 'win32ole'
require 'nokogiri'
require 'csv'
require 'cgi'

LOG = File.open("log.txt", "w:utf-8")
#header indicator. Downcase.
HEADER_INDICATOR = 'key'.downcase

WIN32OLE.codepage = WIN32OLE::CP_UTF8

def getAbsolutePath(filename)
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  return fso.GetAbsolutePathName(filename)
end

def readHTML(htmlfile)
  begin
    csv = File.open("#{htmlfile}.csv", "w:utf-8")
    file = File.open(htmlfile)
    doc = Nokogiri::HTML(file)

    tables = doc.xpath('//div[@class="table-wrap"]/table')
    tables.each {|table|
      #p table.search('tbody').count
      table_count = table.search('tbody').count
      next unless table_count == 1 || table_count == 0
      next unless table.search(:th).inner_text.downcase.include?(HEADER_INDICATOR)

      table_rows = Array.new(table.search(:tr).length)
      table_rows.length.times {|i| table_rows[i] = Array.new(table.css('tr:nth-child(1) td,th').length) }

      rownum = 0
      table.search('tbody,thead').search(:tr).each do |tr|
        colnum = 0
        tr.search("th,td").each{|tag|
          #rowspan skip meged columns in a row
          while table_rows[rownum][colnum] != nil  do
            colnum += 1
          end
          next if colnum >= table_rows[rownum].length

          val = ""
          if tag.search("p").empty? #multi line
            val = CGI.unescapeHTML(tag.inner_text.gsub(/[[:space:]]/, ' ').strip)
          else #single line
            val =  CGI.unescapeHTML(tag.search("p,li,pre,a").map{ |paragraph|
              paragraph.inner_html.gsub("<br>", "\n").gsub(/<.*?>/, "").strip
            }.join("\n")).chomp
          end
          table_rows[rownum][colnum] = val

          #colspan case
          if tag[:colspan]
            (tag[:colspan].to_i - 1).times do |i|
              table_rows[rownum][colnum + i + 1] = val # i starts from 0, thus +1
            end
          end

          #rowspan case
          if tag[:rowspan]
            (tag[:rowspan].to_i - 1).times do |i|
              table_rows[rownum + i + 1][colnum] = val # i starts from 0, thus +1
              #rowspan and colspan case
              if tag[:colspan]
                (tag[:colspan].to_i - 1).times do |j|
                  table_rows[rownum + i + 1][colnum + j + 1] = val # i starts from 0, thus +1
                end
              end
            end
          end

          colnum += 1
        }
        rownum += 1
      end

      #make csv from table_rows
      table_rows.each {|row|
        csv.puts CSV.generate{ |csv|
          csv.add_row(row)
        }
      }
    }
  ensure
    file.close
    csv.close
    doc = nil
  end
end

def writeXLS(html)
  begin
    excel = WIN32OLE.new('Excel.Application')
    book = excel.Workbooks.Add()
    book.Sheets("Sheet1").Select
    book.Sheets("Sheet1").Name = "properties"
    book.Sheets.Add(After: book.ActiveSheet)
    book.Sheets("Sheet2").Select
    book.Sheets("Sheet2").Name = "plurals"

    sheet = book.Sheets("plurals")
    sheet.Activate
    sheet.Cells.Font.Name = "Calibri"
    sheet.Cells.Font.Size = 12
    sheet.Cells.Item(1, 2).value = 'ja'
    sheet.Cells.Item(2, 2).value = 'other'

    sheet = book.Sheets("properties")
    sheet.Activate
    sheet.Cells.Font.Name = "Calibri"
    sheet.Cells.Font.Size = 12
    sheet.Cells.Select
    excel.Selection.RowHeight = 15
    sheet.Columns("A:J").Select
    excel.Selection.NumberFormatLocal = "@"

    rnum = 1
    CSV.foreach("#{html}.csv", encoding:"UTF-8") { |row|
      cnum = 1
      next if rnum != 1 && row.map{|d| d.to_s.downcase}.include?(HEADER_INDICATOR)
      #row.values_at(*COLUMNs).each { |col|
      row.each { |col|
        case col
        when 'en', 'EN', 'En','eN'
          col = 'en'
        when 'ja', 'JA', 'Ja','jA'
          col = 'ja'
        when 'ko', 'KO', 'Ko','kO'
          col = 'ko'
        end
        sheet.Cells.Item(rnum, cnum).value = col
        cnum += 1
      }
      rnum += 1
    }

    sheet.Cells.Item(1, 1).value = ''

    #Save
    xlsfilename = html.sub('.html', '.xlsx')
    book.SaveAs(getAbsolutePath(xlsfilename))
  rescue
    LOG.puts "Error:\n#{$@}\n#{$!}"
  ensure
    excel.Workbooks.Close
    excel.Quit
  end
end

begin
  Dir.glob("**/*.html") { |html|
    p html
    readHTML(html)
    writeXLS(html)
  }
rescue
  LOG.puts "Error:\n#{$@}\n#{$!}"
ensure
  LOG.close
end
