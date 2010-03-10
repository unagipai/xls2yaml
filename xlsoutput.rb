require 'win32ole'

#引数に与えたxlsファイルのセル内容をすべて出力する
def getAbsolutePath filename
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  return fso.GetAbsolutePathName(filename)
end
filename = getAbsolutePath("sample1.xls")

xl = WIN32OLE.new('Excel.Application')

book = xl.Workbooks.Open(filename)
#=begin
begin
  book.Worksheets.each do |sheet|	#対象のブックのすべてのシートに処理
    sheet.UsedRange.Rows.each do |row|	#ワークシートの使用している範囲を一行ごとに取り出す
      record = []			#出力するために使用する配列初期化
      row.Columns.each do |cell|
	if cell.Value.is_a?(String) &&
 	   cell.Value =~ %r(\d\d\d\d/\d\d/\d\d \d\d:\d\d:\d\d)
	  begin
	    record << Time.mktime(*cell.Value.split(%r([:/])))
	  rescue ArgumentError => e
	    STDERR.puts e.inspect
	  end
	else
      	record << cell.Value
	end
      end
      puts record.join(",")
    end
  end
ensure
  book.Close
  xl.Quit
end
#=end
#puts "3"
#puts filename

