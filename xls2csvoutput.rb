require 'win32ole'
require 'find'


#引数に与えたxlsファイルのセル内容をすべて出力する
def getAbsolutePath filename
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  return fso.GetAbsolutePathName(filename)
end

# ファイルに書き出し
def xlsOutput(path, output_path)
  begin
    filename = getAbsolutePath(path)		#パス指定
    xl = WIN32OLE.new('Excel.Application')	#Excel起動
    book = xl.Workbooks.Open(filename)		#ExcelFileを開く
    
    book.Worksheets.each do |sheet|		#対象のブックのすべてのシートに処理
      output_file = File.open("#{output_path}#{path}_#{sheet.name}.csv","w")
      sheet.UsedRange.Rows.each do |row|	#ワークシートの使用している範囲を一行ごとに取り出す
        record = []				#出力するために使用する配列初期化
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
        output_file.write("#{record}\n")    # ファイルにデータ書き込み
      end
      output_file.close #ファイルを閉じる
    end
  ensure
    book.Close		#Bookを閉じる
    xl.Quit		#Excel終了
  end
end

# ディレクトリを処理する
def proc_directory(path, out_path)
  Find.find(path) do |file|
    if(File.file?(file) && File.extname(file) == '.xls' ) then
      xlsOutput(file, out_path)
    end
  end
end

# 使い方
def usage
  puts "usage:"
  puts "  ruby #{__FILE__} [filename]"
  puts "  ruby #{__FILE__} [dirname]"
end

if ARGV[0].empty?
  usage
  exit
end

target = ARGV[0]	#処理対象Excelファイル用引数指定
output_path = ARGV[1]	#保存先指定用引数指定

if(File.file?(target)) then # ファイルが渡された
  xlsOutput(target, output_path)
elsif(File.directory?(target)) then # ディレクトリが渡された
  #proc_directory(target, output_path)
else
  usage
  exit
end
