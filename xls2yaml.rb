require 'win32ole'
require 'find'
require 'yaml'
require 'yaml_waml'


#引数に与えたxlsファイルのセル内容をすべて出力する
def getAbsolutePath filename
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  return fso.GetAbsolutePathName(filename)
end

#ファイルに書き出し
def xls2yaml(path, output_path)
  begin
    if output_path == nil then
        makefile = "#{path}"
    else
    	makefile = "#{output_path}/#{path}"
    end
    filename = getAbsolutePath(path)		#Excelで開くファイルのパス指定
    xl = WIN32OLE.new('Excel.Application')	#Excel起動
    book = xl.Workbooks.Open(filename)		#ExcelFileを開く

    output_file = File.open("#{makefile}.yaml", "w")#出力用yamlファイルの作成と展開
    
    book.Worksheets.each do |sheet|		#対象のブックのすべてのシートに処理
      puts sheet.Name
      sheet.Name.scan(/eet/){			#条件にあったシート名のみ処理
        sheet.UsedRange.Rows.each do |row|	#ワークシートの使用している範囲を一行ごとに取り出す
          record = []				#出力用配列の初期化
          row.Columns.each do |cell|
	    record << cell.Value
	  end
          output_file.write("#{record.to_yaml}\n")    #出力用ファイルに書き込み
        end
        output_file.write("\n")
      }
      end
  ensure
    book.Close		#Bookを閉じる
    xl.Quit		#Excel終了
    output_file.close 	#出力用ファイルを閉じる
  end
end

# ディレクトリを処理する
def proc_directory(path, out_path)
  Find.find(path) do |file|
    if(File.file?(file) && File.extname(file) == '.xls' ) then
      xls2yaml(file, out_path)
    end
  end
end

# 使い方
def usage
  puts "usage:"
  puts "  ruby #{__FILE__} [filename]"
  #puts "  ruby #{__FILE__} [dirname]"
end

def file_format_error
  puts " file format error "
  puts " The file format that can be used is 'xls' "
end

if ARGV[0].empty?
  usage
  exit
end

target = ARGV[0]		#処理対象Excelファイル用引数指定

if(File.file?(target)) then 	#ファイルが渡された
	output_path = []	#保存先指定用配列の初期化
	output_path = ARGV[1] 	#保存先指定用引数指定
    if(File.file?(target) && File.extname(target) == '.xls' ) then
  	xls2yaml(target, output_path)
    else
	file_format_error	#エラー内容表示
    end
#elsif(File.directory?(target)) then #ディレクトリが渡された
  #proc_directory(target, output_path)
else
  usage
  exit
end
