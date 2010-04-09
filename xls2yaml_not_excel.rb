require 'rubygems'
require 'spreadsheet'

require 'find'
require 'yaml'
require 'yaml_waml'
#Spreadsheet.client_encoding = 'UTF-8'
Spreadsheet.client_encoding = 'Windows-31J'
#Spreadsheet.client_encoding = 'cp932'

#ファイルに書き出し
def xls2yaml(path, output_path)
  begin
    if output_path == nil then
        makefile = "#{path}"
    else
    	makefile = "#{output_path}/#{path}"
    end

    book = Spreadsheet.open(path, 'rb')
    output_file = File.open("#{makefile}.yaml", "w")#出力用yamlファイルの作成と展開
    
    puts ("#{makefile}---->#{makefile}.yaml")
    
    book.worksheets.each do |sheet|		#対象のブックのすべてのシートに処理
	record = []
	puts sheet.name
#=begin
	sheet.each do |row|  #stop_fileはここでエラーになり停止する
	  record_row = []
	  row.each do |cell|
	    record_row.push cell.to_s
	  end
	  record.push record_row
	end
	#puts record.to_yaml
      output_file.write("#{record.to_yaml}")    #出力用ファイルに書き込み
#=end
    end
  ensure
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
  puts "  ruby #{__FILE__} [dirname]"
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
elsif(File.directory?(target)) then #ディレクトリが渡された
  proc_directory(target, output_path)
else
  usage
  exit
end
