#! ruby -Ks

require 'rubygems'
require 'spreadsheet'
require 'find'

# Excelクライアントの文字コードを指定
Spreadsheet.client_encoding = 'cp932'

# ワークシートをタブ区切りの文字列に変換
def csv(sheet)
  sheet.inject([]) do |result, item|
    result.push item.join("\t")
  end.join("\n")
end

# ファイルを処理する
def proc_file(path)
  book = Spreadsheet.open(path, 'rb')
  book.worksheets.each do |sheet|
    puts csv(sheet)
  end
end

# ディレクトリを処理する
def proc_directory(path)
  Find.find(path) do |file|
    if(File.file?(file) && File.extname(file) == '.xls' ) then
      proc_file(file)
    end
  end
end

# 使い方
def usage
  puts "usage:"
  puts "  ruby #{__FILE__} [filename]"
  puts "  ruby #{__FILE__} [dirname]"
end

if ARGV.empty?
  usage
  exit
end

target = ARGV[0]

if(File.file?(target)) then # ファイルが渡された
  proc_file(target) 
elsif(File.directory?(target)) then # ディレクトリが渡された
  proc_directory(target)
else
  usage
  exit
end