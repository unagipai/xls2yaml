#! ruby -Ks

require 'rubygems'
require 'spreadsheet'
require 'find'

# Excel�N���C�A���g�̕����R�[�h���w��
Spreadsheet.client_encoding = 'cp932'

# ���[�N�V�[�g���^�u��؂�̕�����ɕϊ�
def csv(sheet)
  sheet.inject([]) do |result, item|
    result.push item.join("\t")
  end.join("\n")
end

# �t�@�C������������
def proc_file(path)
  book = Spreadsheet.open(path, 'rb')
  book.worksheets.each do |sheet|
    puts csv(sheet)
  end
end

# �f�B���N�g������������
def proc_directory(path)
  Find.find(path) do |file|
    if(File.file?(file) && File.extname(file) == '.xls' ) then
      proc_file(file)
    end
  end
end

# �g����
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

if(File.file?(target)) then # �t�@�C�����n���ꂽ
  proc_file(target) 
elsif(File.directory?(target)) then # �f�B���N�g�����n���ꂽ
  proc_directory(target)
else
  usage
  exit
end
