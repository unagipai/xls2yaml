require 'rubygems'
require 'spreadsheet'

require 'find'
require 'yaml'
require 'yaml_waml'
Spreadsheet.client_encoding = 'cp932'

#�t�@�C���ɏ����o��
def xls2yaml(path, output_path)
  begin
    if output_path == nil then
        makefile = "#{path}"
    else
    	makefile = "#{output_path}/#{path}"
    end

    book = Spreadsheet.open(path, 'rb')
    output_file = File.open("#{makefile}.yaml", "w")#�o�͗pyaml�t�@�C���̍쐬�ƓW�J
    
    book.worksheets.each do |sheet|		#�Ώۂ̃u�b�N�̂��ׂẴV�[�g�ɏ���
	record = []
	#sheet.map do |row|
	  #record << row.to_a.join()
	  #puts row.to_a.join().to_yaml
   	#end
	sheet.each do |row|
	  record_row = []
	  row.each do |cell|
	    record_row.push cell.to_s
	  end
	  record.push record_row
	  #puts record_row.to_yaml
	  #output_file.write("#{record_row.to_yaml}")
	end
	#puts record.to_yaml
        output_file.write("#{record.to_yaml}")    #�o�͗p�t�@�C���ɏ�������
    end
  ensure
    output_file.close 	#�o�͗p�t�@�C�������
  end
end

# �f�B���N�g������������
#def proc_directory(path, out_path)
#  Find.find(path) do |file|
#    if(File.file?(file) && File.extname(file) == '.xls' ) then
#      xls2yaml(file, out_path)
#    end
#  end
#end

# �g����
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

target = ARGV[0]		#�����Ώ�Excel�t�@�C���p�����w��

if(File.file?(target)) then 	#�t�@�C�����n���ꂽ
	output_path = []	#�ۑ���w��p�z��̏�����
	output_path = ARGV[1] 	#�ۑ���w��p�����w��
    if(File.file?(target) && File.extname(target) == '.xls' ) then
  	xls2yaml(target, output_path)
    else
	file_format_error	#�G���[���e�\��
    end
#elsif(File.directory?(target)) then #�f�B���N�g�����n���ꂽ
  #proc_directory(target, output_path)
else
  usage
  exit
end
