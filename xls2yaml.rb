require 'win32ole'
require 'find'
require 'yaml'
require 'yaml_waml'


#�����ɗ^����xls�t�@�C���̃Z�����e�����ׂďo�͂���
def getAbsolutePath filename
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  return fso.GetAbsolutePathName(filename)
end

#�t�@�C���ɏ����o��
def xls2yaml(path, output_path)
  begin
    if output_path == nil then
        makefile = "#{path}"
    else
    	makefile = "#{output_path}/#{path}"
    end
    filename = getAbsolutePath(path)		#Excel�ŊJ���t�@�C���̃p�X�w��
    xl = WIN32OLE.new('Excel.Application')	#Excel�N��
    book = xl.Workbooks.Open(filename)		#ExcelFile���J��

    output_file = File.open("#{makefile}.yaml", "w")#�o�͗pyaml�t�@�C���̍쐬�ƓW�J
    
    book.Worksheets.each do |sheet|		#�Ώۂ̃u�b�N�̂��ׂẴV�[�g�ɏ���
      puts sheet.Name
      sheet.Name.scan(/eet/){			#�����ɂ������V�[�g���̂ݏ���
        sheet.UsedRange.Rows.each do |row|	#���[�N�V�[�g�̎g�p���Ă���͈͂���s���ƂɎ��o��
          record = []				#�o�͗p�z��̏�����
          row.Columns.each do |cell|
	    record << cell.Value
	  end
          output_file.write("#{record.to_yaml}\n")    #�o�͗p�t�@�C���ɏ�������
        end
        output_file.write("\n")
      }
      end
  ensure
    book.Close		#Book�����
    xl.Quit		#Excel�I��
    output_file.close 	#�o�͗p�t�@�C�������
  end
end

# �f�B���N�g������������
def proc_directory(path, out_path)
  Find.find(path) do |file|
    if(File.file?(file) && File.extname(file) == '.xls' ) then
      xls2yaml(file, out_path)
    end
  end
end

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
