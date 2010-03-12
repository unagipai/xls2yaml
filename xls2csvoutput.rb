require 'win32ole'
require 'find'


#�����ɗ^����xls�t�@�C���̃Z�����e�����ׂďo�͂���
def getAbsolutePath filename
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  return fso.GetAbsolutePathName(filename)
end

# �t�@�C���ɏ����o��
def xlsOutput(path, output_path)
  begin
    filename = getAbsolutePath(path)		#�p�X�w��
    xl = WIN32OLE.new('Excel.Application')	#Excel�N��
    book = xl.Workbooks.Open(filename)		#ExcelFile���J��
    
    book.Worksheets.each do |sheet|		#�Ώۂ̃u�b�N�̂��ׂẴV�[�g�ɏ���
      output_file = File.open("#{output_path}#{path}_#{sheet.name}.csv","w")
      sheet.UsedRange.Rows.each do |row|	#���[�N�V�[�g�̎g�p���Ă���͈͂���s���ƂɎ��o��
        record = []				#�o�͂��邽�߂Ɏg�p����z�񏉊���
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
        output_file.write("#{record}\n")    # �t�@�C���Ƀf�[�^��������
      end
      output_file.close #�t�@�C�������
    end
  ensure
    book.Close		#Book�����
    xl.Quit		#Excel�I��
  end
end

# �f�B���N�g������������
def proc_directory(path, out_path)
  Find.find(path) do |file|
    if(File.file?(file) && File.extname(file) == '.xls' ) then
      xlsOutput(file, out_path)
    end
  end
end

# �g����
def usage
  puts "usage:"
  puts "  ruby #{__FILE__} [filename]"
  puts "  ruby #{__FILE__} [dirname]"
end

if ARGV[0].empty?
  usage
  exit
end

target = ARGV[0]	#�����Ώ�Excel�t�@�C���p�����w��
output_path = ARGV[1]	#�ۑ���w��p�����w��

if(File.file?(target)) then # �t�@�C�����n���ꂽ
  xlsOutput(target, output_path)
elsif(File.directory?(target)) then # �f�B���N�g�����n���ꂽ
  #proc_directory(target, output_path)
else
  usage
  exit
end
