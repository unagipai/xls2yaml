require 'win32ole'
require 'find'


#�����ɗ^����xls�t�@�C���̃Z�����e�����ׂďo�͂���
def getAbsolutePath filename
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  return fso.GetAbsolutePathName(filename)
end

# �t�@�C���ɏ����o��
def xlsOutput(path)
  begin
    filename = getAbsolutePath(path)		#�p�X�w��
    xl = WIN32OLE.new('Excel.Application')	#Excel�N��
    book = xl.Workbooks.Open(filename)		#ExcelFile���J��
    output_file = File.open("#{path}.csv", "w")	#�t�@�C�����Ȃ���ΐV�K�쐬�A����΃N���A
    
    book.Worksheets.each do |sheet|		#�Ώۂ̃u�b�N�̂��ׂẴV�[�g�ɏ���
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
    end
  ensure
    book.Close		#Book�����
    xl.Quit		#Excel�I��
    output_file.close   #�t�@�C�������
  end
end

# �f�B���N�g������������
def proc_directory(path)
  Find.find(path) do |file|
    if(File.file?(file) && File.extname(file) == '.xls' ) then
      xlsOutput(file)
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
  xlsOutput(target)
elsif(File.directory?(target)) then # �f�B���N�g�����n���ꂽ
  proc_directory(target)
else
  usage
  exit
end
