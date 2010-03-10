require 'win32ole'

#�����ɗ^����xls�t�@�C���̃Z�����e�����ׂďo�͂���
def getAbsolutePath filename
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  return fso.GetAbsolutePathName(filename)
end
filename = getAbsolutePath("sample1.xls")

xl = WIN32OLE.new('Excel.Application')

book = xl.Workbooks.Open(filename)
#=begin
begin
  book.Worksheets.each do |sheet|	#�Ώۂ̃u�b�N�̂��ׂẴV�[�g�ɏ���
    sheet.UsedRange.Rows.each do |row|	#���[�N�V�[�g�̎g�p���Ă���͈͂���s���ƂɎ��o��
      record = []			#�o�͂��邽�߂Ɏg�p����z�񏉊���
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

