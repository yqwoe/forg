class BaseFile
	attr_reader :excel,:file_name
	def initialize(file)
		@file_name = file.original_filename
		begin
			case File.extname(file.path)
				when '.csv'
					csv_file=parse_csv(file)
					@excel = Roo::CSV.new(csv_file.path)
					@excel.default_sheet = @excel.sheets.first
				when '.xls'
					# parse_csv(file)
					@excel = Roo::Spreadsheet.open(file.path)
					# @excel = Roo::Excel.new(path)
					@excel.default_sheet = @excel.sheets.first
				when '.xlsx'
					@excel = Roo::Excelx.new(file.path)
					@excel.default_sheet = @excel.sheets.first
				else
					raise "类型有问题: #{@file_name}"
			end
		rescue Ole::Storage::FormatError => e
			csv_file = html_to_csv(file)
			@excel = Roo::CSV.new(csv_file.path)
			@excel.default_sheet = @excel.sheets.first
		rescue CSV::MalformedCSVError => e
			raise "你的文件已经被损坏了，请另存为再试试: #{@file_name}"
		end
	end
	#清洗excel
	def html_to_csv(file)
		doc = Nokogiri::HTML(file)
		aFile = File.new("#{Rails.root}/public/#{file.original_filename}-#{Time.now.to_i}.csv", "a+")
		doc.css("table tr").each do |d|
			row=[]
			row=d.css("td").map(&:text)
			row=d.css("th").map(&:text) if row.blank?
			next if row.length < 4
			aFile.syswrite("#{row.join(',')}\n") if aFile
		end
		aFile.close
		return aFile
	end
	
	#转换excel
	def parse_csv(file)
		f = File.open(file.path, "r:utf-16")
		aFile = File.new("#{Rails.root}/public/#{file.original_filename}-#{Time.now.to_i}.csv", "a+")
		i = 0
		begin
			f.each do |line|
				aFile.syswrite("#{line.split(/\s/).join(',')}\n") if aFile
			end
		rescue Encoding::InvalidByteSequenceError => e
			f = File.open(file.path, "r:gbk")
			i+=1
			retry if i<=5
		end
		aFile.close
		return aFile
	end
end