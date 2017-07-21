class ExcelToCsv < BaseFile
	include Configable
	attr_reader :row_index,:header_array, :header_indexes, :header_hash,:other_indexes
	
	def initialize(file)
		super(file)
		@row_index = 0
		@header_indexes = []
		@other_indexes = []
		@header_hash = {}
	end
	
	def to_csv
		CSV.generate do |csv|
			csv << column_names #这是文件的headers
			@excel.each_with_index do |row, index|
				if @header_indexes.present?
					row_array=[]
					@header_indexes.each do |header_index|
						# puts header_index
						# puts "#{row}"
						# puts "#{@header_indexes}"
						row_array << if header_index
							              row[header_index].to_s.gsub(/,|"/,"")
							           else
													 nil
						             end
					end
					if row_filter?(row)
						if @file_name.include?("个人")
							row_array << "是"
						else
							row_array << "否"
						end
						csv << row_array
					end
				else
					if valid?(row)
						@row_index = index
						@header_array = row
						row.each_with_index do |col, i|
							@other_indexes << i
							next if col.blank?
							column_name=search_column(col)
							next if column_name.blank?
							@header_indexes[column_names.index(column_name)] = i
						end
					end
				end
			end
		end
	end
	def row_filter?(row)
		puts "#{row}"
		flag = false
		@header_indexes.each do |other_index|
			next if other_index.blank?
			header = @header_array[other_index]
			value = row[other_index].to_s
			price = value.gsub(/,/,"") #判断是否存在金额
			if header.present? && filter_names.include?(header) && (price =~ /^-?(0|[1-9][0-9]*)(\.[0-9]*)?$/ &&  price.to_i > 0)
				flag = true
				break
			end
		end
		return flag
	end
	def search_column(column)
		puts @header_hash
		puts column
		return @header_hash[column] if @header_hash.key?(column)
		headers.each do |k, v|
			if v && v.include?(column)
				@header_hash[column]=k
				break
			end
		end
		return @header_hash[column]
	end
	
	def valid?(row)
		size = 0
		column_values.each do |column|
			size += 1 if row.compact.join("").include?(column)
		end
		puts "#{column_values} \n ====> \n #{row.compact} ==> #{size}" if size > 4
		return size > 4
	end
end