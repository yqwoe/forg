module Configable
	extend ActiveSupport::Concern
	included do
		def column_names
			ConvertConfig.headers.keys.uniq.compact
		end
		def column_values
			ConvertConfig.headers.values.flatten.uniq.compact
		end
		def headers
			ConvertConfig.headers
		end
		def filter_names
			ConvertConfig.filters
		end
		def filter_values
			ConvertConfig.filter_values
		end
	end
end