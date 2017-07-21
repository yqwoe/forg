class HomeController < ApplicationController
  def index
  end
  def upload
    begin
      excel=ExcelToCsv.new(params[:excel])
      respond_to do |format|
        format.csv {
          send_data excel.to_csv, filename:"#{params[:excel].original_filename}.csv"
        }
      end
    rescue => e
      puts e.backtrace
      message = e.message
      message = "你的文件有问题，重新另存为再试试～ : #{params[:excel].original_filename}" if message =~ /comparison of Fixnum with nil failed/
      respond_to do |format|
          format.csv{redirect_back(fallback_location: root_path,notice: "#{message}")}
      end
    end
  end
  
  def time_test(&blk)
    begin_time = Time.now.to_i
    yield(blk)
    puts "运行了 #{ (begin_time - Time.now.to_i) * 1000} ms"
  end
end
