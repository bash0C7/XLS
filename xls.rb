# XLS -- class for parsing XLS data.
#        dependent on win32ole & Microsoft Excel
#
#        Programmed by Toshiaki <koshiba+rubyforge@4038nullpointer.com>
#        License : Ruby License

require 'win32ole'

class XLS
  FIRSTITEM = 1
  
  #  Parse XLS file & execute block
  # 
  # ARGS
  #    path,sheet and block
  # 
  #   path : XLS file fullpath
  #   sheet : Worksheet name or index(from 1)
  # 
  # RETURNS
  #   nil
  #   
  # Example
  #   c:\foo.xls
  #     worksheet(1)
  #       row(1) q,w,e
  #       row(2) 1,2,(empty)
  #       row(3) a,b,TRUE
  #  
  #  <Ruby Program>
  #  result = []
  #  XLS.foreach('c:\foo.xls', 1){ |record|
  #    result << record
  #  }
  #  result #=> [{"q"=>1, "w"=>2, "e"=>nil}, {"q"=>"a", "w"=>"b", "e"=>true}]
  def XLS.foreach(path, sheet, &block)
    open_reader(path, sheet, &block)
  end

  class << self
    private
    def open_reader(path, sheet)
      XLS::Reader.parse(path, sheet) {|record|  yield(record)}
      nil
    end
  end
  
  class Reader
    include Enumerable
    
    FIRSTITEM = 1

    def Reader.create(path, sheet)
      XLS::Reader.new(path, sheet)
    end
    
    def initialize(path, sheet)
      @app = WIN32OLE.new('Excel.Application')
      @book = @app.Workbooks.Open(path,{'ReadOnly' => true})
      @sheet = @book.sheets(sheet)
    rescue
      close
      raise IOError.new
    end

    def Reader.parse(path, sheet)
      reader = Reader.new(path, sheet)
      reader.each {|record| yield(record) }
    ensure
      reader.close if reader
      nil
    end

    def each
      keys = Hash.new
      @sheet.UsedRange.rows.each {|row|
        break if row.columns(FIRSTITEM).value.to_s.empty?
        if row.Row == FIRSTITEM
          keys = create_keys(row)
        else
          yield(create_hash(keys, row))
        end
      }
    end
    
    def close
      @book.close({'SaveChanges' => false}) if @book
      @app.Quit if @app
    end

    private
    def create_keys(row)
      keys = Hash.new
      row.columns.each { |cell|
        cell_value = cell.value.to_s
        break if cell_value.to_s.empty?
        keys.store(cell_value, cell.column)
      }
      keys
    end

    def create_hash(keys, row)
      record = Hash.new
      keys.each {|key, column|
        record.store(key, row.columns(column).value)
      }
      record
    end
  end
  
end
