require 'xls'

describe XLS, "foreach method" do
  before(:each) do
    @fso = WIN32OLE.new('Scripting.FileSystemObject')
    @xls_file = @fso.GetAbsolutePathName('resource\xls.xls')
  end

  it "File not found => IOError" do
    lambda {XLS.foreach(@fso.GetAbsolutePathName('resource\nothing.xls'), 1)}.should raise_error(IOError)
    lambda {XLS.foreach(nil, 1)}.should raise_error(IOError)
    lambda {XLS.foreach(' ', 1)}.should raise_error(IOError)
  end

  it "Sheet not found => IOError" do
    lambda {XLS.foreach(@xls_file, 0)}.should raise_error(IOError)
    lambda {XLS.foreach(@xls_file, 'nothing')}.should raise_error(IOError)
  end
  
  it "empty sheet => nil" do
    result = true
    XLS.foreach(@xls_file, "empty") {|i| result = nil}
    result.should_not be_nil
  end
  
  it "(empty shet & no block) => nil" do
     XLS.foreach(@xls_file, "empty").should be_nil
  end

  it "(not empty sheet & no block) => LocalJumpError" do
    lambda {XLS.foreach(@xls_file, "norm")}.should raise_error(LocalJumpError)
  end

  it "Using the name or index sheet to specify use" do
    lambda {XLS.foreach(@xls_file, 1){|i|i}}.should_not raise_error
    lambda {XLS.foreach(@xls_file, "norm"){|i|i}}.should_not raise_error
  end
  
  it "Hash of the contents of the first line as a key 
      to the second line after UsedRange Excel range of reading lines" do
    input_file = [{"q"=>1, "w"=>2, "e"=>nil}, {"q"=>"a", "w"=>"b", "e"=>true}]
    
    result = []
    XLS.foreach(@xls_file, 1){ |record|
      result << record
    }
    
    input_file.size.times do |i|
      input_file[i]['q'].should == result[i]['q']
      input_file[i]['w'].should == result[i]['w']
      input_file[i]['e'].should == result[i]['e']
    end
  end

  it "If you have duplicate keys, came after a string of key priority" do
    input_file = [{"q"=>3, "w"=>2, "e"=>nil}, {"q"=>"c", "w"=>"b", "e"=>true}]
    
    result = []
    XLS.foreach(@xls_file, "dup"){ |record|
      result << record
    }
    
    input_file.size.times do |i|
      input_file[i]['q'].should == result[i]['q']
      input_file[i]['w'].should == result[i]['w']
      input_file[i]['e'].should == result[i]['e']
    end
  end

  it "Key spaces, if the column is empty since ignore" do
    input_file = [{"q"=>1, "w"=>2, "e"=>nil}, {"q"=>"a", "w"=>"b", "e"=>true}]
    
    result = []
    XLS.foreach(@xls_file, "space"){ |record|
      result << record
    }
    
    input_file.size.times do |i|
      input_file[i]['q'].should == result[i]['q']
      input_file[i]['w'].should == result[i]['w']
      input_file[i]['e'].should == result[i]['e']
    end
  end
  
end

