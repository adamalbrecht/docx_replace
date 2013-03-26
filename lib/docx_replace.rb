require "docx_replace/version"
require 'zip/zip'
require 'tempfile'

module DocxReplace
  class Doc
    def initialize(path, temp_dir=nil)
      @zip_file = Zip::ZipFile.new(path)
      @temp_dir = temp_dir
      read_docx_file
    end

    def replace(pattern, replacement)
      @document_content.sub!(pattern, replacement)
    end

    def commit(new_path=nil)
      write_back_to_file(new_path)
    end

    private
    DOCUMENT_FILE_PATH = 'word/document.xml'

    def read_docx_file
      @document_content = @zip_file.read(DOCUMENT_FILE_PATH)
    end

    def write_back_to_file(new_path=nil)
      if @temp_dir.nil?
        temp_file = Tempfile.new('docxedit-')
      else
        temp_file = Tempfile.new('docxedit-', @temp_dir)
      end
      Zip::ZipOutputStream.open(temp_file.path) do |zos|
        @zip_file.entries.each do |e|
          unless e.name == DOCUMENT_FILE_PATH
            zos.put_next_entry(e.name)
            zos.print e.get_input_stream.read
          end
        end

        zos.put_next_entry(DOCUMENT_FILE_PATH)
        zos.print @document_content
      end

      if new_path.nil?
        path = @zip_file.name
        FileUtils.rm(path)
      else
        path = new_path
      end
      FileUtils.mv(temp_file.path, path)
      @zip_file = Zip::ZipFile.new(path)
    end
  end
end
