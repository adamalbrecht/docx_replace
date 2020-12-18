# encoding: UTF-8

require "docx_replace/version"
require 'zip'
require 'tempfile'

module DocxReplace
  class Doc
    attr_reader :document_content

    def initialize(path, temp_dir=nil)
      @zip_file = Zip::File.new(path)
      @temp_dir = temp_dir
      @document_file_path = get_document_file_path()
      read_docx_file
    end

    def replace(pattern, replacement, multiple_occurrences=false)
      replace = replacement.to_s.encode(xml: :text)
      if multiple_occurrences
        @document_content.force_encoding("UTF-8").gsub!(pattern, replace)
      else
        @document_content.force_encoding("UTF-8").sub!(pattern, replace)
      end
    end

    def matches(pattern)
      @document_content.scan(pattern).map{|match| match.first}
    end

    def unique_matches(pattern)
      matches(pattern)
    end

    alias_method :uniq_matches, :unique_matches


    def commit(new_path=nil)
      write_back_to_file(new_path)
    end

    private
    DOCUMENT_FILE_PATH = 'word/document.xml'
    def get_document_file_path
      content_types_entry = @zip_file.find_entry '[Content_Types].xml'
      raise 'Not an Open XML Document' unless content_types_entry

      content_types = content_types_entry.get_input_stream.read
      if m = content_types.match(/word\/document\d*\.xml/)
        @document_file_path = m[0]
      else
        raise 'Not a Word document'
      end
    end

    def read_docx_file
      @document_content = @zip_file.read(@document_file_path)
    end

    def write_back_to_file(new_path=nil)
      if @temp_dir.nil?
        temp_file = Tempfile.new('docxedit-')
      else
        temp_file = Tempfile.new('docxedit-', @temp_dir)
      end
      Zip::OutputStream.open(temp_file.path) do |zos|
        @zip_file.entries.each do |e|
          unless e.name == @document_file_path
            zos.put_next_entry(e.name)
            zos.print e.get_input_stream.read
          end
        end

        zos.put_next_entry(@document_file_path)
        zos.print @document_content
      end

      if new_path.nil?
        path = @zip_file.name
        FileUtils.rm(path)
      else
        path = new_path
      end
      FileUtils.mv(temp_file.path, path)
      @zip_file = Zip::File.new(path, true)
    end
  end
end
