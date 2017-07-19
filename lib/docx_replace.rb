# encoding: UTF-8

require "docx_replace/version"
require 'zip'
require 'tempfile'

module DocxReplace
  class Doc
    DOCUMENT_FILE_PATH = 'word/document.xml'

    attr_accessor :document_contents

    def initialize(path, temp_dir=nil)
      @zip_file = Zip::File.new(path)
      @temp_dir = temp_dir
      read_docx_file
    end

    def replace(pattern, replacement, multiple_occurrences=false)
      replace = replacement.to_s.encode(xml: :text)
      if multiple_occurrences
        @document_contents.keys.each do |name|
          @document_contents[name].force_encoding("UTF-8").gsub!(pattern, replace)
        end
      else
        @document_contents.keys.each do |name|
          @document_contents[name].force_encoding("UTF-8").sub!(pattern, replace)
        end
      end
    end

    def matches(pattern, unique = false)
      if unique
        @document_contents[DOCUMENT_FILE_PATH].scan(pattern).map{|match| match.first}
      else
        @document_contents[DOCUMENT_FILE_PATH].scan(pattern).flatten.inject(Hash.new(0)) { |hash, key| hash[key] += 1; hash }
      end
    end

    def unique_matches(pattern)
      matches(pattern, true)
    end

    alias_method :uniq_matches, :unique_matches


    def commit(new_path=nil)
      write_back_to_file(new_path)
    end

    private

    def read_docx_file
      @document_contents = {}

      @zip_file.entries.each do |entry|
        if entry.name.match? /^word\/(document|header[0-9]+|footer[0-9]+).xml/
          @document_contents[entry.name] = @zip_file.read(entry.name)
        end
      end

    end

    def write_back_to_file(new_path=nil)
      if @temp_dir.nil?
        temp_file = Tempfile.new('docxedit-')
      else
        temp_file = Tempfile.new('docxedit-', @temp_dir)
      end
      Zip::OutputStream.open(temp_file.path) do |zos|
        @zip_file.entries.each do |e|
          unless @document_contents.keys.include? e.name
            zos.put_next_entry(e.name)
            zos.print e.get_input_stream.read
          end
        end

        @document_contents.each do |name, content|
          zos.put_next_entry(name)
          zos.print content
        end
      end

      if new_path.nil?
        path = @zip_file.name
        FileUtils.rm(path)
      else
        path = new_path
      end
      FileUtils.mv(temp_file.path, path)
      @zip_file = Zip::File.new(path)
    end
  end
end
