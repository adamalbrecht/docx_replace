# encoding: UTF-8

require "docx_replace/version"
require 'zip'
require 'tempfile'

module DocxReplace
  class Doc
    attr_reader :document_content

    def initialize(path, temp_dir=nil)
      @zip_file = Zip::File.new(path)
      @document_file_paths = find_query_file_paths()
      @temp_dir = temp_dir
      read_docx_files
    end

    def replace(pattern, replacement, multiple_occurrences=false)
      replace = replacement.to_s.encode(xml: :text)
      @document_content.each do |path, document|
        if multiple_occurrences
          document.force_encoding("UTF-8").gsub!(pattern, replace)
        else
          document.force_encoding("UTF-8").sub!(pattern, replace)
        end
      end
    end

    def matches(pattern)
      @document_content.values.join.scan(pattern).map{|match| match.first}
    end

    def unique_matches(pattern)
      matches(pattern)
    end

    alias_method :uniq_matches, :unique_matches


    def commit(new_path=nil)
      write_back_to_file(new_path)
    end

    private

    def find_query_file_paths
      @zip_file.entries.map(&:name).select do |entry|
        !(/^word\/(document|footer[0-9]+|header[0-9]+).xml$/ =~ entry).nil?
      end
    end

    def read_docx_files
      @document_content = {}
      @document_file_paths.each do |path|
        @document_content[path] = @zip_file.read(path)
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
          unless @document_file_paths.include?(e.name)
            zos.put_next_entry(e.name)
            zos.print e.get_input_stream.read
          end
        end

        @document_content.each do |path, document|
          zos.put_next_entry(path)
          zos.print document
        end
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
