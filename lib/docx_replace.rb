# encoding: UTF-8

require "docx_replace/version"
require 'zip'

module DocxReplace
  class Doc
    attr_reader :document_content, :entries

    def initialize(stream)
      @stream = stream
      @entries = []
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

    def commit
      write_back_to_file
    end

    private
    DOCUMENT_FILE_PATH = 'word/document.xml'

    def read_docx_file
      Zip::InputStream.open(StringIO.new(@stream)) do |io|
        while entry = io.get_next_entry
          obj = {name: entry.name, content: io.read}
          @entries << obj

          if entry.name == DOCUMENT_FILE_PATH
            @document_content = obj[:content]
          end
        end
      end
    end

    def write_back_to_file
      compressed_filestream = Zip::OutputStream.write_buffer do |zos|
        @entries.each do |entry|
          unless entry[:name] == DOCUMENT_FILE_PATH
            zos.put_next_entry(entry[:name])
            zos.print entry[:content]
          end
        end

        zos.put_next_entry(DOCUMENT_FILE_PATH)
        zos.print @document_content
      end

      compressed_filestream.string
    end
  end
end
