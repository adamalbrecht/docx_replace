require 'spec_helper'
require 'docx_replace'
require 'zip'
require 'tmpdir'

describe DocxReplace::Doc do
  let(:output_file) { Tempfile.new('docx_replace_tmp_output') }
  let(:output_text) { docx_content(output_file) }

  it "replaces a single variable in a very basic document" do
    doc = described_class.new(get_fixture("basic.docx"))
    doc.replace("FOOBAR", "hello world")
    doc.commit(output_file)

    expect(output_text).to match(/hello world/)
    expect(output_text).not_to match(/FOOBAR/)
  end

  it "can replace multiple occurrences of the same variable" do
    doc = described_class.new(get_fixture("multiple.docx"))
    doc.replace("FOOBAR", "hello world", true)
    doc.commit(output_file)

    expect(output_text).not_to match(/FOOBAR/)
    expect(output_text.scan("hello world").size).to eq(2)
  end

  it "does not replace multiple occurrences unless instructed to do so" do
    doc = described_class.new(get_fixture("multiple.docx"))
    doc.replace("FOOBAR", "hello world", false)
    doc.commit(output_file)

    expect(output_text.scan("FOOBAR").size).to eq(1)
    expect(output_text.scan("hello world").size).to eq(1)
  end

  describe "#replace" do
    it "accepts a regular expression pattern" do
      doc = described_class.new(get_fixture("basic.docx"))
      doc.replace(/FOO\w+/, "hello world")
      doc.commit(output_file)

      expect(output_text).to include("hello world")
      expect(output_text).not_to include("FOOBAR")
    end

    it "escapes XML special characters in the replacement" do
      doc = described_class.new(get_fixture("basic.docx"))
      doc.replace("FOOBAR", "Tom & <Jerry>")
      doc.commit(output_file)

      expect(output_text).to include("Tom &amp; &lt;Jerry&gt;")
      expect(output_text).not_to include("<Jerry>")
      expect { parse_xml(output_text) }.not_to raise_error
    end

    it "converts non-string replacements with to_s" do
      doc = described_class.new(get_fixture("basic.docx"))
      doc.replace("FOOBAR", 42)
      doc.commit(output_file)

      expect(output_text).to include(" 42.")
    end

    it "handles non-ASCII replacements" do
      doc = described_class.new(get_fixture("basic.docx"))
      doc.replace("FOOBAR", "Café ñandú 日本語")
      doc.commit(output_file)

      expect(output_text.force_encoding("UTF-8")).to include("Café ñandú 日本語")
    end

    it "leaves the document unchanged when the pattern does not match" do
      original = docx_content(get_fixture("basic.docx"))
      doc = described_class.new(get_fixture("basic.docx"))
      doc.replace("NOT_PRESENT", "hello world", true)
      doc.commit(output_file)

      expect(output_text).to eq(original)
    end

    it "replaces the same variable in the body, headers and footers" do
      path = build_docx_with_headers_and_footers
      doc = described_class.new(path)
      doc.replace("FOOBAR", "hello world", true)
      doc.commit(output_file)

      %w[word/document.xml word/header1.xml word/header2.xml word/footer1.xml].each do |part|
        text = docx_content(output_file, part)
        expect(text).to include("hello world"), "expected #{part} to be replaced"
        expect(text).not_to include("FOOBAR"), "expected #{part} to have no FOOBAR left"
      end
    end

    it "replaces only the first occurrence overall when multiple_occurrences is false" do
      path = build_docx_with_headers_and_footers
      doc = described_class.new(path)
      doc.replace("FOOBAR", "hello world", false)

      # sub! runs once per part, so each part that contains the variable is changed once
      expect(doc.document_contents.values.join.scan("hello world").size).to eq(doc.document_contents.size)
    end

    it "ignores parts that are not the document, headers or footers" do
      path = build_docx_with_headers_and_footers(extra: { "word/other1.xml" => "<x>FOOBAR</x>" })
      doc = described_class.new(path)

      expect(doc.document_contents.keys).not_to include("word/other1.xml")
      expect(doc.document_contents.keys).to match_array(
        %w[word/document.xml word/header1.xml word/header2.xml word/footer1.xml]
      )
    end
  end

  describe "#matches" do
    it "returns the first capture group of each match" do
      doc = described_class.new(get_fixture("multiple.docx"))

      expect(doc.matches(/(FOO)BAR/)).to eq(%w[FOO FOO])
    end

    it "returns an empty array when nothing matches" do
      doc = described_class.new(get_fixture("basic.docx"))

      expect(doc.matches(/(NOT_PRESENT)/)).to eq([])
    end

    it "searches headers and footers as well" do
      doc = described_class.new(build_docx_with_headers_and_footers)

      expect(doc.matches(/(FOOBAR)/).size).to eq(4)
    end
  end

  describe "#unique_matches" do
    it "returns matches for the pattern" do
      doc = described_class.new(get_fixture("multiple.docx"))

      expect(doc.unique_matches(/(FOO)BAR/)).to eq(doc.matches(/(FOO)BAR/))
    end

    it "is aliased as uniq_matches" do
      doc = described_class.new(get_fixture("multiple.docx"))

      expect(doc.uniq_matches(/(FOO)BAR/)).to eq(doc.unique_matches(/(FOO)BAR/))
    end
  end

  describe "#commit" do
    it "overwrites the original file when no path is given" do
      path = copy_fixture_to_tmp("basic.docx")
      doc = described_class.new(path)
      doc.replace("FOOBAR", "hello world")
      doc.commit

      text = docx_content(path)
      expect(text).to include("hello world")
      expect(text).not_to include("FOOBAR")
    end

    it "leaves the original file untouched when a new path is given" do
      path = copy_fixture_to_tmp("basic.docx")
      new_path = File.join(tmp_dir, "new.docx")
      doc = described_class.new(path)
      doc.replace("FOOBAR", "hello world")
      doc.commit(new_path)

      expect(docx_content(path)).to include("FOOBAR")
      expect(docx_content(new_path)).to include("hello world")
    end

    it "preserves every other entry in the archive" do
      source = get_fixture("basic.docx")
      doc = described_class.new(source)
      doc.replace("FOOBAR", "hello world")
      doc.commit(output_file)

      source_entries = zip_entries(source)
      output_entries = zip_entries(output_file)

      expect(output_entries.keys).to match_array(source_entries.keys)
      (source_entries.keys - ["word/document.xml"]).each do |name|
        expect(output_entries[name]).to eq(source_entries[name]), "expected #{name} to be unchanged"
      end
    end

    it "writes a docx that can be reopened and edited again" do
      first = File.join(tmp_dir, "first.docx")
      second = File.join(tmp_dir, "second.docx")

      doc = described_class.new(get_fixture("basic.docx"))
      doc.replace("FOOBAR", "hello world")
      doc.commit(first)

      reopened = described_class.new(first)
      reopened.replace("hello world", "goodbye")
      reopened.commit(second)

      expect(docx_content(second)).to include("goodbye")
      expect(docx_content(second)).not_to include("hello world")
    end

    it "can be called more than once on the same instance" do
      first = File.join(tmp_dir, "first.docx")
      second = File.join(tmp_dir, "second.docx")

      doc = described_class.new(get_fixture("multiple.docx"))
      doc.replace("FOOBAR", "one")
      doc.commit(first)
      doc.replace("FOOBAR", "two")
      doc.commit(second)

      expect(docx_content(first)).to include("one")
      expect(docx_content(first)).to include("FOOBAR")
      expect(docx_content(second)).to include("one")
      expect(docx_content(second)).to include("two")
      expect(docx_content(second)).not_to include("FOOBAR")
    end
  end

  describe "temp_dir" do
    before { output_file } # create the lazy Tempfile before Tempfile.new is stubbed

    it "creates the temporary file in the given directory" do
      expect(Tempfile).to receive(:new).with('docxedit-', tmp_dir).and_call_original

      doc = described_class.new(get_fixture("basic.docx"), tmp_dir)
      doc.replace("FOOBAR", "hello world")
      doc.commit(output_file)

      expect(output_text).to include("hello world")
    end

    it "defaults to the system temp directory" do
      expect(Tempfile).to receive(:new).with('docxedit-').and_call_original

      doc = described_class.new(get_fixture("basic.docx"))
      doc.commit(output_file)
    end
  end

  describe "error handling" do
    it "raises when the file does not exist" do
      expect { described_class.new(File.join(tmp_dir, "missing.docx")) }.to raise_error(Zip::Error)
    end

    it "raises when the file is not a zip archive" do
      path = File.join(tmp_dir, "not_a_docx.docx")
      File.write(path, "this is not a zip file")

      expect { described_class.new(path) }.to raise_error(Zip::Error)
    end
  end

  private

  def get_fixture(name)
    File.join(project_root, "spec", "fixtures", name)
  end

  def project_root
    File.expand_path(File.dirname(File.dirname(__FILE__)))
  end

  def docx_content(path, part = "word/document.xml")
    path = path.path if path.respond_to?(:path)
    Zip::File.open(path) { |zip_file| zip_file.read(part) }
  end

  def zip_entries(path)
    path = path.path if path.respond_to?(:path)
    Zip::File.open(path) do |zip_file|
      zip_file.entries.each_with_object({}) { |e, h| h[e.name] = e.get_input_stream.read }
    end
  end

  def tmp_dir
    @tmp_dir ||= Dir.mktmpdir('docx_replace_spec')
  end

  after { FileUtils.remove_entry(@tmp_dir) if @tmp_dir && File.directory?(@tmp_dir) }

  def copy_fixture_to_tmp(name)
    dest = File.join(tmp_dir, name)
    FileUtils.cp(get_fixture(name), dest)
    dest
  end

  # Builds a docx from basic.docx with two headers and a footer that each
  # contain the FOOBAR variable, so header/footer support can be exercised.
  def build_docx_with_headers_and_footers(extra: {})
    path = File.join(tmp_dir, "headers_footers.docx")
    parts = {
      "word/header1.xml" => "<w:hdr><w:p><w:r><w:t>FOOBAR header one</w:t></w:r></w:p></w:hdr>",
      "word/header2.xml" => "<w:hdr><w:p><w:r><w:t>FOOBAR header two</w:t></w:r></w:p></w:hdr>",
      "word/footer1.xml" => "<w:ftr><w:p><w:r><w:t>FOOBAR footer</w:t></w:r></w:p></w:ftr>"
    }.merge(extra)

    Zip::OutputStream.open(path) do |zos|
      zip_entries(get_fixture("basic.docx")).each do |name, content|
        content = content.sub("</w:t>", " FOOBAR</w:t>") if name == "word/document.xml" && !content.include?("FOOBAR")
        zos.put_next_entry(name)
        zos.print content
      end
      parts.each do |name, content|
        zos.put_next_entry(name)
        zos.print content
      end
    end
    path
  end

  # Raises if the string is not well-formed XML, without adding a dependency.
  def parse_xml(xml)
    require 'rexml/document'
    REXML::Document.new(xml).tap do |doc|
      raise "malformed XML" if doc.root.nil?
    end
  end
end
