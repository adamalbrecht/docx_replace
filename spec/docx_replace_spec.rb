require 'spec_helper'
require 'docx_replace'
require 'zip'

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

  private

  def get_fixture(name)
    File.join(project_root, "spec", "fixtures", name)
  end

  def project_root
    File.expand_path(File.dirname(File.dirname(__FILE__)))
  end

  def docx_content(path)
    zip_file = Zip::File.new(path)
    zip_file.read("word/document.xml")
  end
end
