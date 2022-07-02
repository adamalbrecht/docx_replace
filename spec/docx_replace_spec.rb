require 'spec_helper'
require 'docx_replace'
require 'zip'

describe DocxReplace::Doc do
  let(:output_file) { Tempfile.new('docx_replace_tmp_output') }
  let(:output_text) { docx_content(output_file) }

  it "replaces a single variable in a very basic document" do
    doc = described_class.new(get_fixture("basic.docx"))
    doc.replace("$foobar$", "hello world")
    doc.commit(output_file)

    expect(output_text).to match(/hello world/)
    expect(output_text).not_to match(/\$foobar\$/)
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
