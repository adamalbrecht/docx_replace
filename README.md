# Docx Replace

This gem allows you to generate .docx files in your rails or ruby app by
embedding variables inside of a .docx template. This is purposefully
meant to be simple and feature-light.

## Installation

Add this line to your application's Gemfile:

    gem 'docx_replace'

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install docx_replace

## Usage

Inside of a rails controller, your code might look something like this (although I would recommend extracting most of this into a separate class):

```ruby
def user_report
  @user = User.find(params[:user_id])

  respond_to do |format|
    format.docx do
      # Initialize DocxReplace with your template
      doc = DocxReplace::Doc.new("#{Rails.root}/lib/docx_templates/my_template.docx", "#{Rails.root}/tmp")

      # Replace some variables. $var$ convention is used here, but not required.
      doc.replace("$first_name$", @user.first_name)
      doc.replace("$last_name$", @user.last_name)
      doc.replace("$user_bio$", @user.bio)
      
      # Write the document back to a temporary file
      tmp_file = Tempfile.new('word_tempate', "#{Rails.root}/tmp")
      doc.commit(tmp_file.path)

      # Respond to the request by sending the temp file
      send_file tmp_file.path, filename: "user_#{@user.id}_report.docx", disposition: 'attachment'
    end
  end
end
```


## Contributing

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request

## Credits

Much of this code is based on an older gem called [docxedit](https://github.com/oliamb/docxedit). This has a few more features, but is very sensitive to the formatting of the .docx template.
