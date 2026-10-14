require 'spec_helper'
require 'securerandom'

describe Htmltoword::Document do
  describe "local_images" do
    let(:html) do
      <<-EOL
        <!DOCTYPE html>
        <html>
        <head></head>
        <body>
          <img src="https://placehold.it/50x50.png" style="width:50px;height:50px">
          <a href="http://www.example.com">
            <img src="https://placehold.it/10x10.png" style="width:10px;height:10px">
            Hello!
          </a>
          <p><img src="https://placehold.it/350x150.png" style="width:350px;height:150px"></p>
          <p><img src="https://placehold.it/400x250.jpg" style="width:400px;height:250px"></p>
          <p><img src="https://placehold.it/350x150.tiff" style="width:350px;height:150px"></p>
          <p><img src="https://placehold.it/400x250.gif" style="width:400px;height:250px"></p>
        </body>
        </html>
      EOL
    end

    # placehold.it no longer serves images (HTTP 403), so the download is stubbed and the example
    # keeps covering the embedding of the six remote images in word/media.
    before do
      png = "\x89PNG\r\n\x1a\n".b + 'payload'
      allow(URI).to receive(:open) { |*, &block| block.call(StringIO.new(png)) }
    end

    #in order to test if the images files are embedded correctly we have to create_and_save, then open the zip and look for them. Then ensure that we cleanup the created file.
    #There is probably a better way of doing it as this seems very performance heavy but does have the advantage of testing the create method all the way through...
    it "should only embed local images" do
      filename = SecureRandom.urlsafe_base64
      begin
        docx = Htmltoword::Document.create_and_save(html, tmp_path(filename))
        Zip::File.open(tmp_path(filename)) do |zip_file|
          # Find specific entry
          expect(zip_file.glob('word/media/*').size).to eq 6
        end
      ensure
        File.delete(tmp_path(filename))
      end
    end

    context "when the image name has several dots" do
      let(:html) { '<html><body><p><img src="https://example.com/photos/avatar.v1.webp" style="width:50px;height:50px"></p></body></html>' }

      it "points the relation to the file stored in word/media" do
        docx = Htmltoword::Document.create(html)

        Zip::File.open_buffer(docx) do |zip|
          expect(zip.glob('word/media/*').map(&:name)).to eq ['word/media/image1.webp']
          relations = zip.read('word/_rels/document.xml.rels')
          expect(relations).to include 'Target="media/image1.webp"'
          expect(relations).not_to include 'v1'
        end
      end
    end
  end

  describe "image_resolver" do
    let(:png) { "\x89PNG\r\n\x1a\n".b + 'payload' }
    let(:html) do
      <<-EOL
        <!DOCTYPE html>
        <html>
        <head></head>
        <body>
          <p><img src="https://example.com/photos/avatar.webp" style="width:50px;height:50px"></p>
          <p><img src="https://example.com/photos/logo.png" style="width:50px;height:50px"></p>
        </body>
        </html>
      EOL
    end

    after { Htmltoword.config.image_resolver = nil }

    it "embeds the resolved bytes with their extension and updates the relation" do
      Htmltoword.config.image_resolver = ->(src) { { data: png, ext: 'png' } if src.end_with?('.webp') }
      document = Htmltoword::Document.new(template_file(nil))
      document.replace_files(html)

      expect(document.instance_variable_get(:@image_files)).to eq [
        { filename: 'image1.png', data: png, ext: 'png' },
        { filename: 'image2.png', url: 'https://example.com/photos/logo.png', ext: 'png' }
      ]
      relations = document.instance_variable_get(:@replaceable_files)[Htmltoword::Document.relations_xml_file]
      expect(relations).to include 'Target="media/image1.png"'
      expect(relations).to include 'Target="media/image2.png"'
      expect(relations).not_to include '.webp'
    end

    it "points the relation to the converted file when the image name has several dots" do
      Htmltoword.config.image_resolver = ->(src) { { data: png, ext: 'png' } if src.end_with?('.webp') }
      docx = Htmltoword::Document.create('<html><body><p><img src="https://example.com/photos/avatar.v1.webp" style="width:50px;height:50px"></p></body></html>')

      Zip::File.open_buffer(docx) do |zip|
        expect(zip.glob('word/media/*').map(&:name)).to eq ['word/media/image1.png']
        expect(zip.read('word/media/image1.png')).to eq png
        relations = zip.read('word/_rels/document.xml.rels')
        expect(relations).to include 'Target="media/image1.png"'
        expect(relations).not_to include 'webp'
      end
    end

    it "writes the resolved bytes into the generated document" do
      Htmltoword.config.image_resolver = ->(_src) { { data: png, ext: 'png' } }
      docx = Htmltoword::Document.create(html)

      Zip::File.open_buffer(docx) do |zip|
        expect(zip.read('word/media/image1.png')).to eq png
        expect(zip.read('word/media/image2.png')).to eq png
        expect(zip.read('[Content_Types].xml')).to match(/Extension=["']png["']/)
      end
    end

    it "keeps the inferred extension when the resolver does not return one" do
      Htmltoword.config.image_resolver = ->(_src) { { data: png } }
      document = Htmltoword::Document.new(template_file(nil))
      document.replace_files(html)

      expect(document.instance_variable_get(:@image_files).map { |image| image[:filename] }).to eq ['image1.webp', 'image2.png']
    end

    it "downloads the image as usual when the resolver returns nil" do
      Htmltoword.config.image_resolver = ->(_src) { nil }
      document = Htmltoword::Document.new(template_file(nil))
      document.replace_files(html)

      expect(document.instance_variable_get(:@image_files)).to eq [
        { filename: 'image1.webp', url: 'https://example.com/photos/avatar.webp', ext: 'webp' },
        { filename: 'image2.png', url: 'https://example.com/photos/logo.png', ext: 'png' }
      ]
    end
  end
end
