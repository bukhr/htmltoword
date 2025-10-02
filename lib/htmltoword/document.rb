module Htmltoword
  class Document
    include XSLTHelper
    # 1 cm = 567 twips (Unidad de medida de Word)
    CM_TO_TWIPS = 567
    DEFAULT_HEADER_FOOTER_TWIPS = 708
    DEFAULT_GUTTER_TWIPS = 0

    class << self
      include TemplatesHelper
      def create(content, template_name = nil, extras = false, margins: nil)
        template_name += extension if template_name && !template_name.end_with?(extension)
        document = new(template_file(template_name))
        document.replace_files(content, extras)
        docx_content = document.generate
        margins ? apply_margins_to_docx(docx_content, margins) : docx_content
      end

      def create_and_save(content, file_path, template_name = nil, extras = false)
        File.open(file_path, 'wb') do |out|
          out << create(content, template_name, extras)
        end
      end

      def create_with_content(template, content, extras = false)
        template += extension unless template.end_with?(extension)
        document = new(template_file(template))
        document.replace_files(content, extras)
        document.generate
      end

      def extension
        '.docx'
      end

      def doc_xml_file
        'word/document.xml'
      end

      def numbering_xml_file
        'word/numbering.xml'
      end

      def relations_xml_file
        'word/_rels/document.xml.rels'
      end

      def content_types_xml_file
        '[Content_Types].xml'
      end

      private

      def apply_margins_to_docx(docx_content, margins_cm)
        margin_twips = margins_cm.transform_values { |cm| (cm.to_f * CM_TO_TWIPS).to_i }
        Tempfile.create(['htmltoword_input', '.docx']) do |input_file|
          input_file.binmode
          input_file.write(docx_content)
          input_file.close

          Zip::File.open(input_file.path) do |zip|
            entry = zip.find_entry('word/document.xml')
            xml_content = entry.get_input_stream.read
            modified_xml = modify_document_xml_margins(xml_content, margin_twips)
            zip.get_output_stream(entry.name) { |out| out.write(modified_xml) }
          end

          File.binread(input_file.path)
        end
      end

      def modify_document_xml_margins(xml_content, margin_twips)
        doc = Nokogiri::XML(xml_content)
        # Namespace XML de Word (requerido para consultas XPath).
        # Define el prefijo "w:" para buscar elementos en el XML interno del documento Word.
        # Ejemplo: "w:sectPr" se busca como <w:sectPr> dentro del archivo word/document.xml
        namespaces = { 'w' => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' }

        # XPath para encontrar nodos <w:sectPr> (Section Properties)
        # que contienen la configuración de márgenes y propiedades de página
        # según especificación WordprocessingML de Microsoft
        doc.xpath('//w:sectPr', namespaces).each do |sect_pr|
          pg_mar = sect_pr.at_xpath('w:pgMar', namespaces)

          unless pg_mar
            pg_mar = Nokogiri::XML::Node.new('w:pgMar', doc)
            pg_mar.namespace = sect_pr.namespace
            sect_pr.add_child(pg_mar)
          end

          margin_sides = [:top, :right, :bottom, :left]
          margin_sides.each { |side| pg_mar["w:#{side}"] = margin_twips[side].to_s }
          pg_mar['w:header'] ||= DEFAULT_HEADER_FOOTER_TWIPS.to_s
          pg_mar['w:footer'] ||= DEFAULT_HEADER_FOOTER_TWIPS.to_s
          pg_mar['w:gutter'] ||= DEFAULT_GUTTER_TWIPS.to_s
        end

        doc.to_xml(save_with: Nokogiri::XML::Node::SaveOptions::AS_XML)
      end
    end

    def initialize(template_path)
      @replaceable_files = {}
      @template_path = template_path
      @image_files = []
    end

    #
    # Generate a string representing the contents of a docx file.
    #
    def generate
      Zip::File.open(@template_path) do |template_zip|
        buffer = Zip::OutputStream.write_buffer do |out|
          template_zip.each do |entry|
            out.put_next_entry entry.name
            if @replaceable_files[entry.name] && entry.name == Document.doc_xml_file
              source = entry.get_input_stream.read
              # Change only the body of document. TODO: Improve this...
              source = source.sub(/(<w:body>)((.|\n)*?)(<w:sectPr)/, "\\1#{@replaceable_files[entry.name]}\\4")
              out.write(source)
            elsif @replaceable_files[entry.name]
              out.write(@replaceable_files[entry.name])
            elsif entry.name == Document.content_types_xml_file
              raw_file = entry.get_input_stream.read
              content_types = @image_files.empty? ? raw_file : inject_image_content_types(raw_file)

              out.write(content_types)
            else
              out.write(template_zip.read(entry.name))
            end
          end
          unless @image_files.empty?
          #stream the image files into the media folder using open-uri
            @image_files.each do |hash|
              out.put_next_entry("word/media/#{hash[:filename]}")
              URI.open(hash[:url], 'rb') do |f|
                out.write(f.read)
              end
            end
          end
        end
        buffer.string
      end
    end

    def replace_files(html, extras = false)
      html = '<body></body>' if html.nil? || html.empty?
      original_source = Nokogiri::HTML(html.gsub(/>\s+</, '><'))
      source = xslt(stylesheet_name: 'cleanup').transform(original_source)
      transform_and_replace(source, xslt_path('numbering'), Document.numbering_xml_file)
      transform_and_replace(source, xslt_path('relations'), Document.relations_xml_file)
      transform_doc_xml(source, extras)
      local_images(source)
    end

    def transform_doc_xml(source, extras = false)
      transformed_source = xslt(stylesheet_name: 'cleanup').transform(source)
      transformed_source = xslt(stylesheet_name: 'inline_elements').transform(transformed_source)
      transform_and_replace(transformed_source, document_xslt(extras), Document.doc_xml_file, extras)
    end

    private

    def transform_and_replace(source, stylesheet_path, file, remove_ns = false)
      stylesheet = xslt(stylesheet_path: stylesheet_path)
      content = stylesheet.apply_to(source)
      content.gsub!(/\s*xmlns:(\w+)="(.*?)\s*"/, '') if remove_ns
      @replaceable_files[file] = content
    end

    #generates an array of hashes with filename and full url
    #for all images to be embeded in the word document
    def local_images(source)
      source.css('img').each_with_index do |image,i|
        filename = image['data-filename'] ? image['data-filename'] : image['src'].split("/").last
        ext = File.extname(filename).delete(".").downcase

        @image_files << { filename: "image#{i+1}.#{ext}", url: image['src'], ext: ext }
      end
    end

    #get extension from filename and clean to match content_types
    def content_type_from_extension(ext)
      ext == "jpg" ? "jpeg" : ext
    end

    #inject the required content_types into the [content_types].xml file...
    def inject_image_content_types(source)
      doc = Nokogiri::XML(source)

      #get a list of all extensions currently in content_types file
      existing_exts = doc.css("Default").map { |node| node.attribute("Extension").value }.compact

      #get a list of extensions we need for our images
      required_exts = @image_files.map{ |i| i[:ext] }

      #workout which required extensions are missing from the content_types file
      missing_exts = (required_exts - existing_exts).uniq

      #inject missing extensions into document
      missing_exts.each do |ext|
        doc.at_css("Types").add_child( "<Default Extension='#{ext}' ContentType='image/#{content_type_from_extension(ext)}'/>")
      end

      #return the amended source to be saved into the zip
      doc.to_s
    end
  end
end
