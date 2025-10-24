require_relative 'helpers/watermark_helper'

module Htmltoword
  class Document
    include XSLTHelper
    # 1 cm = 567 twips (Unidad de medida de Word)
    CM_TO_TWIPS = 567
    DEFAULT_HEADER_FOOTER_TWIPS = 708
    DEFAULT_GUTTER_TWIPS = 0

    class << self
      include TemplatesHelper
      include Htmltoword::Helpers::WatermarkHelper
      def create(content, template_name = nil, extras = false, margins: nil, watermark: nil)
        template_name += extension if template_name && !template_name.end_with?(extension)
        document = new(template_file(template_name))
        document.replace_files(content, extras)
        docx_content = document.generate
        docx_content = apply_margins_to_docx(docx_content, margins) if margins
        docx_content = apply_watermark_to_docx(docx_content, watermark) if watermark
        docx_content
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

      def apply_watermark_to_docx(docx_content, watermark_data)
        image = extract_watermark_image(watermark_data)
        return docx_content unless image

        Tempfile.create(['htmltoword_input', '.docx']) do |input_file|
          input_file.binmode
          input_file.write(docx_content)
          input_file.close

          Zip::File.open(input_file.path) do |zip|
            document_xml = zip.find_entry('word/document.xml')&.get_input_stream&.read
            relations_xml = zip.find_entry('word/_rels/document.xml.rels')&.get_input_stream&.read
            content_types_xml = zip.find_entry('[Content_Types].xml')&.get_input_stream&.read

            next File.binread(input_file.path) unless document_xml && relations_xml && content_types_xml

            header_path = 'word/header_watermark.xml'
            header_target = header_path.sub('word/', '')
            header_rel_path = 'word/_rels/header_watermark.xml.rels'
            image_target = "media/#{image[:filename]}"

            relations_doc = Nokogiri::XML(relations_xml)
            existing_ids = relations_doc.xpath('//xmlns:Relationship').map { |node| node['Id'] }
            header_rel_id = generate_unique_relationship_id(existing_ids)

            # Remove previous references to our header to avoid duplicates
            relations_doc.xpath("//xmlns:Relationship[@Target='#{header_target}']").each(&:remove)

            relationship_node = Nokogiri::XML::Node.new('Relationship', relations_doc)
            relationship_node['Id'] = header_rel_id
            relationship_node['Type'] = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'
            relationship_node['Target'] = header_target
            relations_doc.root.add_child(relationship_node)

            extent = calculate_watermark_extent(document_xml, image)
            updated_document_xml = add_header_reference_to_document(document_xml, header_rel_id)
            zip.get_output_stream('word/document.xml') { |out| out.write(updated_document_xml) }
            zip.get_output_stream('word/_rels/document.xml.rels') { |out| out.write(relations_doc.to_xml) }

            header_xml = build_watermark_header_xml(image_rel_id: 'rId1', extent: extent)
            # Target in a part relationships file is resolved relative to the part (word/header_*.xml)
            # The image lives in word/media/, so the correct relative target is 'media/...'
            header_rels_xml = build_header_relationships_xml(image_rel_id: 'rId1', image_target: image_target)

            zip.get_output_stream(header_path) { |out| out.write(header_xml) }
            zip.get_output_stream(header_rel_path) { |out| out.write(header_rels_xml) }
            zip.get_output_stream("word/#{image_target}") { |out| out.write(image[:data]) }

            content_types_doc = Nokogiri::XML(content_types_xml)
            ensure_header_content_type(content_types_doc, header_path)
            ensure_image_content_type(content_types_doc, image[:extension])
            zip.get_output_stream('[Content_Types].xml') { |out| out.write(content_types_doc.to_xml) }
          end

          File.binread(input_file.path)
        end
      end

      def generate_unique_relationship_id(existing_ids)
        index = existing_ids.filter_map do |id|
          id.to_s.sub(/^rId/, '')
             .then { |num| Integer(num, exception: false) }
        end.max.to_i + 1

        loop do
          candidate = "rId#{index}"
          return candidate unless existing_ids.include?(candidate)
          index += 1
        end
      end

      def add_header_reference_to_document(document_xml, header_rel_id)
        doc = Nokogiri::XML(document_xml)
        namespaces = { 'w' => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                       'r' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships' }

        unless doc.root.namespaces.value?('http://schemas.openxmlformats.org/officeDocument/2006/relationships')
          doc.root.add_namespace_definition('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        end

        doc.xpath('//w:sectPr', namespaces).each do |sect_pr|
          sect_pr.xpath("w:headerReference[@r:id='#{header_rel_id}']", namespaces).each(&:remove)
          header_reference = Nokogiri::XML::Node.new('w:headerReference', doc)
          header_reference['w:type'] = 'default'
          header_reference['r:id'] = header_rel_id
          sect_pr.add_child(header_reference)
        end

        doc.to_xml(save_with: Nokogiri::XML::Node::SaveOptions::AS_XML)
      end


      def ensure_header_content_type(content_types_doc, header_path)
        part_name = "/#{header_path}"
        return if content_types_doc.at_xpath("//xmlns:Override[@PartName='#{part_name}']")

        override_node = Nokogiri::XML::Node.new('Override', content_types_doc)
        override_node['PartName'] = part_name
        override_node['ContentType'] = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
        content_types_doc.root.add_child(override_node)
      end

      def ensure_image_content_type(content_types_doc, extension)
        ext = extension == 'jpeg' ? 'jpg' : extension
        return if content_types_doc.at_xpath("//xmlns:Default[@Extension='#{ext}']")

        default_node = Nokogiri::XML::Node.new('Default', content_types_doc)
        default_node['Extension'] = ext
        default_node['ContentType'] = "image/#{extension}"
        content_types_doc.root.add_child(default_node)
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
            # stream the image files into the media folder
            @image_files.each do |hash|
              out.put_next_entry("word/media/#{hash[:filename]}")
              if hash[:data]
                out.write(hash[:data])
              else
                URI.open(hash[:url], 'rb') do |f|
                  out.write(f.read)
                end
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
      source.css('img').each_with_index do |image, i|
        src = image['src'].to_s
        next if src.empty?

        data_image = parse_data_image(src)
        provided_filename = image['data-filename']

        if data_image
          ext = data_image[:ext]
          filename = sanitize_filename(provided_filename.presence || "image#{i + 1}.#{ext}")
          @image_files << { filename: filename, data: data_image[:data], ext: ext }
        else
          # Remote/absolute URL (or relative); derive extension from filename if present
          inferred_filename = provided_filename.presence || src.split('/').last.to_s
          ext = File.extname(inferred_filename).delete('.').downcase
          # fallback if no extension could be inferred
          ext = 'png' if ext.empty?
          filename = sanitize_filename("image#{i + 1}.#{ext}")
          @image_files << { filename: filename, url: src, ext: ext }
        end
      end
    end

    def parse_data_image(src)
      # Supports data URI images: data:image/<type>;base64,<payload>
      m = src.match(%r{\Adata:(image/[^;]+);base64,(.+)\z}m)
      return nil unless m

      mime = m[1]
      b64 = m[2]
      ext = case mime
            when 'image/jpeg', 'image/jpg' then 'jpg'
            when 'image/png' then 'png'
            else nil
            end
      return nil unless ext

      { data: Base64.decode64(b64), ext: ext }
    rescue ArgumentError
      nil
    end

    def sanitize_filename(name)
      # Minimal sanitization to avoid invalid zip entry names
      name.gsub(/[\\\:\\*\?\"\<\>\|]/, '_')
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
