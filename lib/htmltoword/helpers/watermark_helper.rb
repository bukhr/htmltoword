# frozen_string_literal: true

module Htmltoword
  module Helpers
    module WatermarkHelper
      WATERMARK_MIME_EXTENSION = {
        'image/png' => 'png',
        'image/jpeg' => 'jpeg',
        'image/jpg' => 'jpeg'
      }.freeze

      WATERMARK_EXTENT = {
        'cx' => '9000000',
        'cy' => '9000000'
      }.freeze

      WATERMARK_NAMESPACES = {
        'xmlns:w' => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'xmlns:r' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'xmlns:wp' => 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'xmlns:a' => 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'xmlns:pic' => 'http://schemas.openxmlformats.org/drawingml/2006/picture'
      }.freeze

      WATERMARK_ANCHOR_ATTRIBUTES = {
        'distT' => 0,
        'distB' => 0,
        'distL' => 0,
        'distR' => 0,
        'simplePos' => 0,
        'relativeHeight' => 251_658_240,
        'behindDoc' => 1,
        'locked' => 0,
        'layoutInCell' => 1,
        'allowOverlap' => 1
      }.freeze

      def extract_watermark_image(watermark_data)
        return nil if watermark_data.to_s.strip.empty?

        metadata, base64_content = watermark_data.split(',', 2)
        base64_content ||= metadata
        mime = metadata[%r{image/[^;]+}] if metadata&.match?(%r{image/[^;]+})
        extension = WATERMARK_MIME_EXTENSION[mime]
        return nil unless extension

        data = Base64.decode64(base64_content)

        dimensions = detect_image_dimensions(data, extension)

        {
          data: data,
          extension: extension,
          filename: "watermark.#{extension == 'jpeg' ? 'jpg' : extension}",
          dimensions: dimensions
        }
      rescue ArgumentError
        nil
      end

      def build_watermark_header_xml(image_rel_id: 'rId1', extent: WATERMARK_EXTENT)
        Nokogiri::XML::Builder.new do |xml|
          xml['w'].hdr(WATERMARK_NAMESPACES) do
            xml['w'].p do
              xml['w'].r do
                xml['w'].drawing do
                  xml['wp'].anchor(WATERMARK_ANCHOR_ATTRIBUTES) do
                    xml['wp'].simplePos('x' => 0, 'y' => 0)
                    xml['wp'].positionH('relativeFrom' => 'page') { xml['wp'].align('center') }
                    xml['wp'].positionV('relativeFrom' => 'page') { xml['wp'].align('center') }
                    xml['wp'].extent(extent)
                    xml['wp'].effectExtent('l' => 0, 't' => 0, 'r' => 0, 'b' => 0)
                    xml['wp'].wrapNone
                    xml['wp'].docPr('id' => '1', 'name' => 'Watermark')
                    xml['wp'].cNvGraphicFramePr { xml['a'].graphicFrameLocks('noChangeAspect' => '1') }
                    xml['a'].graphic do
                      xml['a'].graphicData('uri' => 'http://schemas.openxmlformats.org/drawingml/2006/picture') do
                        xml['pic'].pic do
                          xml['pic'].nvPicPr do
                            xml['pic'].cNvPr('id' => '0', 'name' => 'Watermark')
                            xml['pic'].cNvPicPr
                          end
                          xml['pic'].blipFill do
                            xml['a'].blip('r:embed' => image_rel_id) { xml['a'].alphaModFix('amt' => '80000') }
                            xml['a'].stretch { xml['a'].fillRect }
                          end
                          xml['pic'].spPr do
                            xml['a'].xfrm do
                              xml['a'].off('x' => 0, 'y' => 0)
                              xml['a'].ext(extent)
                            end
                            xml['a'].prstGeom('prst' => 'rect') { xml['a'].avLst }
                          end
                        end
                      end
                    end
                  end
                end
              end
            end
          end
        end.to_xml
      end

      def build_header_relationships_xml(image_rel_id:, image_target:)
        Nokogiri::XML::Builder.new do |xml|
          xml.Relationships('xmlns' => 'http://schemas.openxmlformats.org/package/2006/relationships') do
            xml.Relationship('Id' => image_rel_id,
                             'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                             'Target' => image_target)
          end
        end.to_xml
      end

      private

        def detect_image_dimensions(data, extension)
          case extension
          when 'png'
            parse_png_dimensions(data)
          when 'jpeg'
            parse_jpeg_dimensions(data)
          else
            nil
          end
        end

        def twips_to_emu(twips)
          (twips.to_f * 635).to_i
        end

        def default_page_metrics
          {
            width_twips: 12_240,  # 8.5 in
            height_twips: 15_840, # 11 in
            left_margin_twips: 1_440,
            right_margin_twips: 1_440,
            top_margin_twips: 1_440,
            bottom_margin_twips: 1_440,
          }
        end

        def extract_content_area_dimensions(document_xml)
          doc = Nokogiri::XML(document_xml)
          namespaces = {
            'w' => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
          }

          metrics = default_page_metrics

          if (pg_sz = doc.at_xpath('//w:sectPr/w:pgSz', namespaces) || doc.at_xpath('//w:pgSz', namespaces))
            metrics[:width_twips] = pg_sz['w:w'].to_i if pg_sz['w:w']
            metrics[:height_twips] = pg_sz['w:h'].to_i if pg_sz['w:h']
          end

          if (pg_mar = doc.at_xpath('//w:sectPr/w:pgMar', namespaces) || doc.at_xpath('//w:pgMar', namespaces))
            metrics[:left_margin_twips] = pg_mar['w:left'].to_i if pg_mar['w:left']
            metrics[:right_margin_twips] = pg_mar['w:right'].to_i if pg_mar['w:right']
            metrics[:top_margin_twips] = pg_mar['w:top'].to_i if pg_mar['w:top']
            metrics[:bottom_margin_twips] = pg_mar['w:bottom'].to_i if pg_mar['w:bottom']
          end

          content_width_twips = [metrics[:width_twips] - metrics[:left_margin_twips] - metrics[:right_margin_twips], 1].max
          content_height_twips = [metrics[:height_twips] - metrics[:top_margin_twips] - metrics[:bottom_margin_twips], 1].max

          {
            width_emu: twips_to_emu(content_width_twips),
            height_emu: twips_to_emu(content_height_twips)
          }
        end

        def calculate_watermark_extent(document_xml, image)
          content_area = extract_content_area_dimensions(document_xml)
          return WATERMARK_EXTENT unless content_area

          if image && image[:dimensions]
            width_px, height_px = image[:dimensions]
            if width_px.to_i.positive? && height_px.to_i.positive?
              emu_per_pixel = 9_144_00.0 / 96.0
              image_width_emu = width_px * emu_per_pixel
              image_height_emu = height_px * emu_per_pixel

              max_width = content_area[:width_emu]
              max_height = content_area[:height_emu]

              scale = [max_width / image_width_emu, max_height / image_height_emu, 1.0].min
              scale *= 0.85

              target_width = [(image_width_emu * scale).to_i, 1].max
              target_height = [(image_height_emu * scale).to_i, 1].max

              return {
                'cx' => [target_width, max_width].min.to_i,
                'cy' => [target_height, max_height].min.to_i
              }
            end
          end

          {
            'cx' => content_area[:width_emu],
            'cy' => content_area[:height_emu]
          }
        end

        def parse_png_dimensions(data)
          return nil unless data && data.bytesize >= 24

          signature = data[0, 8]
          return nil unless signature == "\x89PNG\r\n\x1A\n"

          header = data[8, 16]
          chunk_length = header[0, 4].unpack1('N')
          chunk_type = header[4, 4]
          return nil unless chunk_type == 'IHDR' && chunk_length == 13

          width, height = header[8, 8].unpack('N2')
          [width, height]
        rescue StandardError
          nil
        end

        def parse_jpeg_dimensions(data)
          return nil unless data && data.bytesize > 4

          io = StringIO.new(data)
          return nil unless io.read(2) == "\xFF\xD8"

          loop do
            marker = io.read(2)
            return nil unless marker && marker.bytesize == 2

            while marker.getbyte(0) != 0xFF
              marker = marker[1] + io.read(1)
              return nil unless marker
            end

            marker_byte = marker.getbyte(1)

            if marker_byte >= 0xC0 && marker_byte <= 0xC3
              length = io.read(2).unpack1('n')
              precision = io.read(1)
              height = io.read(2).unpack1('n')
              width = io.read(2).unpack1('n')
              return [width, height]
            elsif marker_byte == 0xD9 || marker_byte == 0xDA
              break
            else
              length = io.read(2).unpack1('n')
              io.seek(length - 2, IO::SEEK_CUR)
            end
          end

          nil
        rescue StandardError
          nil
        end
    end
  end
end
