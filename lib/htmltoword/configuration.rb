module Htmltoword
  class Configuration
    attr_accessor :default_templates_path, :custom_templates_path, :default_xslt_path, :custom_xslt_path

    # Callable invoked with the src of every remote image before it is downloaded.
    # Return nil to keep the default download, or a Hash with the image bytes in :data and,
    # optionally, the file extension in :ext (defaults to the extension inferred from src).
    attr_accessor :image_resolver

    def initialize
      @default_templates_path = File.join(File.expand_path('../', __FILE__), 'templates')
      @custom_templates_path = File.join(File.expand_path('../', __FILE__), 'templates')
      @default_xslt_path = File.join(File.expand_path('../', __FILE__), 'xslt')
      @custom_xslt_path = File.join(File.expand_path('../', __FILE__), 'xslt')
      @image_resolver = nil
    end
  end
end
