require 'docx/containers'
require 'docx/elements'
require 'nokogiri'
require 'zip'

module Docx
  # The Document class wraps around a docx file and provides methods to
  # interface with it.
  #
  #   # get a Docx::Document for a docx file in the local directory
  #   doc = Docx::Document.open("test.docx")
  #
  #   # get the text from the document
  #   puts doc.text
  #
  #   # do the same thing in a block
  #   Docx::Document.open("test.docx") do |d|
  #     puts d.text
  #   end
  class Document
    attr_reader :xml, :doc, :styles, :header_and_footers_xml, :charts_xml, :diagrams_xml
    
    def initialize(path, &block)
      @replace = {}
      zip = Zip::File.open(path)
      document_xml = zip.read('word/document.xml')

      content_types_xml = zip.read('[Content_Types].xml')
      content_types = Nokogiri::XML(content_types_xml)
      
      header_and_footers = []
      content_types.css('Override').each do |override_node|
        if override_node['PartName'].include?("header") || override_node['PartName'].include?("footer")
          header_and_footers << override_node['PartName'][1..-1]
        end
      end
      @header_and_footers_xml = []
      header_and_footers.each do |elem|
        if zip.find_entry(elem)
          temp = Nokogiri::XML(zip.read(elem))
          temp.root['xmlns:wp14'] = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
          temp.root['xmlns:wp'] = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
          temp.root['xmlns:w10'] = "urn:schemas-microsoft-com:office:word"
          temp.root['xmlns:w'] = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          temp.root['xmlns:w14'] = "http://schemas.microsoft.com/office/word/2010/wordml"
          temp.root['xmlns:w15'] = "http://schemas.microsoft.com/office/word/2012/wordml"
          temp.root['xmlns:w16cid'] = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
          temp.root['xmlns:w16se'] = "http://schemas.microsoft.com/office/word/2015/wordml/symex"
          temp.root['mc:Ignorable'] = "w14 w15 w16se w16cid wp14"
          temp.root['xmlns:mc'] = "http://schemas.openxmlformats.org/markup-compatibility/2006"
          temp.root['xmlns:wps'] = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"              
          temp.root['xmlns:r'] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
          temp.root['xmlns:m'] = "http://schemas.openxmlformats.org/officeDocument/2006/math"
          temp.root['xmlns:v'] = "urn:schemas-microsoft-com:vml"          
          
          @header_and_footers_xml << Nokogiri::XML(temp.to_xml(:save_with => Nokogiri::XML::Node::SaveOptions::AS_XML))
        end
      end
      
      temp = Nokogiri::XML(document_xml)
      temp.root['xmlns:wp14'] = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
      temp.root['xmlns:wp'] = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
      temp.root['xmlns:w10'] = "urn:schemas-microsoft-com:office:word"
      temp.root['xmlns:w'] = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      temp.root['xmlns:w14'] = "http://schemas.microsoft.com/office/word/2010/wordml"
      temp.root['xmlns:w15'] = "http://schemas.microsoft.com/office/word/2012/wordml"
      temp.root['xmlns:w16cid'] = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
      temp.root['xmlns:w16se'] = "http://schemas.microsoft.com/office/word/2015/wordml/symex"
      temp.root['mc:Ignorable'] = "w14 w15 w16se w16cid wp14"
      temp.root['xmlns:mc'] = "http://schemas.openxmlformats.org/markup-compatibility/2006"      
      temp.root['xmlns:wps'] = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
      temp.root['xmlns:r'] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
      temp.root['xmlns:m'] = "http://schemas.openxmlformats.org/officeDocument/2006/math"
      temp.root['xmlns:v'] = "urn:schemas-microsoft-com:vml"         
      @doc = Nokogiri::XML(temp.to_xml(:save_with => Nokogiri::XML::Node::SaveOptions::AS_XML))
      
      styles_xml = zip.read('word/styles.xml')
      @styles = Nokogiri::XML(styles_xml)
      if block_given?
        yield self
        zip.close
      end

      charts = []
      content_types.css('Override').each do |override_node|
        if override_node['PartName'].include?("charts/ch")
          charts << override_node['PartName'][1..-1]
        end
      end
      @charts_xml = []
      charts.each do |elem|
        if zip.find_entry(elem)          
          temp = Nokogiri::XML(zip.read(elem))
          temp.root['xmlns:wp14'] = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
          temp.root['xmlns:wp'] = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
          temp.root['xmlns:w10'] = "urn:schemas-microsoft-com:office:word"
          temp.root['xmlns:w'] = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          temp.root['xmlns:w14'] = "http://schemas.microsoft.com/office/word/2010/wordml"
          temp.root['xmlns:w15'] = "http://schemas.microsoft.com/office/word/2012/wordml"
          temp.root['xmlns:w16cid'] = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
          temp.root['xmlns:w16se'] = "http://schemas.microsoft.com/office/word/2015/wordml/symex"
          temp.root['mc:Ignorable'] = "w14 w15 w16se w16cid wp14"
          temp.root['xmlns:mc'] = "http://schemas.openxmlformats.org/markup-compatibility/2006"
          temp.root['xmlns:wps'] = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"              
          temp.root['xmlns:r'] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
          temp.root['xmlns:m'] = "http://schemas.openxmlformats.org/officeDocument/2006/math"
          temp.root['xmlns:v'] = "urn:schemas-microsoft-com:vml"
          temp.root['xmlns:c'] = "urn:schemas-microsoft-com:c"
          temp.root['xmlns:a'] = "urn:schemas-microsoft-com:a"
          @charts_xml << Nokogiri::XML(temp.to_xml(:save_with => Nokogiri::XML::Node::SaveOptions::AS_XML))          
        end
      end

      diagrams = []
      content_types.css('Override').each do |override_node|
        if override_node['PartName'].include?("diagrams")
          diagrams << override_node['PartName'][1..-1]
        end
      end
      @diagrams_xml = []
      diagrams.each do |elem|
        if zip.find_entry(elem)          
          temp = Nokogiri::XML(zip.read(elem))
          temp.root['xmlns:wp14'] = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
          temp.root['xmlns:wp'] = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
          temp.root['xmlns:w10'] = "urn:schemas-microsoft-com:office:word"
          temp.root['xmlns:w'] = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          temp.root['xmlns:w14'] = "http://schemas.microsoft.com/office/word/2010/wordml"
          temp.root['xmlns:w15'] = "http://schemas.microsoft.com/office/word/2012/wordml"
          temp.root['xmlns:w16cid'] = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
          temp.root['xmlns:w16se'] = "http://schemas.microsoft.com/office/word/2015/wordml/symex"
          temp.root['mc:Ignorable'] = "w14 w15 w16se w16cid wp14"
          temp.root['xmlns:mc'] = "http://schemas.openxmlformats.org/markup-compatibility/2006"
          temp.root['xmlns:wps'] = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"              
          temp.root['xmlns:r'] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
          temp.root['xmlns:m'] = "http://schemas.openxmlformats.org/officeDocument/2006/math"
          temp.root['xmlns:v'] = "urn:schemas-microsoft-com:vml"
          temp.root['xmlns:c'] = "urn:schemas-microsoft-com:c"
          temp.root['xmlns:a'] = "urn:schemas-microsoft-com:a"
          @diagrams_xml << Nokogiri::XML(temp.to_xml(:save_with => Nokogiri::XML::Node::SaveOptions::AS_XML))   
        end
      end

    end



    # This stores the current global document properties, for now
    def document_properties
      {
        font_size: font_size
      }
    end


    # With no associated block, Docx::Document.open is a synonym for Docx::Document.new. If the optional code block is given, it will be passed the opened +docx+ file as an argument and the Docx::Document oject will automatically be closed when the block terminates. The values of the block will be returned from Docx::Document.open.
    # call-seq:
    #   open(filepath) => file
    #   open(filepath) {|file| block } => obj
    def self.open(path, &block)
      self.new(path, &block)
    end

    def paragraphs
      @doc.xpath('//w:document//w:body//w:p').map { |p_node| parse_paragraph_from p_node }
    end

    def bookmarks
      bkmrks_hsh = Hash.new
      bkmrks_ary = @doc.xpath('//w:bookmarkStart').map { |b_node| parse_bookmark_from b_node }
      # auto-generated by office 2010
      bkmrks_ary.reject! {|b| b.name == "_GoBack" }
      bkmrks_ary.each {|b| bkmrks_hsh[b.name] = b }
      bkmrks_hsh
    end

    def tables
      @doc.xpath('//w:document//w:body//w:tbl').map { |t_node| parse_table_from t_node }
    end

    # Some documents have this set, others don't.
    # Values are returned as half-points, so to get points, that's why it's divided by 2.
    def font_size
      size_tag = @styles.xpath('//w:docDefaults//w:rPrDefault//w:rPr//w:sz').first
      size_tag ? size_tag.attributes['val'].value.to_i / 2 : nil
    end

    ##
    # *Deprecated*
    #
    # Iterates over paragraphs within document
    # call-seq:
    #   each_paragraph => Enumerator
    def each_paragraph
      paragraphs.each { |p| yield(p) }
    end

    # call-seq:
    #   to_s -> string
    def to_s
      paragraphs.map(&:to_s).join("\n")
    end

    # Output entire document as a String HTML fragment
    def to_html
      paragraphs.map(&:to_html).join('\n')
    end

    # Save document to provided path
    # call-seq:
    #   save(filepath) => void
    def save(path)
      update
      Zip::OutputStream.open(path) do |out|
        zip.each do |entry|
          out.put_next_entry(entry.name)

          if @replace[entry.name]
            out.write(@replace[entry.name])
          else
            out.write(zip.read(entry.name))
          end
        end
      end
      zip.close
    end

    def save_and_return
      update
      stringio = Zip::OutputStream.write_buffer do |out|
        zip.each do |entry|
          out.put_next_entry(entry.name)

          if @replace[entry.name]
            out.write(@replace[entry.name])
          else
            out.write(zip.read(entry.name))
          end
        end
      end
      zip.close
      return stringio
    end

    alias_method :text, :to_s

    def replace_entry(entry_path, file_contents)
      @replace[entry_path] = file_contents
    end

    private

    #--
    # TODO: Flesh this out to be compatible with other files
    # TODO: Method to set flag on files that have been edited, probably by inserting something at the
    # end of methods that make edits?
    #++
    def update
      replace_entry "word/document.xml", doc.serialize(:save_with => 0)

      @header_and_footers.each_with_index do |header_and_footer, index|
        replace_entry header_and_footer, header_and_footers_xml[index].serialize(:save_with => 0) if header_and_footers_xml[index]  
      end

      @charts.each_with_index do |chart, index|
        replace_entry chart, charts_xml[index].serialize(:save_with => 0) if charts_xml[index]  
      end

      @diagrams.each_with_index do |diagram, index|
        replace_entry diagram, diagrams_xml[index].serialize(:save_with => 0) if diagrams_xml[index]  
      end  
    end

    # generate Elements::Containers::Paragraph from paragraph XML node
    def parse_paragraph_from(p_node)
      Elements::Containers::Paragraph.new(p_node, document_properties)
    end

    # generate Elements::Bookmark from bookmark XML node
    def parse_bookmark_from(b_node)
      Elements::Bookmark.new(b_node)
    end

    def parse_table_from(t_node)
      Elements::Containers::Table.new(t_node)
    end
  end
end
