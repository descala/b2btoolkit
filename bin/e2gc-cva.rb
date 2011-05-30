#!/usr/bin/ruby

# Invinet XML Tools 
# Purpose: Ruby tool to create Genericode and CVA files from an OpenOffice Spreadsheet
#
# Input: OpenOffice file with one sheet per codelist and a special sheet specifying the context value associations
# Ouput: gc folder with genericode files and cva folder with cva file
#
# Usage: ./e2gc-cva <openoffice_file> <extension> <path> <pathcva>
#
# Where <openoffice_file> has a sheet per code list and an specific sheet to define the context value associations
#	    <extension> the identifier of the extension that uses this cva
#		<path> is the folder where genericode files will be placed
#		<pathcva> is the folder where CVA file will be placed
#
# Author: Oriol Bausà (2010) Invinet Sistemes



require 'rubygems'
version = ">= 0"
gem 'roo', version
require 'roo'
require 'fileutils'

include FileUtils

puts "Analyzing: #{ARGV[0]}"

def main(fitxer, extension, path, pathcva)

  oo = Openoffice.new(fitxer)
  gcl = nil
  c = nil
  vl = nil
  cva = nil
  valuelist = []
  last_transaction = ""
  if (extension== nil) then extension="" end 
  
  puts "\nCreating Genericode files from an OpenOffice spreadsheet...\n"
     
 if (path == nil) 
      chdir "../gc"   
    else
	  mkdir_p path
      chdir path
  end  
  
  0.upto(oo.sheets.length - 1) do |sheet|

    if (oo.sheets[sheet] != "CVA") then 


      vl = ValueList.new(oo.sheets[sheet])      
      valuelist << vl if vl
            
      codes = []  
      oo.default_sheet = oo.sheets[sheet]
      shortname = oo.cell(2,'A')
      version = oo.cell(2,'B')
      agency = oo.cell(2,'C')
      locationuri = oo.cell(2,'D')
      locale    = oo.cell(2,'E')

      
      gcl = Genericode.new(shortname, version, agency, locationuri, locale)    

      5.upto(oo.last_row) do |line|
        code = oo.cell(line,'A')
        value = oo.cell(line,'B')
        if (code.class == Float) then code = code.truncate end
        if code != nil then 
          c = CodeRow.new(code, value)
          gcl.add_code(c) if c
        end
      end
      puts "Creating Genericode file "+oo.default_sheet+".gc..."
      file = File.new(oo.default_sheet+".gc","w")
      file.puts gcl.to_xml
    end

  
    if (oo.sheets[sheet] == "CVA") then
      oo.default_sheet = oo.sheets[sheet]

      list_of_cva = []
      
      2.upto(oo.last_row) do |line|
        transaction = oo.cell(line,'A')
        id = oo.cell(line,'B')
        item = oo.cell(line,'C')
        scope = oo.cell(line,'D')
        values = oo.cell(line,'E')
        message = oo.cell(line,'F')
        severity = oo.cell(line,'G')
      
        if transaction != last_transaction then
          last_transaction = transaction
		  puts transaction
          cva = ValueListConstraints.new(transaction)
          list_of_cva << cva if cva
          valuelist.each { |vl| cva.add_valuelist(vl) }

        end
      
        ctx = Context.new(id, item, scope, values, severity, message)    
        cva.add_context(ctx) if ctx
        
      end

	  chdir ".."
		
	  if (path == nil) 
		chdir "cva"   
	  else
		mkdir_p pathcva
		chdir pathcva
	  end  
	  	  
      0.upto(list_of_cva.length() - 1) do |cvafile|
        
		
		puts "Creating CVA file for "+extension+list_of_cva[cvafile].id+"..."
        file = File.new(extension+"Codes"+list_of_cva[cvafile].id+".cva", "w")
        file.puts list_of_cva[cvafile].to_xml
      end
    end
  end
end



class ValueListConstraints
  attr_accessor :id, :valuelist, :contexts
  
  def initialize(id)
    @id = id
    @valuelist = []
    @contexts = []
  end
  
  def add_valuelist(vl)
    @valuelist << vl
  end
  
  def add_context(ctx)
    @contexts << ctx
  end
  
  def to_xml
    xml = Builder::XmlMarkup.new(:indent => 2)
    xml.instruct!
    xml.comment! "

        	CVA File #{@id}
        	Oriol Bausà 

    "
    xml.ValueListConstraints("xmlns" => "http://docs.oasis-open.org/codelist/ns/ContextValueAssociation/cd2-1.0/",
      "xmlns:sch" => "http://purl.oclc.org/dsdl/schematron",
      "xmlns:qdt" => "urn:oasis:names:specification:ubl:schema:xsd:QualifiedDatatypes-2",
      "xmlns:cct" => "urn:oasis:names:specification:ubl:schema:xsd:CoreComponentParameters-2",
      "xmlns:cbc" => "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
      "xmlns:cac" => "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
      "xmlns:udt" => "urn:un:unece:uncefact:data:draft:UnqualifiedDataTypesSchemaModule:2",
      "xmlns:stat" => "urn:oasis:names:specification:ubl:schema:xsd:DocumentStatusCode-1.0",
      :name => "#{@extension}Codes#{@id}",
      :version => "Version 0.3") {
        xml.ValueLists {
          valuelist.each { |vl| vl.to_xml(xml) }
        }
        xml.Contexts {
          contexts.each { |ctx| ctx.to_xml(xml) }
        }
      }
  end
end

class ValueList
  attr_accessor :id
  
  def initialize(id)
    @id = id
  end
  
  def to_xml(xml)
    xml.ValueList("xml:id" => "#{@id}", :uri => "../gc/#{@id}.gc")
  end
end

class Context
  attr_accessor :id, :item, :scope, :values, :severity, :message
  
  def initialize(id, item, scope, values, severity, message)
    @id = id
    @item = item
    @scope = scope
    @values = values
    @severity = severity
    @message = message
  end
  
  def to_xml(xml)
    if (scope == nil) then
        xml.Context(:item => "#{@item}", :values => "#{@values}", :mark => "#{@severity}") {
      xml.Message("[#{@id}]-#{@message}")
    }
    else
        xml.Context(:item => "#{@item}", :scope => "#{@scope}", :values => "#{@values}", :mark => "#{@severity}") {
    xml.Message("[#{@id}]-#{@message}")
    }
    end
  end
end

class CodeRow
  attr_accessor :id
  attr_accessor :code, :value
  
  def initialize(code, value)
    @code = code
    @value = value
  end
  
  def to_xml(xml)
    xml.Row {
      xml.Value(:ColumnRef => "code") {
        xml.SimpleValue("#{@code}")
      }  
      xml.Value(:ColumnRef => "name") {
        xml.SimpleValue("#{@value}")
      }  
    }
  end
end

class Genericode
  attr_accessor :id, :shortname, :version, :agency, :locationuri, :locale, :codes
  
  def initialize(shortname, version, agency, locationuri, locale)
    @id = id
    @shortname = shortname
    @version = version
    @agency = agency
    @locationuri = locationuri
    @locale = locale
    @codes = []
  end
  
  def add_code(gcr)
   @codes << gcr
  end
  
  def to_xml
    xml = Builder::XmlMarkup.new(:indent => 2)
    xml.instruct!
    xml.comment! "

        	Genericode File #{@shortname}
        	
        	Oriol Bausà

    "
    xml.CodeList("xmlns" => "http://docs.oasis-open.org/codelist/ns/genericode/1.0/",
                "xmlns:xsi" => "http://www.w3.org/2001/XMLSchema-instance",
                "xsi:schemaLocation" => "http://docs.oasis-open.org/codelist/ns/genericode/1.0/ ../xsd/genericode.xsd") {
    xml.Identification(:xmlns => "") {
      xml.ShortName("#{@shortname}")
      xml.Version("#{@version}")
      xml.CanonicalUri("#{@agency}")
      xml.CanonicalVersionUri("#{@agency}-#{@version}")
      xml.LocationUri("#{@locationuri}")
    }
    xml.ColumnSet(:xmlns => ""){
      xml.Column(:Id => "code", :Use =>"required") {
        xml.ShortName("Code")
        xml.Data(:Type => "normalizedString")
      }
      xml.Column(:Id => "name", :Use =>"optional") {
        xml.ShortName("Name")
        xml.Data(:Type => "string")
      }
      xml.Key(:Id => "codeKey") {
        xml.ShortName("CodeKey")
        xml.ColumnRef(:Ref => "code")
      }
    }
    xml.SimpleCodeList(:xmlns => "") {
      codes.each { |code| code.to_xml(xml) }
    }
  }
  end
end
  
main(ARGV[0],ARGV[1],ARGV[2],ARGV[3])
  
