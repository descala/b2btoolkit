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
# Author: Oriol Bausà (2011) Invinet Sistemes



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

    if (oo.sheets[sheet] != "CVA") and (oo.sheets[sheet] != "Index") then 


      vl = ValueList.new(oo.sheets[sheet])      
      valuelist << vl if vl
            
      codes = []  
      oo.default_sheet = oo.sheets[sheet]
	  listid = oo.cell(2,'A')
      shortname = oo.cell(2,'B')
      version = oo.cell(2,'C')
      agency = oo.cell(2,'D')
      locationuri = oo.cell(2,'E')
      locale    = oo.cell(2,'F')

      
      gcl = Genericode.new(listid, shortname, version, agency, locationuri, locale)    

      5.upto(oo.last_row) do |line|
        code = oo.cell(line,'A')
        value = oo.cell(line,'B')
		valor = oo.cell(line,'C')
        if (code.class == Float) then code = code.truncate end
        if code != nil then 
          c = CodeRow.new(code, value, valor)
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
        values = oo.cell(line,'D')
        message = oo.cell(line,'E')
        severity = oo.cell(line,'F')
		metadata = oo.cell(line,'G')
		
       if transaction != last_transaction then
          last_transaction = transaction
		  puts transaction
          cva = ValueListConstraints.new(transaction)
          list_of_cva << cva if cva
          valuelist.each { |vl| cva.add_valuelist(vl) }

        end
      
        ctx = Context.new(id, item, values, severity, message, metadata)    
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
    xml.ContextValueAssociation("xmlns" => "http://docs.oasis-open.org/codelist/ns/ContextValueAssociation/1.0/",
      "xmlns:sch" => "http://purl.oclc.org/dsdl/schematron",
      "xmlns:cbc" => "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
      "xmlns:cac" => "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
	   :name => "#{@extension}Codes#{@id}",
      :version => "1.0") {
        xml.ValueLists {
          valuelist.each { |vl| vl.to_xml(xml) }
        }
		xml.InstanceMetadataSets {
		  xml.InstanceMetadataSet("xml.id" => "cctsV2.01-amount")  {
			xml.InstanceMetadata(:address => "../@currencyCodeListVersionID",:identification => "Version" )
		  }
		  xml.InstanceMetadataSet("xml.id" => "cctsV2.01-measure")  {
			xml.InstanceMetadata(:address => "../@unitCodeListVersionID",:identification => "Version" )
		  }
		  xml.InstanceMetadataSet("xml.id" => "cctsV2.01-quantity")  {
			xml.InstanceMetadata(:address => "../@unitCodeListID",:identification => "Version" )
			xml.InstanceMetadata(:address => "../@unitCodeListAgencyName",:identification => "Agency/LongName" )
			xml.InstanceMetadata(:address => "../@unitCodeListAgencyID",:identification => "Agency/Identifier" )
		  }
		  xml.InstanceMetadataSet("xml.id" => "cctsV2.01-code")  {
			xml.InstanceMetadata(:address => "@listName",:identification => "LongName[not(@Identifier='listID')]" )
			xml.InstanceMetadata(:address => "@listID",:identification => "LongName[@Identifier='listID']" )
			xml.InstanceMetadata(:address => "@listVersionID",:identification => "Version" )
			xml.InstanceMetadata(:address => "@listSchemeURI",:identification => "CanonicalUri" )
			xml.InstanceMetadata(:address => "@listURI",:identification => "LocationUri" )
			xml.InstanceMetadata(:address => "@listAgencyName",:identification => "Agency/LongName" )
			xml.InstanceMetadata(:address => "@listAgencyID",:identification => "Agency/Identifier" )
		  }
		  xml.InstanceMetadataSet("xml.id" => "cctsV2.01-identifier")  {
			xml.InstanceMetadata(:address => "@schemeName",:identification => "LongName" )
			xml.InstanceMetadata(:address => "@schemeVersionID",:identification => "Version" )
			xml.InstanceMetadata(:address => "@schemeURI",:identification => "CanonicalUri" )
			xml.InstanceMetadata(:address => "@schemeDataURI",:identification => "LocationUri" )
			xml.InstanceMetadata(:address => "@schemeAgencyName",:identification => "Agency/LongName" )
			xml.InstanceMetadata(:address => "@schemeAgencyID",:identification => "Agency/Identifier" )		  }
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
  attr_accessor :id, :item, :values, :severity, :message, :metadata
  
  def initialize(id, item, values, severity, message, metadata)
    @id = id
    @item = item
    @values = values
    @severity = severity
    @message = message
	@metadata = metadata
  end
  
  def to_xml(xml)
    if (metadata == nil) then
        xml.Context(:address => "#{@item}", :values => "#{@values}", :mark => "#{@severity}") {
      xml.Message("[#{@id}]-#{@message}")
    }
    else
	xml.Context(:address => "#{@item}", :values => "#{@values}", :mark => "#{@severity}", :metadata => "#{@metadata}") {
		xml.Message("[#{@id}]-#{@message}")
		}
	end
  end
end

class CodeRow
  attr_accessor :id
  attr_accessor :code, :value, :valor
  
  def initialize(code, value, valor)
    @code = code
    @value = value
	@valor = valor
  end
  
  def to_xml(xml)
    xml.Row {
      xml.Value(:ColumnRef => "code") {
        xml.SimpleValue("#{@code}")
      }  
      xml.Value(:ColumnRef => "name") {
        xml.SimpleValue("#{@value}")
      }  
	  if (valor != nil) then 
	  xml.Value(:ColumnRef => "nombre") {
			xml.SimpleValue("#{@valor}")
	  }
	  end
	  }
  end
end

class Genericode
  attr_accessor :listid, :shortname, :version, :agency, :locationuri, :locale, :codes
  
  def initialize(listid,shortname, version, agency, locationuri, locale)
    @listid = listid
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
    xml.CodeList("xmlns" => "http://docs.oasis-open.org/codelist/ns/genericode/1.0/") {
    xml.Identification {
	  xml.LongName("#{@listid}", :Identifier => "listID") 
	  xml.ShortName("#{@shortname}")
      xml.Version("#{@version}")
    }
    xml.ColumnSet {
      xml.Column(:Id => "code", :Use =>"required") {
        xml.ShortName("Code")
        xml.Data(:Type => "normalizedString")
      }
      xml.Column(:Id => "name", :Use =>"optional") {
        xml.ShortName("Name")
        xml.Data(:Type => "string")
      }
      xml.Column(:Id => "nombre", :Use =>"optional") {
        xml.ShortName("Nombre")
        xml.Data(:Type => "string")
      }
	  xml.Key(:Id => "codeKey") {
        xml.ShortName("CodeKey")
        xml.ColumnRef(:Ref => "code")
      }
    }
    xml.SimpleCodeList {
      codes.each { |code| code.to_xml(xml) }
    }
  }
  end
end
  
main(ARGV[0],ARGV[1],ARGV[2],ARGV[3])
  
