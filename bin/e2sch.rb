#!/usr/bin/ruby

# Invinet XML Tools  
# Purpose: Ruby tool to create Schematron abstract rules and syntax binding from Spreadsheet
# 
#
# Usage: ./e2sch <rules_file> <path> <extension> <codelist.sch>
#
# where <rules_file> is an spreadsheet file in Open Office format with the following sheets:
#     abstract: where abstract rules are defined
#     <syntax>: where define the XPATH expression for rules and contexts
#     artifacts: to define which artifacts should be build
# 
#     <path> is used to define the place where to create the Schematron structure
#     <extension> the identifier for the extension used to prefix the files
#	  <codelist.sch> is the name of the codelist to be applied with the set of rules
#
# Author: Oriol Bausà (2010) Invinet Sistemes


require 'rubygems'
version = ">= 0"
gem 'roo', version
require 'roo'
require 'fileutils'
require_gem 'builder', '~> 2.0'

include FileUtils



puts "Analyzing: #{ARGV[0]}"

def main(fitxer,path,extension,codelist)
  
  oo = Openoffice.new(fitxer)
 # oo = Excel.new(fitxer)

  oo.default_sheet = oo.sheets.first

  if (path == nil) 
      chdir "../schematron" 
    else
	  mkdir_p path
      chdir path
  end
  
  if (codelist==nil) then codelist="" end
  
  mkdir_p oo.default_sheet
  chdir oo.default_sheet
  
  puts "\nCreating "+oo.default_sheet+" "+(oo.last_row-1).to_s+" rules...\n"

  patterns = []
  sheet_list = oo.sheets()
  
  if sheet_list == "artifacts" then puts sheet_list.to_s end
  
  puts sheet_list[1]
  
  p = nil
  r = nil
  a = nil

  last_transaction = ""
  last_context = ""
  new_transaction = false
  if extension==nil then extension="" end
  
  2.upto(oo.last_row) do |line|
    id          = oo.cell(line,'A')
    assertion   = oo.cell(line,'A')
    context     = oo.cell(line,'C')
    severity    = oo.cell(line,'D')
    message     = oo.cell(line,'B')
    transaction = oo.cell(line,'E')
	
	if transaction != nil
		if transaction != last_transaction
			last_transaction = transaction
			new_transaction = true
			p = Pattern.new(transaction)
			patterns << p if p
			r = nil
		end
		if context != last_context or new_transaction == true
			last_context = context
			new_transaction = false
			r = Rule.new(context)
			p.add_rule(r) if r
			a = nil
		end
		a = Assert.new(assertion.to_s,message,severity)
		r.add_assert(a) if a 
	end
  end
  patterns.each do |p|
    file = File.new(extension+"-"+p.id+".sch", "w")
    puts "New pattern "+extension+"-"+p.id+".sch"
    file.puts p.to_xml
  end
  
  # Syntax binding
  
 
  puts "\nIdentified "+(oo.sheets.length-2).to_s+" different syntax bindings"

  1.upto(oo.sheets.length - 2) do |sheet|
	chdir ".."
	oo.default_sheet = oo.sheets[sheet]
	
	mkdir_p oo.default_sheet
	chdir oo.default_sheet

	puts "\nCreating "+oo.default_sheet+" binding...\n"
  
	patterns_binding = []
	p = nil
	r = nil
	last_transaction = ""
	last_param = ""

	2.upto(oo.last_row) do |line|
		transaction = oo.cell(line,'A')
		param       = oo.cell(line,'B')
		value       = oo.cell(line,'C')
		prerequisite= oo.cell(line,'D')
     
		if transaction != last_transaction
			last_transaction = transaction
			p = Pattern_binding.new(transaction, oo.default_sheet)
			puts transaction
			patterns_binding << p if p
			r = nil
		end
		if param != last_param
			last_param = param
			if prerequisite != nil then value = value+' and '+prerequisite+' or not('+prerequisite+')' end 
			r = Param.new(param.to_s, value)
			p.add_param(r) if r
			a = nil
		end
   end
   
   
   patterns_binding.each do |p|
     file = File.new(extension+"-"+oo.default_sheet+"-"+p.id+".sch", "w")
     puts "New pattern "+extension+"-"+oo.default_sheet+"-"+p.id+".sch"
     file.puts p.to_xml
   end
  end
  
  # Create Schematron bundles defined in artifacts sheet
  
  puts "\nCreating Schematron bundles...\n"
  artifact_sheet="artifacts"
  
  
   chdir ".."
   
   last_sheet=oo.sheets.length

   oo.default_sheet = oo.sheets[last_sheet-1]
   
   2.upto(oo.last_row) do |line|
      profile     = oo.cell(line,'A')
      transaction = oo.cell(line,'B')
      binding     = oo.cell(line,'C')
	  namespace   = oo.cell(line,'D')
	  namespacecac   = oo.cell(line,'E')
	  namespacecbc   = oo.cell(line,'F')
	 
	 if (namespacecac == nil) then namespacecac = 'urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2' end 
	 if (namespacecbc == nil) then namespacecbc = 'urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2' end 
	 if (profile == nil) then profile='' end
     puts "Creating Assembly "+extension+"-"+binding+"-"+transaction+".sch..."
     sch = Schematron.new(binding,transaction,profile,codelist, extension,namespace,namespacecac,namespacecbc)
     file = File.new(extension+"-"+binding+"-"+transaction+".sch","w")
     file.puts sch.to_xml
	end
end


class Pattern
  attr_accessor :id, :rules
  
  def initialize(id)
    @id=id
    @rules =[]
    @abstract = true
  end
  
  def add_rule(r)
    @rules << r
  end

  def to_xml
    
    xml = Builder::XmlMarkup.new(:indent => 2)
    xml.comment! "Schematron rules generated automatically."
    xml.comment! "Abstract rules for #{@id}"
    xml.comment! "(2009). Invinet Sistemes"
    xml.pattern(:xmlns => "http://purl.oclc.org/dsdl/schematron",:abstract => "true",:id => @id) { 
       rules.each { |rule| rule.to_xml(xml) } 
    }    
  end
end

class Rule
  attr_accessor :id, :context, :asserts
  def initialize(context)
    @id=context
    @context=context
    @asserts =[]
  end
  
  def add_assert(r)
    @asserts << r
  end
  
  def to_xml(xml)
    @context = @context.gsub(/\b[ \t]+\b/, '_')
    @context = @context.gsub(/[ \t]+\b/, ' $')
    @context = @context.gsub(/^\b/, '$')
    xml.rule(:context => "#{@context}") {
      asserts.each { |assert| assert.to_xml(xml) }
    }
  end
end

class Assert
  attr_accessor :id
  attr_accessor :test, :message, :severity
  def initialize(test, message, severity)
    @test = test
    @testvar = ''
    @message = message
    @severity = severity
  end
  def to_xml(xml)
    @test = @test.gsub(/\b[ \t]+\b/, '_')
    @test = @test.gsub(/[ \t]+\b([^0-9])/, ' $\1')
    @testvar = @test.gsub(/^\b/, '$')
    xml.assert({:test => "#{@testvar}", :flag => "#{@severity}"}, "[#{@test}]-#{@message}")
  end
end

class Pattern_binding
  attr_accessor :id, :params, :binding
  
  def initialize(id,binding)
    @id=id
    @params=[]
	@binding=binding
  end
  
  def add_param(p)
    @params << p
  end

  def to_xml
    xml = Builder::XmlMarkup.new(:indent => 2)
    xml.comment! "Schematron binding rules generated automatically."
    xml.comment! "Data binding to #{@binding.upcase} syntax for #{@id}"
    xml.comment! "(2009). Invinet Sistemes"
    xml.pattern(:xmlns => "http://purl.oclc.org/dsdl/schematron", "is-a" => @id ,:id => "#{@binding.upcase}-#{@id}") { 
       params.each { |param| param.to_xml(xml) } 
    }    
  end
end

class Param
  attr_accessor :id
  attr_accessor :param, :value
  def initialize(param, value)
    @param = param
    @value = value
  end
  
  def to_xml(xml)
    @param = @param.gsub(/\b[ \t]+\b/, '_')
    xml.param(:name=>"#{@param}", :value=>"#{@value}")
  end

end

class Schematron
  attr_accessor :id, :binding, :transaction, :profile, :codelist, :extension, :namespace, :namespacecac, :namespacecbc
  
  def initialize(binding, transaction, profile,codelist,extension,namespace,namespacecac,namespacecbc)
    @binding = binding
    @transaction = transaction
    @profile = profile
	@codelist = codelist
	@extension = extension
	@namespace = namespace
	@namespacecac = namespacecac
	@namespacecbc = namespacecbc
	
  end

  def to_xml
    xml = Builder::XmlMarkup.new(:indent => 2)
    xml.instruct!
    xml.comment! "

        	#{@binding.upcase} syntax binding to the #{@transaction}  #{@profile} 
        	Author: Oriol Bausà

    "
    xml.schema("xmlns" => "http://purl.oclc.org/dsdl/schematron", 
                "xmlns:cbc" => "#{@namespacecbc}",
                "xmlns:cac" => "#{@namespacecac}",
                "xmlns:ubl" => "#{@namespace}",
                :queryBinding => "xslt2") {
                  xml.title(extension.upcase+" #{@profile} #{@transaction} bound to "+binding.upcase)
                  xml.ns(:prefix => "cbc", :uri => "#{@namespacecbc}")
                  xml.ns(:prefix => "cac", :uri => "#{@namespacecac}")
                  xml.ns(:prefix => "#{@binding.downcase}", :uri => "#{@namespace}")
                  
				 xml.phase(:id => "#{@extension}#{@transaction}_phase") {
                    xml.active(:pattern => "#{@binding.upcase}-#{@transaction}")
                  }
				  
				  if @profile.length()>0 then
                  xml.phase(:id => "#{@extension}#{@profile}_phase") {
                    xml.active(:pattern => "#{@binding.upcase}-#{@profile}")
                  }      
				  end
				  
				  if @codelist.length()>0 then
					xml.phase(:id => "codelist_phase") {
						xml.active(:pattern => "Codes#{@transaction}")
					}
                  end
				  
                  xml.comment! "Abstract CEN BII patterns"
                  xml.comment! "========================="
                  
                  xml.include(:href => "abstract/#{@extension}-#{@transaction}.sch") 
				  
                  if @profile.length()>0 then
					xml.include(:href => "abstract/#{@extension}#{@profile}.sch") 
                  end
                  
				  xml.comment! "Data Binding parameters"
                  xml.comment! "======================="
                  
                  xml.include(:href => "#{@binding}/#{@extension}-#{@binding}-#{@transaction}.sch") 
                  if @profile.length()>0 then
					xml.include(:href => "#{@binding}/#{@extension}#{@profile}-#{@binding}.sch") 
			      end
				  if @codelist.length()>0 then
					xml.comment! "Code Lists Binding rules"
					xml.comment! "========================"
					xml.include(:href => "codelist/#{@codelist}")
				  end	
                }
  end
end

main(ARGV[0],ARGV[1],ARGV[2],ARGV[3])


