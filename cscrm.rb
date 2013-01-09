#!/usr/bin/env ruby

require 'rubygems'
require 'win32ole'
require 'ap'
require 'pp'
require 'net/http'
require 'xmlsimple'

if nil
$uri = "http://abt-cscrm.de.bosch.com/cqweb/restapi/CSCRM/CSCRM/QUERY/Personal%20Queries/B_BSW_AR40.7.0.0%20Req?format=XML&loginId=pbt2kor&password=as&noframes=true"
uri = URI($uri)
response = Net::HTTP.get(uri) # => String

File.open("response.xml", "w+") do |aFile|
	aFile.syswrite(response)
   # ... process the file
end
end
file = File.open('response_ref.xml')
content = XmlSimple.xml_in(file)
ap content
