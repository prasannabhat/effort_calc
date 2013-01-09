#!/usr/bin/env ruby

require 'rubygems'
require 'win32ole'
require 'ap'
require 'pp'
require 'net/http'
require 'lib/utils'

a = [1,2,3,4]
b = a.index  5

columnNameMap = {
	"CSCRM" => "taskIdList.id",
	"Effort" => "taskIdList.effortPlanned",
	"Module" => "taskIdList.swComponent" ,
	"Developer" => "taskIdList.Owner.login_name",
	"MonthlyBaseline" => "targetBaselineId" 
	
}
ap Utils.parseCommandArgs ARGV

