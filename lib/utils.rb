#!/usr/bin/env ruby

require 'rubygems'
require 'ap'
require 'pp'

module Utils
	def Utils.printHelp
			print <<"EOF";                
		
		Usage
		<command> <options>
		Possible values for command
		calc : calculate the effort and update the excel sheet
		sync : currently not supported

		options : list of options with key value pairs seperated by comma
		ex : key1 = value , key2 = value
		possible options are
		file = <excel file to process> default : EffortPlanning.xls located in the same directory
EOF
	end

	def Utils.parseCommandArgs args
		#some constants
		valid_commands = ["calc","sync"]
		default_filename = "EffortPlanning.xls"
		
		args_hash = {}
		
		if args.size == 0
			printHelp
			exit
		end
		
		command = args[0]
		
		##Check for valid commands
		unless valid_commands.include? command
			print "Invalid command #{command}"
			printHelp
			exit
			if command == "sync"
				print "sync is currently not supported"
				exit
			end
		end
		args_hash["command"] = command
		
		#Remove the command argument and prepare a string from others
		args.shift
		args_str = args.join(" ")
		
		args_str.split(",").each do |option|
			touple = option.split("=")
			args_hash[touple[0].strip] = touple[1].strip
		end
		
		args_hash["file"] = default_filename unless args_hash["file"]
		file = args_hash["file"]
		unless File.file?(file) 
			print "The specified file #{file} doesnt exist"
			exit
		end

		args_hash
	end 
end