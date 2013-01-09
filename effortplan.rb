#!/usr/bin/env ruby

require 'rubygems'
require 'win32ole'
require 'ap'
require 'pp'
require 'lib/utils'


#Productivity values of developers, wrt modules
#if no value is given, then it is considered as 1
$sridhar = {"Nm" => 1, "CanNm" => 1, "FrNm" => 1}
$david = {"Nm" => 1, "CanNm" => 1, "FrNm" => 1}
$sameer = {"Com" => 1}
$venky = {"Xcp" => 1}
$sachin = {"CanNm" => 1.5}
$lekha = {"CanNm" => 2, "Nm" => 2, "FrNm" => 2}
$vijay = {"CanNm" => 2, "Nm" => 2, "FrNm" => 2}
$krishna = {}
$productivity = {"sridhar" => $sridhar, "david" => $david, "sameer" => $sameer, "venky" => $venky, "sachin" => $sachin, "lekha" => $lekha, "krishna" => $krishna, "vijay" =>$vijay}

$wb
$excel
#Default name for the excel file to process, if specified then read from the command line
$excel_name
$excel_path
$ws1
$ws2
$ws_nonCSCRM
$args

#mapping for column names
$columnNameMap = {
	"CSCRM" => "taskIdList.id",
	"Effort" => "taskIdList.effortPlanned",
	"Module" => "taskIdList.swComponent" ,
	"Developer" => "taskIdList.Owner.login_name",
	"MonthlyBaseline" => "targetBaselineId" ,
	"CSCRM_Req" => "id"
	
}
#Key column numbers in the first sheet (it will be processed and updated)
$columnNumbers1 = {"CSCRM" => -1,"Effort" => -1,"Module" => -1,"Acceptence" => -1,"Developer" => -1,"Normalised Effort" => -1,"TaskProductivity" => -1,"MonthlyBaseline" => -1,
					"CSCRM_Req" => -1}

#Key column numbers in the second sheet
$columnNumbers2 = {"Developer" => -1, "AssignedEffort" => -1, "Misc Tasks" => -1}
$columnNumbers_nonCSCRM = {"Developer" => -1, "Effort(hrs)" => -1, "Task Name" => -1}
$validRow = 0

#get the max valid row number, which contains CSCRM entries
def getValidRowNumber
  cell_value = "CSCRM"
	row = 1
	while  /CSCRM/i =~ cell_value
		row = row + 1
		cell_value = $ws1.cells(row,$columnNumbers1["CSCRM_Req"]).value
	end
	$validRow = row - 1
end

#get the max no of valid rows , based on the column (valid means which contains some entry)
def getMaxRow ws,columnNo
	row = 1
	while cell_value = ws.cells(row,columnNo).value
		row = row + 1
	end
	row = row - 1
end

#Read some column indexs from the workbook and store it
def getColumnIndex map,ws
	#range gives a two dimensional array
	headLines = ws.range("a1:o1").value[0]
	map.each do |key,val|
		#Check if the given column itself exists in the workbook (Ex : CSCRM)
		index = headLines.index key
		if index
			map[key] = index + 1
		else
			key_alias = $columnNameMap[key]
			if key_alias
				#if the column name alias exists, check for the alias name exists in the worksheet
				index = headLines.index key_alias
				map[key] = index + 1 if index
			end
		end
	end
end

def copyColumn src,dest,rowMax
	src_wb = src["wb"];
	dest_wb = dest["wb"];
	rowMax.times do |row|
		row = row + 2
		dest_wb.cells(row,dest["col"]).value = src_wb.cells(row,src["col"]).value
	end
end

def initialise
	$excel_name = $excel_path.split('/')[-1]
	
	$excel.workbooks.each do |wb|
		if wb.name == $excel_name
			$wb = wb
			puts "Workbook #{wb.name} already open, connecting to it"
			break
		end
	end
	unless $wb
		#Open the required workbook 
		$wb = $excel.Workbooks.Open($excel_path)
		puts "Opening file #{$excel_name}"
	end
	
	#get hold of the first worksheet
	$ws1 = $wb.Worksheets("Tasks")
	#bring it to the front -need sometimes to run macros, 
	#not for working with a worksheet from ruby
	#$ws1.Select
	$ws2 = $wb.Worksheets("Effort")
	#Sheet containing non CSCRM activities
	$ws_nonCSCRM = $wb.Worksheets("nonCSCRMActivities")
	
	$excel.Visible = true
	#To make Excel visible, but block user input, set the Application object's Interactive property:
	$excel.Interactive = false
	#To turn off screen updating, set the ScreenUpdating object's Interactive property:
	$excel.ScreenUpdating = false
	
	#Locate the column numbers of key columns in worksheet1
	getColumnIndex $columnNumbers1,$ws1
	#Locate the column numbers of key columns in worksheet2
	getColumnIndex $columnNumbers2,$ws2
	getColumnIndex $columnNumbers_nonCSCRM, $ws_nonCSCRM
	
	#get the rows, untill which CSCRMs are present
	getValidRowNumber
	
	#First copy the effort to normalised effort first
	src = {"wb" => $ws1, "col" => $columnNumbers1["Effort"]}
	dest = {"wb" => $ws1, "col" => $columnNumbers1["Normalised Effort"]}
	copyColumn src,dest , $validRow

end

#calculate the normalised effort based on the productivity
def calcNormEffort
	ws = $ws1
	rowMax = $validRow
	#do for all the CSCRMs
	rowMax.times do |row|
		row = row + 2
		#default factor to multiply
		factor = 1
		effort = ws.cells(row, $columnNumbers1["Effort"]).value
		developer = ws.cells(row, $columnNumbers1["Developer"]).value
		component = ws.cells(row, $columnNumbers1["Module"]).value
		
		#all of these should be defined
		if developer && component
			if $productivity.has_key?(developer) && $productivity[developer].has_key?(component)
				factor = $productivity[developer][component]
			end
		end
		
		#Task specific productivity overrides the regular productivity
		taskFactor = ws.cells(row, $columnNumbers1["TaskProductivity"]).value if $columnNumbers1["TaskProductivity"] != -1
		
		if taskFactor
			factor = taskFactor
		end
		
		if effort
			ws.cells(row,$columnNumbers1["Normalised Effort"]).value = effort * factor
		end
	end
end

#Read the column elements from given column, that meets the given criterion on other columns(from worksheet1 only)
# inputs : criterion => hash with key as column number and value as a regex for performing match
#          columnNumber => column number to read and match
#		   maxColumns => maximum number of columns to search for
#Returns an array

def readColumnsCrieteria ws,columnNumber,criterion,maxColumns
    returnVal = []
    maxColumns.times do |row|
        row = row + 2
        match = true
        criterion.each do |key,val|
            value = ws.cells(row,key).value
            unless criterion[key].match(value.to_s)
                match = false                
            end
        end
        #if the given crieteria matches, push the value to the array
        if match
            value = ws.cells(row,columnNumber).value
            returnVal.push value  
        end
        #check if the given criterion are matching
    end
    returnVal 
end

# matches the contents of the srcColumn with the keys of the srcHash.
# If a match is found, then updated the corresponding value from the hash in destination column
# writeNew : if this is true, after writing matching rows, we write the new elements of the srcHash as well
def updateColumns ws,srcColumn,srcHash,destColumn,writeNew = false
    row = 1
    keys = srcHash.keys
    #as long as the cell content is not nill
    while value = ws.cells(row,srcColumn).value
        #check if the contents of the srcColumn is present in the given hash
        if keys.include? value
            ws.cells(row,destColumn).value = srcHash[value]
			#Delete the entry , after writing into the excel
			srcHash.delete value
        end
		row = row + 1
    end 
	#check if new values should be written
	srcHash.clear unless writeNew

	#Write the remaining entries in the excel sheet
	srcHash.each do |key, value|
		#Since the key dint exist originally, write that as well
		ws.cells(row,srcColumn).value = key
		ws.cells(row,destColumn).value = srcHash[key]
		row = row + 1
	end
    
end

def getUniqueColumnItems ws,columnNo,startRow = 1, endRow = 100
	#Excel column array , alphabet representing the column
	column = ('A'..'ZZ').to_a[columnNo - 1]
	#The values will be returned as a two dimensional array
	values = ws.range("#{column}#{startRow}:#{column}#{endRow}").value
	#get the unique values
	values = values.map {|x| x[0] }.uniq {|x| x}
end

#Update the effort required by each developer
def getDeveloperTaskEffort
    effortHash = {}
	#Read in the developers from theexcel sheet
	developers = getUniqueColumnItems $ws1,$columnNumbers1["Developer"],2,$validRow
    developers.each do |name|
        criterion = {$columnNumbers1["Effort"] => /^\s*\d/i, $columnNumbers1["Developer"] => /^\s*#{name}/i, $columnNumbers1["Acceptence"] => /^\s*Yes/i}
        efforts = readColumnsCrieteria $ws1,$columnNumbers1["Normalised Effort"],criterion,$validRow
        #Store the result in the hash
		totalEffort = efforts.inject{|result,element| result + element} 
		#Default values of zero
		totalEffort = 0 unless totalEffort
        effortHash[name] = totalEffort
    end
    effortHash
end

#Update the effort required by each developer, for non CSCRM activities
def getDeveloperNonTaskEffort
    effortHash = {}
	maxRow = getMaxRow $ws_nonCSCRM, $columnNumbers_nonCSCRM["Task Name"]
	#Read in the developers from the excel sheet
	developers = getUniqueColumnItems $ws_nonCSCRM,$columnNumbers_nonCSCRM["Developer"],2,maxRow
    developers.each do |name|
		criterion = {$columnNumbers_nonCSCRM["Effort(hrs)"] => /^\s*\d/i, $columnNumbers_nonCSCRM["Developer"] => /^\s*#{name}/i}
        efforts = readColumnsCrieteria $ws_nonCSCRM,$columnNumbers_nonCSCRM["Effort(hrs)"],criterion,maxRow
        #Store the result in the hash
		totalEffort = efforts.inject{|result,element| result + element} 
		#Default values of zero
		totalEffort = 0 unless totalEffort
		#Convert to days
		totalEffort = (totalEffort / 8.5)
		#Round to 1 decimal position
		totalEffort = (totalEffort * 10).round() / 10.0
        effortHash[name] = totalEffort
    end
    effortHash
end

#Check for any errors in the excel sheet
def errorCheck

	def checkStatus
		
	end

	ws = $ws1
	row = 1
	column = $columnNumbers1["CSCRM"]
	#Do as long as there are entries in the CSCRM column
	while value = $ws1.cells(row,column).value
		#proceed if there is a valid CSCRM entry
		if /^\s*CSCRM/i =~ value
			status = ws.cells(row,$columnNumbers1["Acceptence"]).value
			status = status ? status.downcase : "no"
			if status == 'yes'
				developer = ws.cells(row,$columnNumbers1["Developer"]).value
				monthlyBaseline = developer = ws.cells(row,$columnNumbers1["MonthlyBaseline"]).value
				unless monthlyBaseline
					puts "Task #{value} status is Yes and  monthly baseline assigned to this task"
				end
			end
		end
		row = row + 1
	end

end


def test
#	cell = $wb.Worksheets(1).cells(1,1).value
	#cell = $wb.Worksheets("Tasks").cells(1,1).value
	#Excel column array
	b = ('A'..'ZZ').to_a
	#alphabet representing the column
	column = b[$columnNumbers1["Developer"] - 1]
	headLines = $ws1.range("#{column}2:#{column}#{$validRow}").value
	headLines = headLines.map {|x| x[0] }.uniq {|x| x}
end

def calc
	#calculate the normalised value of the effort
	calcNormEffort
	#ap test
	#print $validRow
	efforts = getDeveloperTaskEffort
	updateColumns $ws2,$columnNumbers2["Developer"],efforts,$columnNumbers2["AssignedEffort"],true
	nonTaskEfforts =  getDeveloperNonTaskEffort
	updateColumns $ws2,$columnNumbers2["Developer"],nonTaskEfforts,$columnNumbers2["Misc Tasks"],true
	#Check for any errors 
	errorCheck
end

def run
	#Parse the command
	$args = Utils.parseCommandArgs ARGV
	#See if excel is already running
	begin
		$excel = WIN32OLE.connect('Excel.Application')
    rescue
       #Excel is not running , open one
	   $excel = WIN32OLE.new('Excel.Application')
    end
	#Default name for the excel file to process, if specified then read from the command line
	$excel_name = $args["file"]
	$excel_path = File::join(Dir::pwd, $excel_name)   
	initialise
	
	case $args["command"]
	when "calc"
		calc
	when "sync"
	else
	end

	#Save the workbook at the end
	$wb.Save
	$excel.Interactive = true
	#To turn off screen updating, set the ScreenUpdating object's Interactive property:
	$excel.ScreenUpdating = true
	#sleep(3)

	#excel.Quit
end

run