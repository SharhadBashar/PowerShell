Function CreateHTMLFile([string]$Path,[string[]]$HeadData,[string]$BodyData)
# CreateHTMLFile: Functions to output HTML file as record of deployment
# Parameters:
# 	-> Path: Where to save the HTML File
#	-> HeadData: Data for the Head of the HTML File
#   -> BodyData: Data for the Body of the HTML File
{
	$head = $HeadData
	
	$body = $BodyData
		
	$null | ConvertTo-HTML -head $head -body $body | Set-Content $Path
}

Function GenerateReport
{
	$TranscriptPath =  "$Currentdir\test.txt"
	$tableFragment = $testTable | ConvertTo-HTML "Column 1", "Column 2", "Column 3" -fragment
	$testInfoHTML = "<br>Testing a table:<br>$tableFragment<br><hr>"
	
	
	
	$Currentdir = [string](Get-location) 
	$FileName = "test.html"
	$Path =  "$Currentdir\$Filename"
	#note when putting an array on multiple lines, the final closing one ("@) cannot be indented at all, must be right at the beginning of the line
	$headTag = @"
	<style>
	TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
	TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #CAE8EA;}
	TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
	</style>
	<title>
	$Title
	</title>
"@
	
	$BodyStart = "<h1 style=""color:#9fb11e;font-size:30px"">Testing how to make a table</h1><br><h3 style=""color:#9fb11e;margin-left:30px;""></h3><hr>"
	$BodyTag = "$Bodystart $testInfoHTML"
	CreateHTMLFile $Path $headTag $BodyTag
}



##build tables
#create table object
$testTable = New-Object system.Data.DataTable "Test Table"

#Define Columns
$col_1 = New-Object system.Data.DataColumn colName_1,([string])
$col_1.ColumnName = "Column 1"
$col_2 = New-Object system.Data.DataColumn colName_2,([string])
$col_2.ColumnName = "Column 2"
$col_3 = New-Object system.Data.DataColumn colName_3,([string])
$col_3.ColumnName = "Column 3"

#add them
$testTable.columns.add($col_1)
$testTable.columns.add($col_2)
$testTable.columns.add($col_3)

#create row
$row_1 = $testTable.NewRow()
$row_2 = $testTable.NewRow()
$row_3 = $testTable.NewRow()

#enter data in rows
$row_1["Column 1"] = "1-1"
$row_1["Column 2"] = "1-2"
$row_1["Column 3"] = "1-3"

$row_2["Column 1"] = "2-1"
$row_2["Column 2"] = "2-2"
$row_2["Column 3"] = "2-3"

$row_3["Column 1"] = "3-1"
$row_3["Column 2"] = "3-2"
$row_3["Column 3"] = "3-3"

#add rows
$testTable.Rows.add($row_1)
$testTable.Rows.add($row_2)
$testTable.Rows.add($row_3)

#display table

$testTable | format-table -AutoSize 
GenerateReport