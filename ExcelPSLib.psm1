<#
.NOTES
	NAME: ExcelPSLib.psm1
	AUTHOR: Tomson Philip
    CONTRIBUTORS: Singelé Cédric, Haot Vincent, Elliston Jack
	DATE: 31/07/13
	KEYWORDS: OOXML, MICROSOFT EXCEL
	VERSION : 0.7.0
    LICENSE: LGPL 2.1

    This PowerShell Module allow simple creation of XLSX file by using the EPPlus 4.1 .Net DLL 
    made by Jan Kallman and Licensed under LGPL 2.1 and available at http://epplus.codeplex.com/ .
    Copyright (C) 2014  Tomson Philip

    This library is free software; you can redistribute it and/or
    modify it under the terms of the GNU Lesser General Public
    License as published by the Free Software Foundation; either
    version 2.1 of the License, or (at your option) any later version.

    This library is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
    Lesser General Public License for more details.

    You should have received a copy of the GNU Lesser General Public
    License along with this library; if not, write to the Free Software
    Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA

.SYNOPSIS

    VERSION 0.7.0 (27/03/2019)
        - Improved Cmdlet : Export-OOXML
          It is now possible to do a cumulative export with the -AddToExistingDocument switch parameter
          You MUST provide another worksheetname each time or it won't work

    VERSION 0.6.9 (04/10/2017)
        - Improved Cmdlet : Add-OOXMLWorksheet
          It is now possible to push an array of identical objects at the worksheet creation
          There is no style support for the moment
          Autofilter will not work
        - Fixed Cmdlet : New-OOXMLStyleSheet
          If a style name already exists, this is the corresponding stylesheet that will be returned
          TODO : Implement data type auto style.
          IDEA : Modify Export-OOXML and Add-OOXMLWorksheet so they could use the same code and features
          

    VERSION 0.6.8 (26/04/2016)
        - Improved Cmdlet : Import-OOXML
          The Import-OOXML always consider that first data row is actualy de data header
          that will be used to generate the data object model. In the previous versions,
          There was no check to see if a header was empty or null. Now the Cmdlet will read
          each header and will stop if no header value is found.

    VERSION 0.6.7 (19/04/2016)
        - Improved Cmdlet : Import-OOXML
          Added worksheetname back as an option to select which worksheet to use
          Added range so a range could be used as the import source rather than
          the whole worksheet
        - Added a new Cmdlet : Open-OOXML
          This simply opens an existing Excel file for use 
        - Added a new Cmdlet : Convert-OOXMLFromExcelCoordinates
          This converts an Excel cell address or range into a hash containing
          the related row and column numbers
        - Added a new Cmdlet : Read-OOXMLCell
          Given a worksheet this will return the contents of a cell either
          using row/column number or the Excel address

    VERSION 0.6.6 (09/11/2016)
        - Improved the Export-OOXML function by allowing you to add data validation per column
        - Added a new Cmdlet : Get-OOXMLDataValidationCustomObject
          This Cmdlet is a user friendly way to create a custom object
        - Added a new Cmdlet : Get-OOXMLDataValidationAssignementCustomObject
          This Cmdlet is a user friendly way to create a custom object
        - Added a new Cmdlet : Add-OOXMLDataValidation
          This Cmdlet assign a data validation constrain on a target range.
        - Optimised some part of the code
        - Added a complete example of how to use the Export-OOXML Cmdlet

        #TODO
        - Refactor the Export-OOXML Cmdlet to a more readable and maintable Cmdlet !

    version 0.6.5 (14/10/2016)
        - Fixed a bug introduced in version 0.6.4
          This bug always setted the Conditional formating to the precise mode !
        - "UnPrecise" mode (default) has been improved and can (only) handle the following conditions :
            * BeginsWith
            * ContainsText
            * EndsWith
            * Equal
            * GreaterThan
            * GreaterThanOrEqual
            * LessThan
            * LessThanOrEqual

    version 0.6.4 (13/10/2016)
        - Improved the Export-OOXML function by allowing you to select the columns you want as output instead of
          the current behaviour that output each property to a column. (It is an ordered list !!!)
        - Imporved the Export-OOXML function by allowing you to order the columns : Ascending or Descending.
          This can be combined with the properties/columns selection
        - Improved the Export-OOXML function by allowing you to set a freeze pane at the selected column

    version 0.6.3 (07/09/2016)
        - Improved the Export-OOXML function, which by default will color a whole row in place of a single cell
          For the one who want to keep the previous functionnality just use the switch "Precise" as parameter.

    version 0.6.2 (29/08/2016)
	    - Added a Dll containing the 3 Enum : 
            EnumConditionalFormattingRuleTypeand
            EnumOperations
            EnumColors 
        - Refactored the .psm1 module file

    version 0.6.1 (10/08/2016)
        - Added the parameter "TextRotation" to Set-OOXMLRangeTextOptions
        - Added the parameter "TextRotation" to New-OOXMLStyleSheet
        - Added the parameter "$HeaderTextRotation" to Export-OOXML
        - Added the possibility to define and assign a custom style to one or more column header with the Export-OOXML Cmdlet
        - Added the parameters "Merge" + "RowEnd", "ColEnd" to Set-OOXMLRangeValue so you can now merge a range of cells 
        - Improved Get-OOXMLColumnString (20% faster) by using a native static function : OfficeOpenXml.ExcelCellAddress.GetColumnLetter(int column) 
        - Added/fixed parameter comments that were wrong or missing
        - EPPLUS DLL is now version 4.1 (Stable version)
    
    version 0.6.0 (08/09/2015)
        - Fixed the exception thrown if no HeaderStyle parameter was provided to Export-OOXML (CodePlex - Issue ID #3)
        - Added some try/catch pattern into Cmdlet (It's a WIP so not every Cmdlet was updated)

    version 0.5.9 (04/09/2015)
        This update in mainly focused on the functionnalities of the Export-OOXML Cmdlet
        All The next versions will be added both to Chocolatey and to Codeplex
        - Added 2 New Enum: "EnumColors"(141 color name extracted from ) & EnumOperations (5 Basic & 3 Conditional Excel Formula Operators)
        - Export-OOXML => Added 137 color styles to the original 4 to use with the Cmdlet Get-OOXMLConditonalFormattingCustomObject
        - Export-OOXML => Added Support for basic math operations on columns : "SUM","AVERAGE","COUNT","MAX","MIN"
        - Export-OOXML => Added Support for Conditional math operations on columns : "SUMIF","AVERAGEIF","COUNTIF"
        - Export-OOXML => Added 137 color Styles for the column headers
        - EPPLUS DLL is now version 4.0.4 (Stable version)

    version 0.5.8 (07/08/2015)

        - Improved Import-OOXML cmdlet so it can "auto-sense" data types if asked by adding the "KeepDataType" switch parameter.
          *** Warning, for this to work the data type must be the same in the whole column, if one single cell in 
              the column is of a different data type the data type will always be set to "string" even if the "KeepDataType" 
              was set.

    version 0.5.7 (29/09/2014)

        - Improved the Export-OOXML cmdlet so it can "auto-sense" data types and apply the correct formatting to cells
        - Improved the Export-OOXML cmdlet so it can recognize URI and set the HyperLink propertie of the cell accordingly
        - Added Import-OOXML cmdlet to convert an XLSX file to an array of object this function is still basic and requires
          some fixed Excel sheet format. If you do an Export-OOXML and then an Import-OOXML with the generated XLSX as 
          input file, everything should be fine.
        - Fixed some small "gliches"

    version 0.5.6 (24/09/2014)

        - Added a new command-line Get-OOXMLConditonalFormattingCustomObject that returns a "PSCustomObject"
          ready to be used with the Export-OOXML "ConditionalFormatings" parameter. It has Auto-Complete for "Style" and "Condition"
        - Improved the Export-OOXML cmdlet with a new switch parameter "AutoFit" that will resize all the column according
          to the size of their content
        - Fixed the way that conditional formatting was applied in the Export-OOXML because the range was row count +1
        - Fixed the "invoke member w/ expression name" exception introduced in version 0.5.5 for those using PS 3.0... Sorry about this !

        #TODO
        - Add condition priority
        - Check if properties are defined in a style sheet before using them with Add-OOXMLConditionalFormatting (DONE)
        - Add the possibility to set cell Text/Numberformat per column with the Export-OOXML cmdlet (DONE => Auto-Sensing)

    version 0.5.5 (23/09/2014)

        - Added a reduced enum EnumConditionalFormattingRuleType based on the
          OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingRuleType enum
        - Added a new cmdlet Add-OOXMLConditionalFormatting that let you add conditional formating rules
          to single or multiple ranges
        - Improved the Export-OOXML Cmdlet with a new parameter that allow you to use 4 Conditional Styles
          that are named: Red, Orange, Yellow and Green. This new "ConditionalFormatings" parameter can receives
          array of objects of the following format:

              @(
                [PSCustomObject]@{
                    Name="DeviceID";
                    Style="Red";
                    Condition="BeginsWith";
                    Value="L"
                }
              )

              "Name" is the name of one propertie of the array of object you want to export
              "Style" is the style you want to apply Red, Orange, Yellow and Green
              "Condition" is a string coming from the enum "EnumConditionalFormattingRuleType"
              "Value" is the condition value.

              For more informations see the Demo.ps1 file
    
    version 0.5.4 (16/09/2014)

        - Added a new command-line Export-OOXML that will export an array of object to an XLSX file

    version 0.5.3 (10/07/2014)

        - Fixed the problem with the value in the Set-OOXMLRangeValue that was always a "string"

    Version 0.5.2 (26/03/2014)

        - Added worksheet's name maximum length check in the Add-OOXMLWorksheet cmdlet

    Version 0.5.1 (06/03/2014)

        - Added "OutlineLevel" parameter to the Set-OOXMLRangeValue cmdlet

    Version 0.5 (26/02/2014)

        - The "Color" is no more a "string" but uses the "System.Drawing.Color" type
        - Added Set-OOXMLStyleSheet and New-OOXMLStyleSheet so you can now define style and then recall and apply them
        - Added a really basic Pivot Table cmdlet New-OOXMLPivotTable
        - Modified the Set-OOXMLRangeValue cmdlet to accept 2 more parameters "StyleSheet" and "StyleSheetName"
        - Introduction of ParametersetName in some cmdlets to enforce good cmdlet usage.
        - Added complete cmdlet info bloc for each cmdlet
        - Renamed Get-ColumnString to Get-OOXMLColumnString and created an alias for backward compatibility

    Version 0.4 (25/09/2013)

        - Added ValueFromPipeline to allmost all functions (If relevant)
        - Fixed some Class Type casting (OfficeOpenXml.ExcelRange => OfficeOpenXml.ExcelRangeBase)
        - Added Return to allmost all functions (If relevant) so you can now chain them like :
          $Worksheet | Set-OOXMLRangeValue -row 1 -col 1 -value "Test Value" | Set-OOXMLRangeBorder -borderStyle "DashDotDot" -color "Green"
        - Modified the Save-OOXMLPackage "CmdLet" so it now uses the "Dispose" method if the "Dispose" switch is used
        - The "BorderStyle" is no more a "string" but uses the Enum "OfficeOpenXml.Style.ExcelBorderStyle"
        - The "FillStyle" is no more a "string" but it uses the Enum "OfficeOpenXml.Style.ExcelFillStyle"
        - The "HorizontalAlignment" is no more a "string" but it uses the Enum "OfficeOpenXml.Style.ExcelHorizontalAlignment"
        - The "VerticalAlignment" is no more a "string" but it uses the Enum "OfficeOpenXml.Style.ExcelVerticalAlignment"


	VERSION 0.3a (13/08/2013)
		
		- Fixed the Set-OOXMLRangeBorder cmdlet that was still using the old cmdlet
		- Removed the usage example present in this module

	VERSION 0.3 (13/08/2013)
	
	    - Renamed all cmdlets to respect standards
	    - Added the "Get-OOXMLDeprecatedCommand" to allow you to use the old cmdlets
	    - Added the "Repair-OOXMLLib" cmdlet to set Aliases
	    - Added the "Convert-OOXMLOldScripts" to convert you script with 0.2 style cmdlets to the 0.3 style
	      This cmdlet is very basic and should work in many case but it is more a "brute force" conversion than
	      something "smart" so use it if you dare !
	    - Adde the "Get-OOXMLHelp" to print the syntax of all cmdlets at once (Ex: output it to a file)
	    - Reformated all comments and infos
	    - Compatible PowerShell 2.0(*)
		    * There was an issue with PS 2.0 :
		      Mmulti-dimentional .Net tables like cells[1,1] or cells[1,1,1,5] were not understood !
		      So If you use power try to use literal addressing like "A1" or "A1:E1"
	    - Added the "Convert-OOXMLCellsCoordinates" cmdlet to convert coordinate like [1,1] to "A1" or like [1,1,1,5] to "A1:E1"
	    - Added the "Get-ColumnString" cmdlet that is normaly used by "Convert-OOXMLCellsCoordinates" but you can use it to
	      Convert coordinate like [1,1] to "A1". I recommend the usage of "Convert-OOXMLCellsCoordinates" instead of "Get-ColumnString"
	      for single cell coordinate convertion.

    VERSION 0.2 (02/08/2013)
    
        - Added Default Row, Col Size and AutoFilter range at Sheet Creation
        - Added a new Cmdlet SetTextOptions allowing to set Text formating options for a cell range
        - Added a new Cmdlet SetPrinterSettings allowing to set some printer settings

    VERSION 0.1 (31/07/13)

        This PowerShell Module to allow simple creation of XLSX file by using the EPPlus 3.1 .Net DLL
        available at http://epplus.codeplex.com/ and was made by Jan Kallman.

        The current set of feature is the following :
        - Create a Microsoft Excel Workbook
        - Add Worksheet to a Workbook
        - Select a Worksheet in a Workbook
        - Define the font style
        - Define border style
        - Define Cell color and Fill type
        - Save the Workbook to a file
        - Set the value of a Cell as Text or Hyperlink
        - Set AutoFitColumns minimum width
        - Select a range of cell

	TODO:

        - Add more error checking within function

#>

<#---------------------------------------[ Variables ]---------------------------------------#>


<#---------------------------------------[ Functions ]---------------------------------------#>

Function New-OOXMLPackage {
    <#
    .SYNOPSIS
    Create an ExcelPackage instance, configure the workbook and return the ExcelPackage instance. 
    
    .DESCRIPTION
    Create an ExcelPackage instance, configure the workbook and return the ExcelPackage instance.
    
    .PARAMETER Author
    An author to be added to the workbook.

    .PARAMETER Title
    A title to be added to the workbook.
    
    .PARAMETER Comment
    A comment to be added to the workbook.

    .PARAMETER Path
    The path of XLSX file

    .EXAMPLE
    [OfficeOpenXml.ExcelPackage]$excel = $(New-OOXMLPackage -Author "Mr.Excel" -Title "Workbook title" -Comment "Workbook comment")

    Description
    -----------
    Calls a function which create and returns a "OfficeOpenXml.ExcelPackage" object
    
    .NOTES
    
    .LINK 
    
    #> 
    [CmdletBinding()]
	param (
        [parameter(Mandatory=$true)]
		[string]$Author,
        [parameter(Mandatory=$true)]
		[string]$Title,
        [string]$Comment,
        [string]$Path
	)
	process{
        
        if($Path){
            [System.IO.FileInfo]$XLSXFile = New-Object System.IO.FileInfo($Path)
		    $ExcelInstance = New-Object OfficeOpenXml.ExcelPackage($XLSXFile)
        }else{
            $ExcelInstance = New-Object OfficeOpenXml.ExcelPackage
        }

		$ExcelInstance.Workbook.Properties.Author = $Author
		if($Title){$ExcelInstance.Workbook.Properties.Title = $Title}
        if($Comment){$ExcelInstance.Workbook.Properties.Comments = $Comment}
		return [OfficeOpenXml.ExcelPackage]$ExcelInstance
	}
}

Function Add-OOXMLWorksheet {
    <#
    .SYNOPSIS
    Add a worksheet to the workbook and configure the worksheet. 
    
    .DESCRIPTION
    Add a worksheet to the workbook and configure the worksheet. 
    
    .PARAMETER DefColWidth
    Default width of the columns in the worksheet.

    .PARAMETER DefRowHeight
    Default height of the rows in the worksheet.
    
    .PARAMETER AutofilterRange
    Set a range on which you want to enable the Auto-Filter feature

    .PARAMETER WorkSheetName
    The name of the worksheet

    .PARAMETER ExcelInstance
    The Current ExcelPackage instance

    .PARAMETER InputObject
    An Optional array of objects and insert it

    .EXAMPLE
    Add-OOXMLWorksheet -ExcelInstance $excel -WorkSheetName "New Worksheet"
    $excel | Add-OOXMLWorksheet -WorkSheetName "New Worksheet"
    $excel | Add-OOXMLWorksheet -WorkSheetName "New Worksheet" -DefColWidth 20 -DefRowHeight 10

    Description
    -----------
    Calls a function which create a new worksheet in the workbook of the current ExcelInstance Object
    
    .NOTES
    
    .LINK 
    
    #> 
    [CmdletBinding()]
	param (
        [int]$DefColWidth,
        [int]$DefRowHeight,
		[string]$AutofilterRange,
        [parameter(Mandatory=$true)]
        [ValidateScript({$_.length -lt 30})]
		[String]$WorkSheetName,
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]
		[OfficeOpenXml.ExcelPackage]$ExcelInstance,
        [Object[]]$InputObject
	)
	process{
		    $ExcelInstance.Workbook.Worksheets.Add($WorkSheetName) | Out-Null
            $SheetNumber = $ExcelInstance.Workbook.Worksheets.Count
            $loop = $true
            $i=1
            while($loop){

                if($ExcelInstance.Workbook.Worksheets[$i].Name -eq $WorkSheetName){

                    $sheet = $ExcelInstance.Workbook.Worksheets[$i]
                    $book = $ExcelInstance.Workbook

                    if($InputObject)
                    {

                        $StyleNormal = New-OOXMLStyleSheet -WorkBook $book -Name "NormalStyle" -borderStyle Thin -BorderColor Black -HAlign Right
                        $StyleURI = New-OOXMLStyleSheet -WorkBook $book -Name "URIStyle" -borderStyle Thin -BorderColor Black -HAlign Left -ForeGroundColor Blue -Underline
        
                        $StyleDate = New-OOXMLStyleSheet -WorkBook $book -Name "DateStyle" -borderStyle Thin -BorderColor Black -HAlign Right -NFormat "$([System.Globalization.DateTimeFormatInfo]::CurrentInfo.ShortDatePattern) $([System.Globalization.DateTimeFormatInfo]::CurrentInfo.ShortTimePattern)"
                        $StyleNumber = New-OOXMLStyleSheet -WorkBook $book -Name "NumberStyle" -borderStyle Thin -BorderColor Black -HAlign Right -NFormat "0"
                        $StyleFloat = New-OOXMLStyleSheet -WorkBook $book -Name "FloatStyle" -borderStyle Thin -BorderColor Black -HAlign Right -NFormat "0.00"

                        $ReferencePropertySet = $($InputObject[0].PSObject.Properties).Name
                        $ColumnNumber = $ReferencePropertySet.Length
                        $RowPosition = 1
                        
                        $j=1
                        foreach($Property in $ReferencePropertySet)
                        {
                            $sheet | Set-OOXMLRangeValue -row $RowPosition -col $j -value $Property | Out-Null
                            $sheet.Column($j).Width = 32
                            $j++
                        }

                        $RowPosition++

                        foreach($Object in $InputObject){
                            $j=1
                            foreach($Property in $ReferencePropertySet){
                                $Value = "Empty Value"

                                
                                if($($Object.$Property) -ne $null){
                                    $Value = $($Object.$Property)
                    
                                }                               

                                $IsURI = $false
                                $AppliedStyle = $StyleNormal
                                switch -regex ($($Value.GetType())){
                                    "(^uint[0-9]{2}$)|(^int[0-9]{2}$)|(^long$)|(^int$)" {
                                        $AppliedStyle = $StyleNumber
                                    }
                                    "(double)|(float)|(decimal)" {
                                        $AppliedStyle = $StyleFloat
                                    }
                                    "datetime" {
                                        $AppliedStyle = $StyleDate
                                    }
                                    "^string$"{
                                        if($([System.URI]::IsWellFormedUriString([System.URI]::EscapeUriString($Value),[System.UriKind]::Absolute)) -and $($Value -match "(^\\\\)|(^http://)|(^ftp://)|(^[a-zA-Z]:(//|\\))|(^https://)"))
                                        {
                                            $AppliedStyle = $StyleURI
                                            $IsURI = $true
                                        }
                                    }
                                }

                                <#

                                if($IsURI){
                                    $sheet | Set-OOXMLRangeValue -Row $RowPosition -Col $j -Value $Value -StyleSheet $AppliedStyle -Uri | Out-Null
                                }else{
                                    $sheet | Set-OOXMLRangeValue -Row $RowPosition -Col $j -Value $Value -StyleSheet $AppliedStyle | Out-Null
                                }
                                #>

                                $sheet | Set-OOXMLRangeValue -Row $RowPosition -Col $j -Value $Value | Out-Null

                                $j++
                                

                            }
                            $RowPosition++
                        }

                        $EndColumn = Get-OOXMLColumnString -ColNumber $($ReferencePropertySet.Length)
                        $FirstColumn = Get-OOXMLColumnString -ColNumber 1
                        $Sheet.Cells["$FirstColumn$($Sheet.Dimension.Start.Row):$EndColumn$LastRow"].AutoFitColumns()

                    }

                    if($DefColWidth){$ExcelInstance.Workbook.Worksheets[$i].DefaultColWidth = $DefColWidth}
                    if($DefRowHeight){$ExcelInstance.Workbook.Worksheets[$i].DefaultRowHeight = $DefRowHeight}
				    if($AutofilterRange){$ExcelInstance.Workbook.Worksheets[$i].Cells[$AutofilterRange].AutoFilter=$true}
                    $loop = $false
                    
                }
                $i++
            }
	}
}

Function Get-OOXMLWorkbook {
    <#
    .SYNOPSIS
    Get the workbook in the ExcelInstance object
    
    .DESCRIPTION
    Get the workbook in the ExcelInstance object

    .PARAMETER ExcelInstance
    The Current ExcelPackage instance

    .EXAMPLE
    $book = $excel | Get-OOXMLWorkbook
    $book = Get-OOXMLWorkbook -ExcelInstance ExcelPackage

    Description
    -----------
    Calls a function which return the workbook of the current ExcelInstance Object
    
    .NOTES
    
    .LINK 
    
    #> 
    [CmdletBinding()]
	param(
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]
		[OfficeOpenXml.ExcelPackage]$ExcelInstance
	)
	process{
		return [OfficeOpenXml.ExcelWorkbook]$ExcelInstance.Workbook
	}
}

Function Select-OOXMLWorkSheet {
    <#
    .SYNOPSIS
    Get a worksheet by name or by number from the given workbook in the ExcelInstance object
    
    .DESCRIPTION
    Get a worksheet by name or by number from the given workbook in the ExcelInstance object

    .PARAMETER WorkBook
    The workbook in the Excel instance 

    .PARAMETER WorkSheetNumber
    The worksheet index number

    .PARAMETER WorkSheetName
    The worksheet name

    .EXAMPLE
    $sheet = $book | Select-OOXMLWorkSheet -WorkSheetNumber 1
    $sheet = $book | Select-OOXMLWorkSheet -WorkSheetName "My Worksheet"

    Description
    -----------
    Calls a function which return a worksheet in the given workbook object
    
    .NOTES
    
    .LINK 
    
    #> 
    [CmdletBinding()]
	param (
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]
		[OfficeOpenXml.ExcelWorkbook]$WorkBook,
        [parameter(ParameterSetName="WorksheetIndex", Mandatory=$true)]
		[int]$WorkSheetNumber,
        [parameter(ParameterSetName="WorksheetName", Mandatory=$true)]
        [string]$WorkSheetName
	)
	process{
        if($WorkSheetName){
            $SheetNumberPlusOne = $($Workbook.Worksheets.Count + 1)
            $i=1
            while($i -lt $SheetNumberPlusOne){
                if($Workbook.Worksheets[$i].Name -like $WorkSheetName){
                    return [OfficeOpenXml.ExcelWorksheet]$WorkBook.Worksheets[$i]
                }
                $i++
            }
        }
        $WorkSheet = [OfficeOpenXml.ExcelWorksheet]$WorkBook.Worksheets[$WorkSheetNumber]
		return [OfficeOpenXml.ExcelWorksheet]$WorkSheet
	}
}
<#
Function Get-OOXMLRangeValue
{
    
}
#>

Function Set-OOXMLRangeValue {
    <#
    .SYNOPSIS
    Set the value in a cell and optionally apply a stylesheet to it
    
    .DESCRIPTION
    Set the value in a cell and optionally apply a stylesheet to it

    .PARAMETER Row
    The start row index expressed as an integer

    .PARAMETER Col
    The start column index expressed as an integer

    .PARAMETER RowEnd
    The end row index expressed as an integer

    .PARAMETER ColEnd
    The end column index expressed as an integer

    .PARAMETER Value
    The value you want to set in the cell

    .PARAMETER WorkSheet
    The WorkSheet object where the cell is located

    .PARAMETER Uri
    This option will try to convert the value to an hyperlink

    .PARAMETER StyleSheet
    The style sheet you want to apply to the cell

    .PARAMETER StyleSheetName
    the style sheet name that you want to apply to the cell

    .PARAMETER OutlineLevel
    The ouline Level for the whole row containing the cell

    .EXAMPLE
    $sheet | Set-OOXMLRangeValue -row 1 -col 1 -value "http:\\excelpslib.codeplex.com" -StyleSheetName "New Style" -Uri
    $sheet | Set-OOXMLRangeValue -Merge -Row 1 -Col 1 -RowEnd 10 -ColEnd 10 -Value "Merged Cells" -StyleSheetName "New Style"
    $Range = $sheet | Set-OOXMLRangeValue -row 1 -col 1 -value "http:\\excelpslib.codeplex.com" -StyleSheetName "New Style" -Uri
    
    Description
    -----------
    Calls a function which set the value in a cell and optionally apply a stylesheet to it and return a ExcelRangeBase object in the given workbook object
    
    .NOTES
    
    .LINK 
    
    #> 
    [CmdletBinding()]
	param (
        [parameter(ParameterSetName="CellRangeMerge", Mandatory=$true)]
        [switch]$Merge,
        [parameter(ParameterSetName="CellRangeMerge", Mandatory=$true)]
        [parameter(ParameterSetName="NoCellRangeMerge", Mandatory=$true)]
		[string]$Row,
        [parameter(ParameterSetName="CellRangeMerge", Mandatory=$true)]
        [parameter(ParameterSetName="NoCellRangeMerge", Mandatory=$true)]
        [string]$Col,
        [parameter(ParameterSetName="CellRangeMerge", Mandatory=$true)]
		[string]$RowEnd,
        [parameter(ParameterSetName="CellRangeMerge", Mandatory=$true)]
        [string]$ColEnd,
        [parameter(ParameterSetName="CellRangeMerge", Mandatory=$true)]
        [parameter(ParameterSetName="NoCellRangeMerge", Mandatory=$true)]
		[object]$Value,
        [parameter(ParameterSetName="CellRangeMerge", Mandatory=$true, ValueFromPipeline=$true)]
        [parameter(ParameterSetName="NoCellRangeMerge", Mandatory=$true, ValueFromPipeline=$true)]
		[OfficeOpenXml.ExcelWorksheet]$WorkSheet,
        [parameter(ParameterSetName="CellRangeMerge")]
        [parameter(ParameterSetName="NoCellRangeMerge")]

        [switch]$Uri,
        [parameter(ParameterSetName="CellRangeMerge")]
        [parameter(ParameterSetName="NoCellRangeMerge")]
		[OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml]$StyleSheet,
        [parameter(ParameterSetName="CellRangeMerge")]
        [parameter(ParameterSetName="NoCellRangeMerge")]
		[String]$StyleSheetName,
        [parameter(ParameterSetName="CellRangeMerge")]
        [parameter(ParameterSetName="NoCellRangeMerge")]
        [int]$OutlineLevel
        
	)
	process{

        if($Merge)
        {
            $Coordinates = Convert-OOXMLCellsCoordinates -StartRow $row -StartCol $col -EndRow $RowEnd -EndCol $ColEnd
            $workSheet.Cells[$Coordinates].Merge = $true
        }
        else
        {
            $Coordinates = Convert-OOXMLCellsCoordinates -StartRow $row -StartCol $col
        }
        
        $WorkSheet.SetValue($row, $col, $value) | Out-Null
        
        if($OutlineLevel){
            $WorkSheet.Row($row).OutlineLevel($OutlineLevel)
        }

        if($Uri){
			$workSheet.Cells[$Coordinates].Hyperlink = new-object System.Uri($value)
        }

        if($StyleSheet){
            $workSheet.Cells[$Coordinates].StyleName = $StyleSheet.Name
        }elseif($StyleSheetName){
            $workSheet.Cells[$Coordinates].StyleName = $StyleSheetName
        }
        
        return [OfficeOpenXml.ExcelRangeBase]$workSheet.Cells[$Coordinates]
	}
}

Function Set-OOXMLRangeBorder {
    <#
    .SYNOPSIS
    Set the border style options for range of cell
    
    .DESCRIPTION
    Set the border style options for range of cell

    .PARAMETER BorderStyle
    The border style that will be applied to the range of cell

    .PARAMETER CellRange
    The cell range that where the options are to be applied

    .PARAMETER Color
    The color that will be applied to the range of cell

    .EXAMPLE
    $Range = Set-OOXMLRangeBorder -cellRange $range -borderStyle Thick -color red
    $Range = $Range | Set-OOXMLRangeBorder -borderStyle Thick -color red

    Description
    -----------
    Calls a function which set the border style options for range of cells and return an ExcelRangeBase object
    
    .NOTES
    
    .LINK 
    
    #> 
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true)]
        [OfficeOpenXml.Style.ExcelBorderStyle]$BorderStyle,
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [OfficeOpenXml.ExcelRangeBase]$CellRange,
        [parameter(Mandatory=$true)]
        [System.Drawing.Color]$Color
    )
    process{
        Set-OOXMLRangeBorderTop -borderStyle $BorderStyle -cellRange $CellRange -color $Color
        Set-OOXMLRangeBorderRight -borderStyle $BorderStyle -cellRange $CellRange -color $Color
        Set-OOXMLRangeBorderBottom -borderStyle $BorderStyle -cellRange $CellRange -color $Color
        Set-OOXMLRangeBorderLeft -borderStyle $BorderStyle -cellRange $CellRange -color $Color
        return [OfficeOpenXml.ExcelRangeBase]$CellRange
    }

}

Function Set-OOXMLRangeBorderTop {
    <#
    .SYNOPSIS
    Set the top border style options for range of cell
    
    .DESCRIPTION
    Set the top border style options for range of cell

    .PARAMETER BorderStyle
    The border style that will be applied to the range of cell

    .PARAMETER CellRange
    The cell range that where the options are to be applied

    .PARAMETER Color
    The color that will be applied to the range of cell

    .EXAMPLE
    $Range = Set-OOXMLRangeBorderTop -cellRange $range -borderStyle Thick -color red
    $Range = $Range | Set-OOXMLRangeBorderTop -borderStyle Thick -color red

    Description
    -----------
    Calls a function which set the top border style options for range of cells and return an ExcelRangeBase object
    
    .NOTES
    
    .LINK 
    
    #> 
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true)]
        [OfficeOpenXml.Style.ExcelBorderStyle]$borderStyle,
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [OfficeOpenXml.ExcelRangeBase]$cellRange,
        [parameter(Mandatory=$true)]
        [System.Drawing.Color]$color
    )
    process{
        $cellRange.Style.Border.Top.Style = $borderStyle
        $cellRange.Style.Border.Top.Color.SetColor($color)
        return [OfficeOpenXml.ExcelRangeBase]$cellRange
    }
}

Function Set-OOXMLRangeBorderRight {
    <#
    .SYNOPSIS
    Set the right border style options for range of cell
    
    .DESCRIPTION
    Set the right border style options for range of cell

    .PARAMETER BorderStyle
    The border style that will be applied to the range of cell

    .PARAMETER CellRange
    The cell range that where the options are to be applied

    .PARAMETER Color
    The color that will be applied to the range of cell

    .EXAMPLE
    $Range = Set-OOXMLRangeBorderRight -cellRange $range -borderStyle Thick -color red
    $Range = $Range | Set-OOXMLRangeBorderRight -borderStyle Thick -color red

    Description
    -----------
    Calls a function which set the right border style options for range of cells and return an ExcelRangeBase object
    
    .NOTES
    
    .LINK 
    
    #> 
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true)]
        [OfficeOpenXml.Style.ExcelBorderStyle]$borderStyle,
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [OfficeOpenXml.ExcelRangeBase]$cellRange,
        [parameter(Mandatory=$true)]
        [System.Drawing.Color]$color
    )
    process{
        $cellRange.Style.Border.Right.Style = $borderStyle
        $cellRange.Style.Border.Right.Color.SetColor($color)
        return [OfficeOpenXml.ExcelRangeBase]$cellRange
    }
}

Function Set-OOXMLRangeBorderBottom {
    <#
    .SYNOPSIS
    Set the bottom border style options for range of cell
    
    .DESCRIPTION
    Set the bottom border style options for range of cell

    .PARAMETER BorderStyle
    The border style that will be applied to the range of cell

    .PARAMETER CellRange
    The cell range that where the options are to be applied

    .PARAMETER Color
    The color that will be applied to the range of cell

    .EXAMPLE
    $Range = Set-OOXMLRangeBorderBottom -cellRange $range -borderStyle Thick -color red
    $Range = $Range | Set-OOXMLRangeBorderBottom -borderStyle Thick -color red

    Description
    -----------
    Calls a function which set the bottom border style options for range of cells and return an ExcelRangeBase object
    
    .NOTES
    
    .LINK 
    
    #>
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true)]
        [OfficeOpenXml.Style.ExcelBorderStyle]$borderStyle,
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [OfficeOpenXml.ExcelRangeBase]$cellRange,
        [parameter(Mandatory=$true)]
        [System.Drawing.Color]$color
    )
    process{
        $cellRange.Style.Border.Bottom.Style = $borderStyle
        $cellRange.Style.Border.Bottom.Color.SetColor($color)
        return [OfficeOpenXml.ExcelRangeBase]$cellRange
    }
}

Function Set-OOXMLRangeBorderLeft {
    <#
    .SYNOPSIS
    Set the left border style options for range of cell
    
    .DESCRIPTION
    Set the left border style options for range of cell

    .PARAMETER BorderStyle
    The border style that will be applied to the range of cell

    .PARAMETER CellRange
    The cell range that where the options are to be applied

    .PARAMETER Color
    The color that will be applied to the range of cell

    .EXAMPLE
    $Range = Set-OOXMLRangeBorderLeft -cellRange $range -borderStyle Thick -color red
    $Range = $Range | Set-OOXMLRangeBorderLeft -borderStyle Thick -color red

    Description
    -----------
    Calls a function which set the left border style options for range of cells and return an ExcelRangeBase object
    
    .NOTES
    
    .LINK 
    
    #>
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true)]
        [OfficeOpenXml.Style.ExcelBorderStyle]$borderStyle,
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [OfficeOpenXml.ExcelRangeBase]$cellRange,
        [System.Drawing.Color]$color
    )
    process{
        $cellRange.Style.Border.Left.Style = $borderStyle
        $cellRange.Style.Border.Left.Color.SetColor($color)
        return [OfficeOpenXml.ExcelRangeBase]$cellRange
    }
}

Function Set-OOXMLRangeFill {
    <#
    .SYNOPSIS
    Set the fill style options for range of cell
    
    .DESCRIPTION
    Set the fill style options for range of cell

    .PARAMETER Type
    The fill type that will be applied to the range of cell

    .PARAMETER Color
    The color that will be applied to the range of cell
    
    .PARAMETER CellRange
    The cell range that where the options are to be applied

    .EXAMPLE
    $Range = Set-OOXMLRangeFill -cellRange $range -Type Solid -color red
    $Range = $Range | Set-OOXMLRangeFill -Type Solid -color red

    Description
    -----------
    Calls a function which set the fill style options for range of cells and return an ExcelRangeBase object
    
    .NOTES
    
    .LINK 
    
    #>
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true)]
        [OfficeOpenXml.Style.ExcelFillStyle]$Type,
        [parameter(Mandatory=$true)]
        [System.Drawing.Color]$color,
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [OfficeOpenXml.ExcelRangeBase]$cellRange,
        [switch]$Pass
    )
    process{
        $cellRange.Style.Fill.PatternType = $Type
        $cellRange.Style.Fill.BackgroundColor.SetColor($color)
        return [OfficeOpenXml.ExcelRangeBase]$cellRange
    }
}

Function Set-OOXMLRangeFont {
    <#
    .SYNOPSIS
    Set the font options for range of cell
    
    .DESCRIPTION
    Set the font options for range of cell

    .PARAMETER Bold
    Set the font to bold

    .PARAMETER Italic
    Set the font to italic

    .PARAMETER Underline
    Set the font to underlined

    .PARAMETER Strike
    Set the font to striked

    .PARAMETER Size
    Set the font size

    .PARAMETER Color
    The color that will be applied to the range of cell
    
    .PARAMETER CellRange
    The cell range that where the options are to be applied

    .EXAMPLE
    $Range = Set-OOXMLRangeFont -bold -italic -underline -strike -size 12 -color red -cellRange $Range
    $Range = $Range | Set-OOXMLRangeFont -bold -italic -underline -strike -size 12 -color red
    Set-OOXMLRangeFont -bold -italic -underline -strike -size 12 -color red

    Description
    -----------
    Calls a function which set the font options for range of cells and return an ExcelRangeBase object
    
    .NOTES
    
    .LINK 
    
    #>
    [CmdletBinding()]
    param(
        [switch]$bold,
        [switch]$italic,
        [switch]$underline,
        [switch]$strike,
        [float]$size,
        [System.Drawing.Color]$color,
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [OfficeOpenXml.ExcelRangeBase]$cellRange
    )
    process{
        if($bold){$cellRange.Style.Font.Bold = $true}else{$cellRange.Style.Font.Bold = $false}
        if($italic){$cellRange.Style.Font.Italic = $true}else{$cellRange.Style.Font.Italic = $false}
        if($underline){$cellRange.Style.Font.UnderLine = $true}else{$cellRange.Style.Font.UnderLine = $false}
        if($strike){$cellRange.Style.Font.Strike = $true}else{$cellRange.Style.Font.Strike = $false}
        if($size){$cellRange.Style.Font.Size = $size}
        if($color){$cellRange.Style.Font.Color.SetColor($color)}
        return [OfficeOpenXml.ExcelRangeBase]$cellRange
    }
}

Function Set-OOXMLRangeTextOptions {
    <#
    .SYNOPSIS
    Set some text options linked to how text must be displayed within a range of cell
    
    .DESCRIPTION
    Set some text options linked to how text must be displayed within a range of cell

    .PARAMETER HAlign
    Set the horizontal text alignement

    .PARAMETER VAlign
    Set the vertical alignement type

    .PARAMETER NFormat
    Format a number according to a definited patern

    .PARAMETER Wrap
    Force end of the for line that are bigger than the cell 

    .PARAMETER Shrink
    Reduce the size of the text to fit in cell

    .PARAMETER Locked
    Prevent text edition within a cell

    .PARAMETER TextRotation
    Set the angle of the text

    .PARAMETER CellRange
    The cell range that where the options are to be applied

    .EXAMPLE
    $Range = Set-OOXMLRangeTextOptions -cellRange $Range -HAlign Center -VAlign Bottom -Wrap -Locked
    $Range = $Range | Set-OOXMLRangeTextOptions -HAlign Center -VAlign Bottom -Wrap -Locked
    Set-OOXMLRangeTextOptions -cellRange $Range -HAlign Center -VAlign Bottom -Wrap -Locked -TextRotation 90

    Description
    -----------
    Calls a function which set some text options linked to how text must be displayed within a range of cells and return an ExcelRangeBase object
    
    .NOTES
    
    .LINK 
    
    #>
    [CmdletBinding()]
    param(
        [OfficeOpenXml.Style.ExcelHorizontalAlignment]$HAlign,
        [OfficeOpenXml.Style.ExcelVerticalAlignment]$VAlign,
        [string]$NFormat,
        [switch]$Wrap,
        [switch]$Shrink,
        [switch]$Locked,
        [ValidateScript({($_ -ge 0) -and ($_ -le 180)})]
        [int]$TextRotation,
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [OfficeOpenXml.ExcelRangeBase]$cellRange
    )
    process{
        if($HAlign){$cellRange.Style.HorizontalAlignment = $HAlign}
        if($VAlign){$cellRange.Style.VerticalAlignment = $VAlign}
        if($Wrap){$cellRange.Style.WrapText = $true}else{$cellRange.Style.WrapText = $false}
        if($Shrink){$cellRange.Style.ShrinkToFit = $true}else{$cellRange.Style.ShrinkToFit = $false}
        if($NFormat){$cellRange.Style.Numberformat.Format = $NFormat}
        if($Locked){$cellRange.Style.Locked = $true}else{$cellRange.Style.Locked = $false}
        if($TextRotation){$cellRange.Style.TextRotation = $TextRotation}
        return [OfficeOpenXml.ExcelRangeBase]$cellRange
    }
}

Function Set-OOXMLPrinterSettings {
    <#
    .SYNOPSIS
    Set some general printer settings
    
    .DESCRIPTION
    Set some general printer settings

    .PARAMETER HorizontalCentered
    Center horizontaly the sheet on the page

    .PARAMETER VerticalCentered
    Center verticaly the sheet on the page

    .PARAMETER ShowGridLines
    Print gridlines on the page

    .PARAMETER BlackAndWhite
    Print in black and white only

    .PARAMETER FitToPage
    Resize the sheet to fit in the page

    .PARAMETER RowRange
    The cell row range to repeat on every pages

    .PARAMETER ColRange
    The cell column range to repeat on every pages

    .PARAMETER WorkSheet
    The WorkSheet object where the cell is located

    .EXAMPLE
    $sheet = $sheet | Set-OOXMLPrinterSettings -HorizontalCentered -VerticalCentered -FitToPage -ShowGridLines
    $sheet | Set-OOXMLPrinterSettings -HorizontalCentered -VerticalCentered -FitToPage -ShowGridLines
    Set-OOXMLPrinterSettings -WorkSheet $sheet -HorizontalCentered -VerticalCentered -FitToPage -ShowGridLines

    Description
    -----------
    Calls a function which set some print options linked to how the sheet must be printed and return an ExcelWorksheet object
    
    .NOTES
    
    .LINK 
    
    #>
    [CmdletBinding()]
    param(
        [switch]$HorizontalCentered,
        [switch]$VerticalCentered,
        [switch]$ShowGridLines,
        [switch]$BlackAndWhite,
        [switch]$FitToPage,
        [OfficeOpenXml.ExcelRangeBase]$RowRange,
        [OfficeOpenXml.ExcelRangeBase]$ColRange,
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]
		[OfficeOpenXml.ExcelWorksheet]$WorkSheet
    )
    process{
        if($ColRange){$WorkSheet.PrinterSettings.RepeatColumns = $ColRange}
        if($RowRange){$WorkSheet.PrinterSettings.RepeatRows = $RowRange}
        if($BlackAndWhite){$WorkSheet.PrinterSettings.BlackAndWhite = $true}else{$WorkSheet.PrinterSettings.BlackAndWhite = $false}
        if($ShowGridLines){$WorkSheet.PrinterSettings.ShowGridLines = $true}else{$WorkSheet.PrinterSettings.ShowGridLines = $false}
        if($HorizontalCentered){$WorkSheet.PrinterSettings.HorizontalCentered = $true}else{$WorkSheet.PrinterSettings.HorizontalCentered = $false}
        if($VerticalCentered){$WorkSheet.PrinterSettings.VerticalCentered = $true}else{$WorkSheet.PrinterSettings.VerticalCentered = $false}
        if($FitToPage){$WorkSheet.PrinterSettings.FitToPage = $true}else{$WorkSheet.PrinterSettings.FitToPage = $false}
        return [OfficeOpenXml.ExcelWorksheet]$WorkSheet
    }

}

Function Get-OOXMLHelp {
    <#
    .SYNOPSIS
    Display the full OOXML Module Help File
    
    .DESCRIPTION
    Display the full OOXML Module Help File

    .EXAMPLE
    Get-OOXMLHelp

    Description
    -----------
    Calls a function that will display the full OOXML Module Help File
    
    .NOTES
    
    .LINK 
    
    #>
    foreach($Command in $(Get-Command -Module excelpslib)){
        Get-Help -Name $($Command.Name)
    }
}

Function Get-OOXMLDeprecatedCommand {
    <#
    .SYNOPSIS
    Generate aliases for backward compatibility
    
    .DESCRIPTION
    Generate aliases for backward compatibility

    .EXAMPLE
    Get-OOXMLDeprecatedCommand

    Description
    -----------
    Calls a function that will generate aliases for backward compatibility
    
    .NOTES
    
    .LINK 
    
    #>
	Set-Alias CreateExcelInstance New-OOXMLPackage -Scope "Global"
	Set-Alias CreateWorkSheet Add-OOXMLWorksheet -Scope "Global"
	Set-Alias GetWorkBook Get-OOXMLWorkbook -Scope "Global"
	Set-Alias SelectWorkSheet Select-OOXMLWorkSheet -Scope "Global"
	Set-Alias SetValueAt Set-OOXMLRangeValue -Scope "Global"
	Set-Alias SetBorder Set-OOXMLRangeBorder -Scope "Global"
	Set-Alias SetBorderTop Set-OOXMLRangeBorderTop -Scope "Global"
	Set-Alias SetBorderRight Set-OOXMLRangeBorderRight -Scope "Global"
	Set-Alias SetBorderBottom Set-OOXMLRangeBorderBottom -Scope "Global"
	Set-Alias SetBorderLeft Set-OOXMLRangeBorderLeft -Scope "Global"
	Set-Alias SetFont Set-OOXMLRangeFont -Scope "Global"
	Set-Alias SetFill Set-OOXMLRangeFill -Scope "Global"
	Set-Alias SetTextOptions Set-OOXMLRangeTextOptions -Scope "Global"
	Set-Alias SetPrinterSettings Set-OOXMLPrinterSettings -Scope "Global"
	Set-Alias SaveFile Save-OOXMLPackage -Scope "Global"
    Set-Alias Get-ColumnString Get-OOXMLColumnString -Scope "Global"
}

Function Convert-OOXMLOldScripts {
    <#
    .SYNOPSIS
    Convert to the new format all the old commands
    
    .DESCRIPTION
    Convert to the new format all the old commands

    .PARAMETER InputFile
    The row decimal coordinate

    .PARAMETER OutputFile
    The column decimal coordinate

    .EXAMPLE
    Convert-OOXMLOldScripts -InputFile "C:\Old_Script.ps1" -OutpuFile "C:\New_Script.ps1"

    Description
    -----------
    Calls a function to convert to the new format all the old commands
    
    .NOTES
    
    .LINK 
    
    #>
    [CmdletBinding()]
	param(
		[parameter(Mandatory=$true)]
		[string]$InputFile,
		[parameter(Mandatory=$true)]
		[string]$OutputFile
	)
	process {
		$lookupTable = @{
			'CreateExcelInstance' = 'New-OOXMLPackage' 
			'CreateWorkSheet' = 'Add-OOXMLWorksheet' 
			'GetWorkBook' = 'Get-OOXMLWorkbook' 
			'SelectWorkSheet' = 'Select-OOXMLWorkSheet' 
			'SetValueAt' = 'Set-OOXMLRangeValue' 
			'SetBorder' = 'Set-OOXMLRangeBorder'
			'SetFont' = 'Set-OOXMLRangeFont'
			'SetFill' = 'Set-OOXMLRangeFill'
			'SetPrinterSettings' = 'Set-OOXMLPrinterSettings'
			'SetTextOptions' = 'Set-OOXMLRangeTextOptions'
            'Get-ColumnString' ='Get-OOXMLColumnString'
			'SaveFile' = 'Save-OOXMLPackage'
		}
		if(Test-Path -Path $inputFile){
			Get-Content -Path $inputFile | ForEach-Object { 
			    $line = $_
			    $lookupTable.GetEnumerator() | ForEach-Object {
			        if ($line -match $_.Key) {$line = $line -replace $_.Key, $_.Value}
			    }
			   $line
			} | Set-Content -Path $outputFile
		}
	}
}

Function Convert-OOXMLCellsCoordinates {
    <#
    .SYNOPSIS
    Convert decimal coordinate(s) to an excel coordinate style numbering
    
    .DESCRIPTION
    Convert decimal coordinate(s) to an excel coordinate style numbering

    .PARAMETER StartRow
    The row decimal coordinate

    .PARAMETER StartCol
    The column decimal coordinate

    .PARAMETER EndRow
    The row decimal coordinate

    .PARAMETER EndCol
    The column decimal coordinate

    .EXAMPLE
    $coordinates = Convert-OOXMLCellsCoordinates -StartRow 1 -StartCol 1 -EndRow 100 -EndCol 16

    Description
    -----------
    Calls a function to convert decimal coordinate(s) to an excel coordinate style numbering and return it
    
    .NOTES
    
    .LINK 
    
    #>
    [CmdletBinding(DefaultParametersetName="None")]
	param (

        [parameter(Mandatory=$true)]
		[int]$StartRow,

        [parameter(Mandatory=$true)]
		[int]$StartCol,

        [parameter(ParameterSetName="MultipleCells", Mandatory=$true)]
		[int]$EndRow,

        [parameter(ParameterSetName="MultipleCells", Mandatory=$true)]
		[int]$EndCol
	)
	process{
		[string]$Coordinates = [string]::Empty
		[string]$StartString = Get-OOXMLColumnString -ColNumber $StartCol
		$Coordinates = $StartString+$StartRow
		if($EndRow -and $EndCol){
			[string]$EndString = Get-OOXMLColumnString -ColNumber $EndCol
			$Coordinates = $Coordinates+":"+$EndString+$EndRow
		}
		return [string]$Coordinates
	}
}

Function Get-OOXMLColumnString {
    <#
    .SYNOPSIS
    Convert a decimal number to an excel column style numbering
    
    .DESCRIPTION
    Convert a decimal number to an excel column style numbering

    .PARAMETER ColNumber
    The decimal colomn number to be converted

    .EXAMPLE
    Get-OOXMLColumnString -ColNumber 34

    Description
    -----------
    Calls a function Convert a decimal number to an excel column style numbering and return it
    
    .NOTES
    
    .LINK 
    
    #>
    [CmdletBinding()]
	param (
        [ValidateScript({$_ -ge 1})]
		[int]$ColNumber
	)
    return [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($ColNumber)
}

Function New-OOXMLPivotTable {
    <#
    .SYNOPSIS
    Create a PivotTable in the given worksheet and return the created PivotTable
    
    .DESCRIPTION
    Create a PivotTable in the given worksheet and return the created PivotTable

    .PARAMETER WorkSheet
    Worksheet where are the data located

    .PARAMETER Origin
    Location of the upper left corner of the pivot table

    .PARAMETER Datas
    The data to be processed by the pivot table

    .PARAMETER Name
    The name of the pivot table

    .EXAMPLE
    New-OOXMLPivotTable -WorkSheet $sheet -origin "$A$932" -Datas "$A$1:$E$910" -Name "New Pivot Table"

    Description
    -----------
    Calls a function which create a PivotTable in the given worksheet and return an ExcelPivotTable object
    
    .NOTES
    This function is really basic and must be improved
    
    .LINK 
    
    #>
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]
		[OfficeOpenXml.ExcelWorksheet]$WorkSheet,
        [parameter(Mandatory=$true)]
        [string]$Origin = [string]::Empty,
        [parameter(Mandatory=$true)]
        [string]$Datas = [string]::Empty,
        [parameter(Mandatory=$true)]
        [string]$Name = [string]::Empty
    )
    process{
        [OfficeOpenXml.Table.PivotTable.ExcelPivotTable]$Pivot = $WorkSheet.PivotTables.Add($WorkSheet.Cells[$Origin],$WorkSheet.Cells[$Datas], $Name)
        return [OfficeOpenXml.Table.PivotTable.ExcelPivotTable]$Pivot
    }

}

Function Set-OOXMLStyleSheet {
    <#
    .SYNOPSIS
    Assign a style sheet to a range of cell
    
    .DESCRIPTION
    Assign a style sheet to a range of cell

    .PARAMETER CellRange
    The cell range that style sheet is to be applied

    .PARAMETER StyleSheet
    The style sheet object that you want to use

    .PARAMETER StyleSheetName
    The style sheet object name that you want to use


    .EXAMPLE
    $Range | Set-OOXMLStyleSheet -StyleSheet $Style
    Set-OOXMLStyleSheet -cellRange $Range -StyleSheet $Style

    Description
    -----------
    Calls a function which assign a style sheet to a range of cell and return an ExcelRangeBase object
    
    .NOTES
    
    .LINK 
    
    #>
    [CmdletBinding()]
    param(
        [parameter(ParameterSetName="WithStyleSheetObject", Mandatory=$true, ValueFromPipeline=$true)]
        [parameter(ParameterSetName="WithStyleSheetName", Mandatory=$true, ValueFromPipeline=$true)]
        [OfficeOpenXml.ExcelRangeBase]$cellRange,

        [parameter(ParameterSetName="WithStyleSheetObject", Mandatory=$true)]
        [OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml]$StyleSheet,

        [parameter(ParameterSetName="WithStyleSheetName", Mandatory=$true)]
        [string]$StyleSheetName
    )
    process{
        if($StyleSheet){
            $cellRange.StyleName = $StyleSheet.Name
        }else{
            $cellRange.StyleName = $StyleSheetName
        }
        return [OfficeOpenXml.ExcelRangeBase]$cellRange 
    }
}

Function New-OOXMLStyleSheet {
    <#
    .SYNOPSIS
    Create a style sheet object
    
    .DESCRIPTION
    Create a style sheet object

    .PARAMETER WorkBook
    The workbook in the Excel instance 

    .PARAMETER Name
    The name you want to give to your style sheet

    .PARAMETER HAlign
    Set the horizontal text alignement

    .PARAMETER VAlign
    Set the vertical alignement type

    .PARAMETER NFormat
    Format a number according to a definited patern

    .PARAMETER Wrap
    Force end of the for line that are bigger than the cell 

    .PARAMETER Shrink
    Reduce the size of the text to fit in cell

    .PARAMETER Locked
    Prevent text edition within a cell

    .PARAMETER Bold
    Set the font to bold

    .PARAMETER Italic
    Set the font to italic

    .PARAMETER Underline
    Set the font to underlined

    .PARAMETER Strike
    Set the font to striked

    .PARAMETER Size
    Set the font size

    .PARAMETER TextRotation
    Set the angle of the text

    .PARAMETER ForeGroundColor
    The color that will be applied to the text

    .PARAMETER FillType
    The type of fill style to use on the background

    .PARAMETER BackGroundColor
    The color that will be applied to the background

    .PARAMETER BorderStyle
    The border style that will be applied to the range of cell

    .PARAMETER BorderColor
    The color that will be applied to the border

    .EXAMPLE
    $Style = New-OOXMLStyleSheet -WorkBook $book -Name "FirstStyle" -FillType solid -HAlign Center -Italic -Size 14 -BackGroundColor Red -TextRotation 90
    $Style = $book | New-OOXMLStyleSheet -Name "FirstStyle" -FillType solid -HAlign Center -Italic -Size 14 -BackGroundColor Red

    Description
    -----------
    Calls a function which will create, configure and return an ExcelNamedStyleXml object
    
    .NOTES
    
    .LINK 
    
    #>
    [CmdletBinding(DefaultParametersetName="None")]
    param(
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]
		[OfficeOpenXml.ExcelWorkbook]$WorkBook,
        [parameter(Mandatory=$true)]  
        [string]$Name,
        [OfficeOpenXml.Style.ExcelHorizontalAlignment]$HAlign,
        [OfficeOpenXml.Style.ExcelVerticalAlignment]$VAlign,
        [string]$NFormat,
        [switch]$Wrap,
        [switch]$Shrink,
        [switch]$Locked,
        [switch]$Bold,
        [switch]$Italic,
        [switch]$Underline,
        [switch]$Strike,
        [float]$Size,
        [ValidateScript({($_ -ge 0) -and ($_ -le 180)})]
        [int]$TextRotation,
        [System.Drawing.Color]$ForeGroundColor,
        [OfficeOpenXml.Style.ExcelFillStyle]$FillType,
        [System.Drawing.Color]$BackGroundColor,
        [OfficeOpenXml.Style.ExcelBorderStyle]$borderStyle,
        [System.Drawing.Color]$BorderColor
    )
    process{

        if($WorkBook.Styles.NamedStyles.Name -notcontains $Name)
        {
            $StyleSheet = $WorkBook.Styles.CreateNamedStyle($Name)

            if($borderStyle){
                $StyleSheet.Style.Border.Left.Style = $borderStyle
                $StyleSheet.Style.Border.Bottom.Style = $borderStyle
                $StyleSheet.Style.Border.Right.Style = $borderStyle
                $StyleSheet.Style.Border.Top.Style = $borderStyle
            }

            if($BorderColor){
                $StyleSheet.Style.Border.Left.Color.SetColor($BorderColor)
                $StyleSheet.Style.Border.Bottom.Color.SetColor($BorderColor)
                $StyleSheet.Style.Border.Right.Color.SetColor($BorderColor)
                $StyleSheet.Style.Border.Top.Color.SetColor($BorderColor)
            }
        
            if($FillType){$StyleSheet.Style.Fill.PatternType = $FillType}
            if($BackGroundColor){$StyleSheet.Style.Fill.BackgroundColor.SetColor($BackGroundColor)}

            if($bold){$StyleSheet.Style.Font.Bold = $true}else{$StyleSheet.Style.Font.Bold = $false}
            if($italic){$StyleSheet.Style.Font.Italic = $true}else{$StyleSheet.Style.Font.Italic = $false}
            if($underline){$StyleSheet.Style.Font.UnderLine = $true}else{$StyleSheet.Style.Font.UnderLine = $false}
            if($strike){$StyleSheet.Style.Font.Strike = $true}else{$StyleSheet.Style.Font.Strike = $false}
            if($size){$StyleSheet.Style.Font.Size = $size}
            if($ForeGroundColor){$StyleSheet.Style.Font.Color.SetColor($ForeGroundColor)}

            if($HAlign){$StyleSheet.Style.HorizontalAlignment = $HAlign}
            if($VAlign){$StyleSheet.Style.VerticalAlignment = $VAlign}
            if($Wrap){$StyleSheet.Style.WrapText = $true}else{$StyleSheet.Style.WrapText = $false}
            if($Shrink){$StyleSheet.Style.ShrinkToFit = $true}else{$StyleSheet.Style.ShrinkToFit = $false}
            if($NFormat){$StyleSheet.Style.Numberformat.Format = $NFormat}
            if($Locked){$StyleSheet.Style.Locked = $true}else{$StyleSheet.Style.Locked = $false}
            if($TextRotation){$StyleSheet.Style.TextRotation = $TextRotation}
        }
        else
        {
            for($i=0;$i -lt $Book.Styles.NamedStyles.Count; $i++)
            {
                if($Book.Styles.NamedStyles[$i].Name -like $Name)
                {
                  return [OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml]$WorkBook.Styles.NamedStyles[$i]
                }
            }
        }

        return [OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml]$StyleSheet
    }
}

Function New-OOXMLStyleSheetData {
    <#
    .SYNOPSIS
    Create a style sheet data object
    
    .DESCRIPTION
    Create a style sheet data object

    .PARAMETER Name
    The name you want to give to your style sheet

    .PARAMETER HAlign
    Set the horizontal text alignement

    .PARAMETER VAlign
    Set the vertical alignement type

    .PARAMETER NFormat
    Format a number according to a definited patern

    .PARAMETER Wrap
    Force end of the for line that are bigger than the cell 

    .PARAMETER Shrink
    Reduce the size of the text to fit in cell

    .PARAMETER Locked
    Prevent text edition within a cell

    .PARAMETER Bold
    Set the font to bold

    .PARAMETER Italic
    Set the font to italic

    .PARAMETER Underline
    Set the font to underlined

    .PARAMETER Strike
    Set the font to striked

    .PARAMETER Size
    Set the font size

    .PARAMETER TextRotation
    Set the angle of the text

    .PARAMETER ForeGroundColor
    The color that will be applied to the text

    .PARAMETER FillType
    The type of fill style to use on the background

    .PARAMETER BackGroundColor
    The color that will be applied to the background

    .PARAMETER BorderStyle
    The border style that will be applied to the range of cell

    .PARAMETER BorderColor
    The color that will be applied to the border

    .EXAMPLE
    $Style = New-OOXMLStyleSheetData -Name "FirstStyle" -FillType solid -HAlign Center -Italic -Size 14 -BackGroundColor Red -TextRotation 90

    Description
    -----------
    Calls a function which will return an style data object to be used in Export-OOXML
    
    .NOTES
    
    .LINK 
    
    #>
    param(
        [parameter(Mandatory=$true)]  
        [string]$Name,
        [OfficeOpenXml.Style.ExcelHorizontalAlignment]$HAlign = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center,
        [OfficeOpenXml.Style.ExcelVerticalAlignment]$VAlign = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center,
        [string]$NFormat,
        [switch]$Wrap = $true,
        [switch]$Shrink = $false,
        [switch]$Locked = $false,
        [switch]$Bold = $false,
        [switch]$Italic = $false,
        [switch]$Underline = $false,
        [switch]$Strike = $false,
        [float]$Size = 14,
        [ValidateScript({($_ -ge 0) -and ($_ -le 180)})]
        [int]$TextRotation,
        [System.Drawing.Color]$ForeGroundColor = [System.Drawing.Color]::White,
        [OfficeOpenXml.Style.ExcelFillStyle]$FillType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid,
        [System.Drawing.Color]$BackGroundColor = [System.Drawing.Color]::Black,
        [OfficeOpenXml.Style.ExcelBorderStyle]$borderStyle = [OfficeOpenXml.Style.ExcelBorderStyle]::Thick,
        [System.Drawing.Color]$BorderColor = [System.Drawing.Color]::Black
    )
    process{
        return [pscustomobject]@{
            Name = $Name
            HAlign = $HAlign
            VAlign = $VAlign
            NFormat= $NFormat
            Wrap = $Wrap
            Shrink = $Shrink
            Locked = $Locked
            Bold = $Bold
            Italic = $Italic
            Underline = $Underline
            Strike = $Strike
            Size = $Size
            TextRotation = $TextRotation
            ForeGroundColor = $ForeGroundColor
            FillType = $FillType
            BackGroundColor = $BackGroundColor
            borderStyle = $borderStyle
            BorderColor = $BorderColor
        }
    }
}
 
Function Export-OOXML {
    <#
    .SYNOPSIS
    Export an array of objects to an XLSX File
    
    .DESCRIPTION
    Export an array of objects to an XLSX File

    .PARAMETER InputObject
    The array object that will be exported to an XLSX File

    .PARAMETER FileFullPath
    The full path of the XLSX File

    .PARAMETER DocumentName
    The name of the XLSX Document

    .PARAMETER WorksheetName
    The name of the worksheet in the XLSX Document

    .PARAMETER IncludedProperties
    An array object containing the names of the object properties you want to export. This list is ordered !!!

    .PARAMETER ConditionalFormatings
    The conditional formating you want to apply

    .PARAMETER FormulaObjects
    The formula you want to apply to a whole column

    .PARAMETER OrderedProperties
    This allow you to order the columns either Ascending or Descending

    .PARAMETER AutoFit
    Auto size the columns

    .PARAMETER HeaderStyle
    Set the style of the header to a predefinited style

    .PARAMETER HeaderTextRotation
    Set the orientation of the header

    .PARAMETER HeaderCustomStyles
    This parameter will change the style of one or more header You must give an array of Hashtable as argument : @{Name=<string>;Data=<PSCustomObject>}

    .PARAMETER Precise
    This parameter will toogle the conditional formating mode to the precise mode by formating only a given column
    in place of the entire row.

    .PARAMETER DataValidationLists
    This parameter shoud receive an array of objects in the format produced by the Get-OOXMLDataValidationCustomObject this will only add data on an
    addtional "REF_DATA" worksheet. This parameter must by used in combination with ....

    .PARAMETER DataValidationAssignements



    .EXAMPLE

        $p = Get-Process

        Import-Module -Name ExcelPSLib -Force

        $FirstList = Get-OOXMLDataValidationCustomObject -Name "FirstList" -Values @("Value001","Value002","Value003")
        $secondList = Get-OOXMLDataValidationCustomObject -Name "SecondList" -Values @("Value00X","Value00Y","Value00Z")

        $FirstAssignement = Get-OOXMLDataValidationAssignementCustomObject -DataValidationName "FirstList" -ColumnNames @("Name","Handles")
        $SecondAssignement = Get-OOXMLDataValidationAssignementCustomObject -DataValidationName "SecondList" -ColumnNames @("WS","VM")

        $Red = Get-OOXMLConditonalFormattingCustomObject -Name "Name" -Style Red -Condition ContainsText -Value "host"
        $Green = Get-OOXMLConditonalFormattingCustomObject -Name "Name" -Style Green -Condition ContainsText -Value "32"

        $FormulaOne = Get-OOXMLFormulaObject -Name "Handles" -Style Beige -Operation AVERAGE

        $HeaderStyle = New-OOXMLStyleSheetData -Name "HeaderDemoStyle" -FillType solid -HAlign Center -Italic -Size 14 -BackGroundColor Red -TextRotation 90

        Export-OOXML -InputObject $p `
                     -FileFullPath "C:\temp\datavalidationtestX.xlsx" `
                     -DataValidationLists @($FirstList,$secondList) `
                     -DataValidationAssignements @($FirstAssignement,$SecondAssignement) `
                     -FreezedColumnName "VM" `
                     -AutoFit `
                     -DocumentName "OOXMLDemo" `
                     -HeaderStyle Gray `
                     -WorksheetName "OOXMLDemo" `
                     -ConditionalFormatings @($Red,$Green) `
                     -FormulaObjects @($FormulaOne) `
                     -HeaderCustomStyles @(@{Name="Company";Data=$HeaderStyle}) `
                     -OrderedProperties Ascending

    Description
    -----------
    Calls a function that will export the content of an array to an XLSX file
    
    .NOTES
    
    .LINK 
    
    #>
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [object[]]$InputObject,
        [ValidateScript({Test-Path -Path $_ -PathType Leaf -IsValid})]
        [parameter(Mandatory=$true)]
        [string]$FileFullPath,
        [string]$DocumentName = "ExcelPSLib",
        [string]$WorksheetName = "ExcelPSLib",
        [string[]]$IncludedProperties,
        [string]$FreezedColumnName,
        [object[]]$ConditionalFormatings,
        [object[]]$FormulaObjects,
        [switch]$AutoFit,
        [ValidateSet("Ascending","Descending")]
        [string]$OrderedProperties,
        [ExcelPSLib.EnumColors]$HeaderStyle = [ExcelPSLib.EnumColors]::Black,
        [ValidateScript({($_ -ge 0) -and ($_ -le 180)})]
        [int]$HeaderTextRotation = 0,
        [object[]]$HeaderCustomStyles,
        [switch]$Precise,
        [parameter(ParameterSetName="DataValidation")]
        [object[]]$DataValidationLists,
        [parameter(ParameterSetName="DataValidation")]
        [object[]]$DataValidationAssignements,
        [object[]]$CustomFormatings,
        [switch]$AddToExistingDocument
    )
    process{

        try
        {
            if($IncludedProperties){
                
                $RawReferencePropertySet = $($InputObject[0].PSObject.Properties).Name
                [string[]]$ReferencePropertySet = @()

                foreach($IncludedProperty in $IncludedProperties)
                {
                    if($RawReferencePropertySet -contains $IncludedProperty)
                    {
                        $ReferencePropertySet += $IncludedProperty
                    }
                }
            }
            else
            {
                $ReferencePropertySet = $($InputObject[0].PSObject.Properties).Name
            }

            if($OrderedProperties -like "Ascending")
            {
                [System.Array]::Sort($ReferencePropertySet)
            }

            if($OrderedProperties -like "Descending")
            {
                [System.Array]::Sort($ReferencePropertySet)
                [System.Array]::Reverse($ReferencePropertySet)
            }

            $ColumnNumber = $ReferencePropertySet.Length

            $RowPosition = 2

            if($AddToExistingDocument)
            {
                [OfficeOpenXml.ExcelPackage]$excel = New-OOXMLPackage -author "ExcelPSLib" -title $DocumentName -Path $FileFullPath
                [OfficeOpenXml.ExcelWorkbook]$book = $excel | Get-OOXMLWorkbook

                $AutofilterRange = Convert-OOXMLCellsCoordinates -StartRow $RowPosition -EndRow $RowPosition -StartCol 1 -EndCol $ColumnNumber

                $excel | Add-OOXMLWorksheet -WorkSheetName $WorksheetName -AutofilterRange $AutofilterRange
                $sheet = $book | Select-OOXMLWorkSheet -WorkSheetName $WorksheetName
            }
            else
            {
                [OfficeOpenXml.ExcelPackage]$excel = New-OOXMLPackage -author "ExcelPSLib" -title $DocumentName
                [OfficeOpenXml.ExcelWorkbook]$book = $excel | Get-OOXMLWorkbook

                $AutofilterRange = Convert-OOXMLCellsCoordinates -StartRow $RowPosition -EndRow $RowPosition -StartCol 1 -EndCol $ColumnNumber

                $excel | Add-OOXMLWorksheet -WorkSheetName $WorksheetName -AutofilterRange $AutofilterRange
                $sheet = $book | Select-OOXMLWorkSheet -WorkSheetName $WorksheetName
            }
            
            
            
            $StyleHeaderCollection = @{
                "AliceBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "AliceBlueStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor AliceBlue -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "AntiqueWhite" = New-OOXMLStyleSheet -WorkBook $book -Name "AntiqueWhiteStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor AntiqueWhite -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Aqua" = New-OOXMLStyleSheet -WorkBook $book -Name "AquaStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Aqua -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Aquamarine" = New-OOXMLStyleSheet -WorkBook $book -Name "AquamarineStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Aquamarine -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Azure" = New-OOXMLStyleSheet -WorkBook $book -Name "AzureStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Azure -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Beige" = New-OOXMLStyleSheet -WorkBook $book -Name "BeigeStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Beige -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Bisque" = New-OOXMLStyleSheet -WorkBook $book -Name "BisqueStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Bisque -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Black" = New-OOXMLStyleSheet -WorkBook $book -Name "BlackStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Black -FillType Solid -ForeGroundColor White  -TextRotation $HeaderTextRotation
                "BlanchedAlmond" = New-OOXMLStyleSheet -WorkBook $book -Name "BlanchedAlmondStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor BlanchedAlmond -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Blue" = New-OOXMLStyleSheet -WorkBook $book -Name "BlueStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Blue -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "BlueViolet" = New-OOXMLStyleSheet -WorkBook $book -Name "BlueVioletStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor BlueViolet -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Brown" = New-OOXMLStyleSheet -WorkBook $book -Name "BrownStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Brown -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "BurlyWood" = New-OOXMLStyleSheet -WorkBook $book -Name "BurlyWoodStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor BurlyWood -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "CadetBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "CadetBlueStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor CadetBlue -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Chartreuse" = New-OOXMLStyleSheet -WorkBook $book -Name "ChartreuseStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Chartreuse -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Chocolate" = New-OOXMLStyleSheet -WorkBook $book -Name "ChocolateStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Chocolate -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Coral" = New-OOXMLStyleSheet -WorkBook $book -Name "CoralStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Coral -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "CornflowerBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "CornflowerBlueStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor CornflowerBlue -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Cornsilk" = New-OOXMLStyleSheet -WorkBook $book -Name "CornsilkStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Cornsilk -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Crimson" = New-OOXMLStyleSheet -WorkBook $book -Name "CrimsonStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Crimson -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Cyan" = New-OOXMLStyleSheet -WorkBook $book -Name "CyanStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Cyan -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "DarkBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkBlueStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor DarkBlue -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "DarkCyan" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkCyanStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor DarkCyan -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "DarkGoldenrod" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkGoldenrodStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor DarkGoldenrod -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "DarkGray" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkGrayStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor DarkGray -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "DarkGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkGreenStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor DarkGreen -FillType Solid -ForeGroundColor White -TextRotation $HeaderTextRotation
                "DarkKhaki" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkKhakiStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor DarkKhaki -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "DarkMagenta" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkMagentaStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor DarkMagenta -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "DarkOliveGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkOliveGreenStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor DarkOliveGreen -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "DarkOrange" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkOrangeStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor DarkOrange -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "DarkOrchid" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkOrchidStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor DarkOrchid -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "DarkRed" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkRedStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor DarkRed -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "DarkSalmon" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkSalmonStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor DarkSalmon -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "DarkSeaGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkSeaGreenStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor DarkSeaGreen -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "DarkSlateBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkSlateBlueStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor DarkSlateBlue -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "DarkSlateGray" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkSlateGrayStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor DarkSlateGray -FillType Solid -ForeGroundColor White -TextRotation $HeaderTextRotation
                "DarkTurquoise" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkTurquoiseStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor DarkTurquoise -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "DarkViolet" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkVioletStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor DarkViolet -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "DeepPink" = New-OOXMLStyleSheet -WorkBook $book -Name "DeepPinkStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor DeepPink -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "DeepSkyBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "DeepSkyBlueStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor DeepSkyBlue -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "DimGray" = New-OOXMLStyleSheet -WorkBook $book -Name "DimGrayStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor DimGray -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "DodgerBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "DodgerBlueStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor DodgerBlue -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Firebrick" = New-OOXMLStyleSheet -WorkBook $book -Name "FirebrickStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Firebrick -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "FloralWhite" = New-OOXMLStyleSheet -WorkBook $book -Name "FloralWhiteStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor FloralWhite -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "ForestGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "ForestGreenStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor ForestGreen -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Fuchsia" = New-OOXMLStyleSheet -WorkBook $book -Name "FuchsiaStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Fuchsia -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Gainsboro" = New-OOXMLStyleSheet -WorkBook $book -Name "GainsboroStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Gainsboro -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "GhostWhite" = New-OOXMLStyleSheet -WorkBook $book -Name "GhostWhiteStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor GhostWhite -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Gold" = New-OOXMLStyleSheet -WorkBook $book -Name "GoldStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Gold -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Goldenrod" = New-OOXMLStyleSheet -WorkBook $book -Name "GoldenrodStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Goldenrod -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Gray" = New-OOXMLStyleSheet -WorkBook $book -Name "GrayStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Gray -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Green" = New-OOXMLStyleSheet -WorkBook $book -Name "GreenStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Green -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "GreenYellow" = New-OOXMLStyleSheet -WorkBook $book -Name "GreenYellowStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor GreenYellow -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Honeydew" = New-OOXMLStyleSheet -WorkBook $book -Name "HoneydewStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Honeydew -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "HotPink" = New-OOXMLStyleSheet -WorkBook $book -Name "HotPinkStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor HotPink -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "IndianRed" = New-OOXMLStyleSheet -WorkBook $book -Name "IndianRedStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor IndianRed -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Indigo" = New-OOXMLStyleSheet -WorkBook $book -Name "IndigoStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Indigo -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Ivory" = New-OOXMLStyleSheet -WorkBook $book -Name "IvoryStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Ivory -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Khaki" = New-OOXMLStyleSheet -WorkBook $book -Name "KhakiStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Khaki -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Lavender" = New-OOXMLStyleSheet -WorkBook $book -Name "LavenderStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Lavender -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "LavenderBlush" = New-OOXMLStyleSheet -WorkBook $book -Name "LavenderBlushStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor LavenderBlush -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "LawnGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "LawnGreenStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor LawnGreen -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "LemonChiffon" = New-OOXMLStyleSheet -WorkBook $book -Name "LemonChiffonStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor LemonChiffon -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "LightBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "LightBlueStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor LightBlue -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "LightCoral" = New-OOXMLStyleSheet -WorkBook $book -Name "LightCoralStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor LightCoral -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "LightCyan" = New-OOXMLStyleSheet -WorkBook $book -Name "LightCyanStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor LightCyan -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "LightGoldenrodYellow" = New-OOXMLStyleSheet -WorkBook $book -Name "LightGoldenrodYellowStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor LightGoldenrodYellow -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "LightGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "LightGreenStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor LightGreen -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "LightGray" = New-OOXMLStyleSheet -WorkBook $book -Name "LightGrayStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor LightGray -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "LightPink" = New-OOXMLStyleSheet -WorkBook $book -Name "LightPinkStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor LightPink -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "LightSalmon" = New-OOXMLStyleSheet -WorkBook $book -Name "LightSalmonStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor LightSalmon -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "LightSeaGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "LightSeaGreenStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor LightSeaGreen -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "LightSkyBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "LightSkyBlueStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor LightSkyBlue -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "LightSlateGray" = New-OOXMLStyleSheet -WorkBook $book -Name "LightSlateGrayStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor LightSlateGray -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "LightSteelBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "LightSteelBlueStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor LightSteelBlue -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "LightYellow" = New-OOXMLStyleSheet -WorkBook $book -Name "LightYellowStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor LightYellow -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Lime" = New-OOXMLStyleSheet -WorkBook $book -Name "LimeStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Lime -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "LimeGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "LimeGreenStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor LimeGreen -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Linen" = New-OOXMLStyleSheet -WorkBook $book -Name "LinenStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Linen -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Magenta" = New-OOXMLStyleSheet -WorkBook $book -Name "MagentaStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Magenta -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Maroon" = New-OOXMLStyleSheet -WorkBook $book -Name "MaroonStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Maroon -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "MediumAquamarine" = New-OOXMLStyleSheet -WorkBook $book -Name "MediumAquamarineStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor MediumAquamarine -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "MediumBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "MediumBlueStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor MediumBlue -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "MediumOrchid" = New-OOXMLStyleSheet -WorkBook $book -Name "MediumOrchidStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor MediumOrchid -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "MediumPurple" = New-OOXMLStyleSheet -WorkBook $book -Name "MediumPurpleStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor MediumPurple -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "MediumSeaGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "MediumSeaGreenStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor MediumSeaGreen -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "MediumSlateBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "MediumSlateBlueStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor MediumSlateBlue -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "MediumSpringGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "MediumSpringGreenStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor MediumSpringGreen -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "MediumTurquoise" = New-OOXMLStyleSheet -WorkBook $book -Name "MediumTurquoiseStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor MediumTurquoise -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "MediumVioletRed" = New-OOXMLStyleSheet -WorkBook $book -Name "MediumVioletRedStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor MediumVioletRed -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "MidnightBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "MidnightBlueStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor MidnightBlue -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "MintCream" = New-OOXMLStyleSheet -WorkBook $book -Name "MintCreamStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor MintCream -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "MistyRose" = New-OOXMLStyleSheet -WorkBook $book -Name "MistyRoseStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor MistyRose -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Moccasin" = New-OOXMLStyleSheet -WorkBook $book -Name "MoccasinStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Moccasin -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "NavajoWhite" = New-OOXMLStyleSheet -WorkBook $book -Name "NavajoWhiteStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor NavajoWhite -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Navy" = New-OOXMLStyleSheet -WorkBook $book -Name "NavyStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Navy -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "OldLace" = New-OOXMLStyleSheet -WorkBook $book -Name "OldLaceStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor OldLace -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Olive" = New-OOXMLStyleSheet -WorkBook $book -Name "OliveStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Olive -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "OliveDrab" = New-OOXMLStyleSheet -WorkBook $book -Name "OliveDrabStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor OliveDrab -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Orange" = New-OOXMLStyleSheet -WorkBook $book -Name "OrangeStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Orange -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "OrangeRed" = New-OOXMLStyleSheet -WorkBook $book -Name "OrangeRedStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor OrangeRed -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Orchid" = New-OOXMLStyleSheet -WorkBook $book -Name "OrchidStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Orchid -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "PaleGoldenrod" = New-OOXMLStyleSheet -WorkBook $book -Name "PaleGoldenrodStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor PaleGoldenrod -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "PaleGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "PaleGreenStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor PaleGreen -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "PaleTurquoise" = New-OOXMLStyleSheet -WorkBook $book -Name "PaleTurquoiseStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor PaleTurquoise -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "PaleVioletRed" = New-OOXMLStyleSheet -WorkBook $book -Name "PaleVioletRedStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor PaleVioletRed -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "PapayaWhip" = New-OOXMLStyleSheet -WorkBook $book -Name "PapayaWhipStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor PapayaWhip -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "PeachPuff" = New-OOXMLStyleSheet -WorkBook $book -Name "PeachPuffStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor PeachPuff -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Peru" = New-OOXMLStyleSheet -WorkBook $book -Name "PeruStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Peru -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Pink" = New-OOXMLStyleSheet -WorkBook $book -Name "PinkStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Pink -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Plum" = New-OOXMLStyleSheet -WorkBook $book -Name "PlumStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Plum -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "PowderBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "PowderBlueStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor PowderBlue -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Purple" = New-OOXMLStyleSheet -WorkBook $book -Name "PurpleStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Purple -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Red" = New-OOXMLStyleSheet -WorkBook $book -Name "RedStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Red -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "RosyBrown" = New-OOXMLStyleSheet -WorkBook $book -Name "RosyBrownStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor RosyBrown -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "RoyalBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "RoyalBlueStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor RoyalBlue -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "SaddleBrown" = New-OOXMLStyleSheet -WorkBook $book -Name "SaddleBrownStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor SaddleBrown -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Salmon" = New-OOXMLStyleSheet -WorkBook $book -Name "SalmonStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Salmon -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "SandyBrown" = New-OOXMLStyleSheet -WorkBook $book -Name "SandyBrownStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor SandyBrown -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "SeaGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "SeaGreenStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor SeaGreen -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "SeaShell" = New-OOXMLStyleSheet -WorkBook $book -Name "SeaShellStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor SeaShell -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Sienna" = New-OOXMLStyleSheet -WorkBook $book -Name "SiennaStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Sienna -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Silver" = New-OOXMLStyleSheet -WorkBook $book -Name "SilverStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Silver -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "SkyBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "SkyBlueStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor SkyBlue -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "SlateBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "SlateBlueStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor SlateBlue -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "SlateGray" = New-OOXMLStyleSheet -WorkBook $book -Name "SlateGrayStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor SlateGray -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Snow" = New-OOXMLStyleSheet -WorkBook $book -Name "SnowStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Snow -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "SpringGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "SpringGreenStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor SpringGreen -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "SteelBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "SteelBlueStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor SteelBlue -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Tan" = New-OOXMLStyleSheet -WorkBook $book -Name "TanStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Tan -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Teal" = New-OOXMLStyleSheet -WorkBook $book -Name "TealStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Teal -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Thistle" = New-OOXMLStyleSheet -WorkBook $book -Name "ThistleStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Thistle -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Tomato" = New-OOXMLStyleSheet -WorkBook $book -Name "TomatoStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Tomato -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Turquoise" = New-OOXMLStyleSheet -WorkBook $book -Name "TurquoiseStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Turquoise -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Violet" = New-OOXMLStyleSheet -WorkBook $book -Name "VioletStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Violet -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Wheat" = New-OOXMLStyleSheet -WorkBook $book -Name "WheatStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Wheat -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "White" = New-OOXMLStyleSheet -WorkBook $book -Name "WhiteStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor White -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "WhiteSmoke" = New-OOXMLStyleSheet -WorkBook $book -Name "WhiteSmokeStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor WhiteSmoke -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "Yellow" = New-OOXMLStyleSheet -WorkBook $book -Name "YellowStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor Yellow -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
                "YellowGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "YellowGreenStyleHeader" -Size 14 -Bold -HAlign Center -VAlign Center -BackGroundColor YellowGreen -FillType Solid -ForeGroundColor Black -TextRotation $HeaderTextRotation
            }

            $StyleNormal = New-OOXMLStyleSheet -WorkBook $book -Name "NormalStyle" -borderStyle Thin -BorderColor Black -HAlign Right
            $StyleURI = New-OOXMLStyleSheet -WorkBook $book -Name "URIStyle" -borderStyle Thin -BorderColor Black -HAlign Left -ForeGroundColor Blue -Underline
        
            $StyleDate = New-OOXMLStyleSheet -WorkBook $book -Name "DateStyle" -borderStyle Thin -BorderColor Black -HAlign Right -NFormat "$([System.Globalization.DateTimeFormatInfo]::CurrentInfo.ShortDatePattern) $([System.Globalization.DateTimeFormatInfo]::CurrentInfo.ShortTimePattern)"
            $StyleNumber = New-OOXMLStyleSheet -WorkBook $book -Name "NumberStyle" -borderStyle Thin -BorderColor Black -HAlign Right -NFormat "0"
            $StyleFloat = New-OOXMLStyleSheet -WorkBook $book -Name "FloatStyle" -borderStyle Thin -BorderColor Black -HAlign Right -NFormat "0.00"

            $StyleCollection = @{
                "AliceBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "AliceBlueStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor AliceBlue -FillType Solid -ForeGroundColor Black
                "AntiqueWhite" = New-OOXMLStyleSheet -WorkBook $book -Name "AntiqueWhiteStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor AntiqueWhite -FillType Solid -ForeGroundColor Black
                "Aqua" = New-OOXMLStyleSheet -WorkBook $book -Name "AquaStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Aqua -FillType Solid -ForeGroundColor Black
                "Aquamarine" = New-OOXMLStyleSheet -WorkBook $book -Name "AquamarineStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Aquamarine -FillType Solid -ForeGroundColor Black
                "Azure" = New-OOXMLStyleSheet -WorkBook $book -Name "AzureStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Azure -FillType Solid -ForeGroundColor Black
                "Beige" = New-OOXMLStyleSheet -WorkBook $book -Name "BeigeStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Beige -FillType Solid -ForeGroundColor Black
                "Bisque" = New-OOXMLStyleSheet -WorkBook $book -Name "BisqueStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Bisque -FillType Solid -ForeGroundColor Black
                "Black" = New-OOXMLStyleSheet -WorkBook $book -Name "BlackStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Black -FillType Solid -ForeGroundColor White
                "BlanchedAlmond" = New-OOXMLStyleSheet -WorkBook $book -Name "BlanchedAlmondStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor BlanchedAlmond -FillType Solid -ForeGroundColor Black
                "Blue" = New-OOXMLStyleSheet -WorkBook $book -Name "BlueStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Blue -FillType Solid -ForeGroundColor Black
                "BlueViolet" = New-OOXMLStyleSheet -WorkBook $book -Name "BlueVioletStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor BlueViolet -FillType Solid -ForeGroundColor Black
                "Brown" = New-OOXMLStyleSheet -WorkBook $book -Name "BrownStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Brown -FillType Solid -ForeGroundColor Black
                "BurlyWood" = New-OOXMLStyleSheet -WorkBook $book -Name "BurlyWoodStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor BurlyWood -FillType Solid -ForeGroundColor Black
                "CadetBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "CadetBlueStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor CadetBlue -FillType Solid -ForeGroundColor Black
                "Chartreuse" = New-OOXMLStyleSheet -WorkBook $book -Name "ChartreuseStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Chartreuse -FillType Solid -ForeGroundColor Black
                "Chocolate" = New-OOXMLStyleSheet -WorkBook $book -Name "ChocolateStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Chocolate -FillType Solid -ForeGroundColor Black
                "Coral" = New-OOXMLStyleSheet -WorkBook $book -Name "CoralStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Coral -FillType Solid -ForeGroundColor Black
                "CornflowerBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "CornflowerBlueStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor CornflowerBlue -FillType Solid -ForeGroundColor Black
                "Cornsilk" = New-OOXMLStyleSheet -WorkBook $book -Name "CornsilkStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Cornsilk -FillType Solid -ForeGroundColor Black
                "Crimson" = New-OOXMLStyleSheet -WorkBook $book -Name "CrimsonStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Crimson -FillType Solid -ForeGroundColor Black
                "Cyan" = New-OOXMLStyleSheet -WorkBook $book -Name "CyanStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Cyan -FillType Solid -ForeGroundColor Black
                "DarkBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkBlueStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor DarkBlue -FillType Solid -ForeGroundColor Black
                "DarkCyan" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkCyanStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor DarkCyan -FillType Solid -ForeGroundColor Black
                "DarkGoldenrod" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkGoldenrodStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor DarkGoldenrod -FillType Solid -ForeGroundColor Black
                "DarkGray" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkGrayStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor DarkGray -FillType Solid -ForeGroundColor Black
                "DarkGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkGreenStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor DarkGreen -FillType Solid -ForeGroundColor White
                "DarkKhaki" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkKhakiStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor DarkKhaki -FillType Solid -ForeGroundColor Black
                "DarkMagenta" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkMagentaStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor DarkMagenta -FillType Solid -ForeGroundColor Black
                "DarkOliveGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkOliveGreenStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor DarkOliveGreen -FillType Solid -ForeGroundColor Black
                "DarkOrange" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkOrangeStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor DarkOrange -FillType Solid -ForeGroundColor Black
                "DarkOrchid" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkOrchidStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor DarkOrchid -FillType Solid -ForeGroundColor Black
                "DarkRed" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkRedStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor DarkRed -FillType Solid -ForeGroundColor Black
                "DarkSalmon" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkSalmonStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor DarkSalmon -FillType Solid -ForeGroundColor Black
                "DarkSeaGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkSeaGreenStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor DarkSeaGreen -FillType Solid -ForeGroundColor Black
                "DarkSlateBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkSlateBlueStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor DarkSlateBlue -FillType Solid -ForeGroundColor Black
                "DarkSlateGray" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkSlateGrayStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor DarkSlateGray -FillType Solid -ForeGroundColor White
                "DarkTurquoise" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkTurquoiseStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor DarkTurquoise -FillType Solid -ForeGroundColor Black
                "DarkViolet" = New-OOXMLStyleSheet -WorkBook $book -Name "DarkVioletStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor DarkViolet -FillType Solid -ForeGroundColor Black
                "DeepPink" = New-OOXMLStyleSheet -WorkBook $book -Name "DeepPinkStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor DeepPink -FillType Solid -ForeGroundColor Black
                "DeepSkyBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "DeepSkyBlueStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor DeepSkyBlue -FillType Solid -ForeGroundColor Black
                "DimGray" = New-OOXMLStyleSheet -WorkBook $book -Name "DimGrayStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor DimGray -FillType Solid -ForeGroundColor Black
                "DodgerBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "DodgerBlueStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor DodgerBlue -FillType Solid -ForeGroundColor Black
                "Firebrick" = New-OOXMLStyleSheet -WorkBook $book -Name "FirebrickStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Firebrick -FillType Solid -ForeGroundColor Black
                "FloralWhite" = New-OOXMLStyleSheet -WorkBook $book -Name "FloralWhiteStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor FloralWhite -FillType Solid -ForeGroundColor Black
                "ForestGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "ForestGreenStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor ForestGreen -FillType Solid -ForeGroundColor Black
                "Fuchsia" = New-OOXMLStyleSheet -WorkBook $book -Name "FuchsiaStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Fuchsia -FillType Solid -ForeGroundColor Black
                "Gainsboro" = New-OOXMLStyleSheet -WorkBook $book -Name "GainsboroStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Gainsboro -FillType Solid -ForeGroundColor Black
                "GhostWhite" = New-OOXMLStyleSheet -WorkBook $book -Name "GhostWhiteStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor GhostWhite -FillType Solid -ForeGroundColor Black
                "Gold" = New-OOXMLStyleSheet -WorkBook $book -Name "GoldStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Gold -FillType Solid -ForeGroundColor Black
                "Goldenrod" = New-OOXMLStyleSheet -WorkBook $book -Name "GoldenrodStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Goldenrod -FillType Solid -ForeGroundColor Black
                "Gray" = New-OOXMLStyleSheet -WorkBook $book -Name "GrayStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Gray -FillType Solid -ForeGroundColor Black
                "Green" = New-OOXMLStyleSheet -WorkBook $book -Name "GreenStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Green -FillType Solid -ForeGroundColor Black
                "GreenYellow" = New-OOXMLStyleSheet -WorkBook $book -Name "GreenYellowStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor GreenYellow -FillType Solid -ForeGroundColor Black
                "Honeydew" = New-OOXMLStyleSheet -WorkBook $book -Name "HoneydewStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Honeydew -FillType Solid -ForeGroundColor Black
                "HotPink" = New-OOXMLStyleSheet -WorkBook $book -Name "HotPinkStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor HotPink -FillType Solid -ForeGroundColor Black
                "IndianRed" = New-OOXMLStyleSheet -WorkBook $book -Name "IndianRedStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor IndianRed -FillType Solid -ForeGroundColor Black
                "Indigo" = New-OOXMLStyleSheet -WorkBook $book -Name "IndigoStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Indigo -FillType Solid -ForeGroundColor Black
                "Ivory" = New-OOXMLStyleSheet -WorkBook $book -Name "IvoryStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Ivory -FillType Solid -ForeGroundColor Black
                "Khaki" = New-OOXMLStyleSheet -WorkBook $book -Name "KhakiStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Khaki -FillType Solid -ForeGroundColor Black
                "Lavender" = New-OOXMLStyleSheet -WorkBook $book -Name "LavenderStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Lavender -FillType Solid -ForeGroundColor Black
                "LavenderBlush" = New-OOXMLStyleSheet -WorkBook $book -Name "LavenderBlushStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor LavenderBlush -FillType Solid -ForeGroundColor Black
                "LawnGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "LawnGreenStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor LawnGreen -FillType Solid -ForeGroundColor Black
                "LemonChiffon" = New-OOXMLStyleSheet -WorkBook $book -Name "LemonChiffonStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor LemonChiffon -FillType Solid -ForeGroundColor Black
                "LightBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "LightBlueStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor LightBlue -FillType Solid -ForeGroundColor Black
                "LightCoral" = New-OOXMLStyleSheet -WorkBook $book -Name "LightCoralStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor LightCoral -FillType Solid -ForeGroundColor Black
                "LightCyan" = New-OOXMLStyleSheet -WorkBook $book -Name "LightCyanStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor LightCyan -FillType Solid -ForeGroundColor Black
                "LightGoldenrodYellow" = New-OOXMLStyleSheet -WorkBook $book -Name "LightGoldenrodYellowStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor LightGoldenrodYellow -FillType Solid -ForeGroundColor Black
                "LightGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "LightGreenStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor LightGreen -FillType Solid -ForeGroundColor Black
                "LightGray" = New-OOXMLStyleSheet -WorkBook $book -Name "LightGrayStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor LightGray -FillType Solid -ForeGroundColor Black
                "LightPink" = New-OOXMLStyleSheet -WorkBook $book -Name "LightPinkStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor LightPink -FillType Solid -ForeGroundColor Black
                "LightSalmon" = New-OOXMLStyleSheet -WorkBook $book -Name "LightSalmonStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor LightSalmon -FillType Solid -ForeGroundColor Black
                "LightSeaGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "LightSeaGreenStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor LightSeaGreen -FillType Solid -ForeGroundColor Black
                "LightSkyBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "LightSkyBlueStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor LightSkyBlue -FillType Solid -ForeGroundColor Black
                "LightSlateGray" = New-OOXMLStyleSheet -WorkBook $book -Name "LightSlateGrayStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor LightSlateGray -FillType Solid -ForeGroundColor Black
                "LightSteelBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "LightSteelBlueStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor LightSteelBlue -FillType Solid -ForeGroundColor Black
                "LightYellow" = New-OOXMLStyleSheet -WorkBook $book -Name "LightYellowStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor LightYellow -FillType Solid -ForeGroundColor Black
                "Lime" = New-OOXMLStyleSheet -WorkBook $book -Name "LimeStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Lime -FillType Solid -ForeGroundColor Black
                "LimeGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "LimeGreenStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor LimeGreen -FillType Solid -ForeGroundColor Black
                "Linen" = New-OOXMLStyleSheet -WorkBook $book -Name "LinenStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Linen -FillType Solid -ForeGroundColor Black
                "Magenta" = New-OOXMLStyleSheet -WorkBook $book -Name "MagentaStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Magenta -FillType Solid -ForeGroundColor Black
                "Maroon" = New-OOXMLStyleSheet -WorkBook $book -Name "MaroonStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Maroon -FillType Solid -ForeGroundColor Black
                "MediumAquamarine" = New-OOXMLStyleSheet -WorkBook $book -Name "MediumAquamarineStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor MediumAquamarine -FillType Solid -ForeGroundColor Black
                "MediumBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "MediumBlueStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor MediumBlue -FillType Solid -ForeGroundColor Black
                "MediumOrchid" = New-OOXMLStyleSheet -WorkBook $book -Name "MediumOrchidStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor MediumOrchid -FillType Solid -ForeGroundColor Black
                "MediumPurple" = New-OOXMLStyleSheet -WorkBook $book -Name "MediumPurpleStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor MediumPurple -FillType Solid -ForeGroundColor Black
                "MediumSeaGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "MediumSeaGreenStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor MediumSeaGreen -FillType Solid -ForeGroundColor Black
                "MediumSlateBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "MediumSlateBlueStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor MediumSlateBlue -FillType Solid -ForeGroundColor Black
                "MediumSpringGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "MediumSpringGreenStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor MediumSpringGreen -FillType Solid -ForeGroundColor Black
                "MediumTurquoise" = New-OOXMLStyleSheet -WorkBook $book -Name "MediumTurquoiseStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor MediumTurquoise -FillType Solid -ForeGroundColor Black
                "MediumVioletRed" = New-OOXMLStyleSheet -WorkBook $book -Name "MediumVioletRedStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor MediumVioletRed -FillType Solid -ForeGroundColor Black
                "MidnightBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "MidnightBlueStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor MidnightBlue -FillType Solid -ForeGroundColor Black
                "MintCream" = New-OOXMLStyleSheet -WorkBook $book -Name "MintCreamStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor MintCream -FillType Solid -ForeGroundColor Black
                "MistyRose" = New-OOXMLStyleSheet -WorkBook $book -Name "MistyRoseStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor MistyRose -FillType Solid -ForeGroundColor Black
                "Moccasin" = New-OOXMLStyleSheet -WorkBook $book -Name "MoccasinStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Moccasin -FillType Solid -ForeGroundColor Black
                "NavajoWhite" = New-OOXMLStyleSheet -WorkBook $book -Name "NavajoWhiteStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor NavajoWhite -FillType Solid -ForeGroundColor Black
                "Navy" = New-OOXMLStyleSheet -WorkBook $book -Name "NavyStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Navy -FillType Solid -ForeGroundColor Black
                "OldLace" = New-OOXMLStyleSheet -WorkBook $book -Name "OldLaceStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor OldLace -FillType Solid -ForeGroundColor Black
                "Olive" = New-OOXMLStyleSheet -WorkBook $book -Name "OliveStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Olive -FillType Solid -ForeGroundColor Black
                "OliveDrab" = New-OOXMLStyleSheet -WorkBook $book -Name "OliveDrabStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor OliveDrab -FillType Solid -ForeGroundColor Black
                "Orange" = New-OOXMLStyleSheet -WorkBook $book -Name "OrangeStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Orange -FillType Solid -ForeGroundColor Black
                "OrangeRed" = New-OOXMLStyleSheet -WorkBook $book -Name "OrangeRedStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor OrangeRed -FillType Solid -ForeGroundColor Black
                "Orchid" = New-OOXMLStyleSheet -WorkBook $book -Name "OrchidStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Orchid -FillType Solid -ForeGroundColor Black
                "PaleGoldenrod" = New-OOXMLStyleSheet -WorkBook $book -Name "PaleGoldenrodStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor PaleGoldenrod -FillType Solid -ForeGroundColor Black
                "PaleGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "PaleGreenStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor PaleGreen -FillType Solid -ForeGroundColor Black
                "PaleTurquoise" = New-OOXMLStyleSheet -WorkBook $book -Name "PaleTurquoiseStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor PaleTurquoise -FillType Solid -ForeGroundColor Black
                "PaleVioletRed" = New-OOXMLStyleSheet -WorkBook $book -Name "PaleVioletRedStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor PaleVioletRed -FillType Solid -ForeGroundColor Black
                "PapayaWhip" = New-OOXMLStyleSheet -WorkBook $book -Name "PapayaWhipStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor PapayaWhip -FillType Solid -ForeGroundColor Black
                "PeachPuff" = New-OOXMLStyleSheet -WorkBook $book -Name "PeachPuffStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor PeachPuff -FillType Solid -ForeGroundColor Black
                "Peru" = New-OOXMLStyleSheet -WorkBook $book -Name "PeruStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Peru -FillType Solid -ForeGroundColor Black
                "Pink" = New-OOXMLStyleSheet -WorkBook $book -Name "PinkStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Pink -FillType Solid -ForeGroundColor Black
                "Plum" = New-OOXMLStyleSheet -WorkBook $book -Name "PlumStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Plum -FillType Solid -ForeGroundColor Black
                "PowderBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "PowderBlueStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor PowderBlue -FillType Solid -ForeGroundColor Black
                "Purple" = New-OOXMLStyleSheet -WorkBook $book -Name "PurpleStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Purple -FillType Solid -ForeGroundColor Black
                "Red" = New-OOXMLStyleSheet -WorkBook $book -Name "RedStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Red -FillType Solid -ForeGroundColor Black
                "RosyBrown" = New-OOXMLStyleSheet -WorkBook $book -Name "RosyBrownStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor RosyBrown -FillType Solid -ForeGroundColor Black
                "RoyalBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "RoyalBlueStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor RoyalBlue -FillType Solid -ForeGroundColor Black
                "SaddleBrown" = New-OOXMLStyleSheet -WorkBook $book -Name "SaddleBrownStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor SaddleBrown -FillType Solid -ForeGroundColor Black
                "Salmon" = New-OOXMLStyleSheet -WorkBook $book -Name "SalmonStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Salmon -FillType Solid -ForeGroundColor Black
                "SandyBrown" = New-OOXMLStyleSheet -WorkBook $book -Name "SandyBrownStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor SandyBrown -FillType Solid -ForeGroundColor Black
                "SeaGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "SeaGreenStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor SeaGreen -FillType Solid -ForeGroundColor Black
                "SeaShell" = New-OOXMLStyleSheet -WorkBook $book -Name "SeaShellStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor SeaShell -FillType Solid -ForeGroundColor Black
                "Sienna" = New-OOXMLStyleSheet -WorkBook $book -Name "SiennaStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Sienna -FillType Solid -ForeGroundColor Black
                "Silver" = New-OOXMLStyleSheet -WorkBook $book -Name "SilverStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Silver -FillType Solid -ForeGroundColor Black
                "SkyBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "SkyBlueStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor SkyBlue -FillType Solid -ForeGroundColor Black
                "SlateBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "SlateBlueStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor SlateBlue -FillType Solid -ForeGroundColor Black
                "SlateGray" = New-OOXMLStyleSheet -WorkBook $book -Name "SlateGrayStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor SlateGray -FillType Solid -ForeGroundColor Black
                "Snow" = New-OOXMLStyleSheet -WorkBook $book -Name "SnowStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Snow -FillType Solid -ForeGroundColor Black
                "SpringGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "SpringGreenStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor SpringGreen -FillType Solid -ForeGroundColor Black
                "SteelBlue" = New-OOXMLStyleSheet -WorkBook $book -Name "SteelBlueStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor SteelBlue -FillType Solid -ForeGroundColor Black
                "Tan" = New-OOXMLStyleSheet -WorkBook $book -Name "TanStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Tan -FillType Solid -ForeGroundColor Black
                "Teal" = New-OOXMLStyleSheet -WorkBook $book -Name "TealStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Teal -FillType Solid -ForeGroundColor Black
                "Thistle" = New-OOXMLStyleSheet -WorkBook $book -Name "ThistleStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Thistle -FillType Solid -ForeGroundColor Black
                "Tomato" = New-OOXMLStyleSheet -WorkBook $book -Name "TomatoStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Tomato -FillType Solid -ForeGroundColor Black
                "Turquoise" = New-OOXMLStyleSheet -WorkBook $book -Name "TurquoiseStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Turquoise -FillType Solid -ForeGroundColor Black
                "Violet" = New-OOXMLStyleSheet -WorkBook $book -Name "VioletStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Violet -FillType Solid -ForeGroundColor Black
                "Wheat" = New-OOXMLStyleSheet -WorkBook $book -Name "WheatStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Wheat -FillType Solid -ForeGroundColor Black
                "White" = New-OOXMLStyleSheet -WorkBook $book -Name "WhiteStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor White -FillType Solid -ForeGroundColor Black
                "WhiteSmoke" = New-OOXMLStyleSheet -WorkBook $book -Name "WhiteSmokeStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor WhiteSmoke -FillType Solid -ForeGroundColor Black
                "Yellow" = New-OOXMLStyleSheet -WorkBook $book -Name "YellowStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor Yellow -FillType Solid -ForeGroundColor Black
                "YellowGreen" = New-OOXMLStyleSheet -WorkBook $book -Name "YellowGreenStyle" -borderStyle Thin -BorderColor Black -HAlign Right -BackGroundColor YellowGreen -FillType Solid -ForeGroundColor Black
            }
        
            $i=1
            $AssociatedConditionalFormattings = @()
            $AssociatedFormulaInformations = @()

            $DefaultHeaderStyle = $StyleHeaderCollection[$HeaderStyle.ToString()]

            foreach($Property in $ReferencePropertySet)
            {
                $StyleHeader = $DefaultHeaderStyle
                foreach($HeaderCustomStyle in $HeaderCustomStyles)
                {
                    if($($HeaderCustomStyle.Name) -eq $Property){
                        
                        $StyleHeader = New-OOXMLStyleSheet -WorkBook $book `
                                                           -Name $HeaderCustomStyle.Data.Name `
                                                           -HAlign $HeaderCustomStyle.Data.HAlign `
                                                           -VAlign $HeaderCustomStyle.Data.VAlign `
                                                           -NFormat $HeaderCustomStyle.Data.NFormat `
                                                           -Wrap:$($HeaderCustomStyle.Data.Wrap) `
                                                           -Shrink:$($HeaderCustomStyle.Data.Shrink) `
                                                           -Locked:$($HeaderCustomStyle.Data.Locked) `
                                                           -Bold:$($HeaderCustomStyle.Data.Bold) `
                                                           -Italic:$($HeaderCustomStyle.Data.Italic) `
                                                           -Underline:$($HeaderCustomStyle.Data.Underline) `
                                                           -Strike:$($HeaderCustomStyle.Data.Strike) `
                                                           -Size $HeaderCustomStyle.Data.Size `
                                                           -TextRotation $HeaderCustomStyle.Data.TextRotation `
                                                           -ForeGroundColor $HeaderCustomStyle.Data.ForeGroundColor `
                                                           -FillType $HeaderCustomStyle.Data.FillType `
                                                           -BackGroundColor $HeaderCustomStyle.Data.BackGroundColor `
                                                           -borderStyle $HeaderCustomStyle.Data.borderStyle `
                                                           -BorderColor $HeaderCustomStyle.Data.BorderColor
                    }
                }

                $sheet | Set-OOXMLRangeValue -row $RowPosition -col $i -value $Property -StyleSheet $StyleHeader | Out-Null
                $sheet.Column($i).Width = 32

                foreach($FormulaObject in $FormulaObjects)
                {
                    if($Property -eq $FormulaObject.Name)
                    {
                        $AssociatedFormulaInformations += [PSCustomObject]@{
                            FormulaObject = $FormulaObject
                            ColumnName = $Property
                            ColumnIndex = $i
                        }
                    }
                }

                foreach($ConditionalFormating in $ConditionalFormatings)
                {
                    if($Property -eq $ConditionalFormating.Name)
                    {
                        $AssociatedConditionalFormattings += [PSCustomObject]@{
                            FormattingObject = $ConditionalFormating
                            ColumnName = $Property
                            ColumnIndex = $i
                        }
                    }
                }
                $i++
            }

            $RowPosition++

            foreach($Object in $InputObject){
                $i=1
                foreach($Property in $ReferencePropertySet){
                    $Value = "Empty Value"
                    $IsURI = $false
                    if($($Object.$Property) -ne $null){
                        $Value = $($Object.$Property)
                    
                    }
                    $AppliedStyle = $StyleNormal
                    switch -regex ($($Value.GetType())){
                        "(^uint[0-9]{2}$)|(^int[0-9]{2}$)|(^long$)|(^int$)" {
                            $AppliedStyle = $StyleNumber
                        }
                        "(double)|(float)|(decimal)" {
                            $AppliedStyle = $StyleFloat
                        }
                        "datetime" {
                            $AppliedStyle = $StyleDate
                        }
                        "^string$"{
                            if($([System.URI]::IsWellFormedUriString([System.URI]::EscapeUriString($Value),[System.UriKind]::Absolute)) -and $($Value -match "(^\\\\)|(^http://)|(^ftp://)|(^[a-zA-Z]:(//|\\))|(^https://)"))
                            {
                                $AppliedStyle = $StyleURI
                                $IsURI = $true
                            }
                        }
                    }

                    if($IsURI){
                        $sheet | Set-OOXMLRangeValue -Row $RowPosition -Col $i -Value $Value -StyleSheet $AppliedStyle -Uri | Out-Null
                    }else{
                        $sheet | Set-OOXMLRangeValue -Row $RowPosition -Col $i -Value $Value -StyleSheet $AppliedStyle | Out-Null
                    }
                    $i++
                }
                $RowPosition++
            }

            $LastRow = $($RowPosition - 1)
            $FirstDataRowIndex = $($Sheet.Dimension.Start.Row + 1)
            
            $StartColumn = Get-OOXMLColumnString -ColNumber $($Sheet.Dimension.Start.Column)
            $EndColumn = Get-OOXMLColumnString -ColNumber $($Sheet.Dimension.End.Column)

            foreach($AssociatedFormulaInformation in $AssociatedFormulaInformations)
            {
                $FormulaColumnName = Get-OOXMLColumnString -ColNumber $($AssociatedFormulaInformation.ColumnIndex)
                $FormulaRangeAddress = $("$FormulaColumnName" + "$FirstDataRowIndex" + ":" + "$FormulaColumnName$LastRow")
                $FormulaAddress = $("$FormulaColumnName$RowPosition")

                Switch($AssociatedFormulaInformation.FormulaObject.Operation)
                {
                    "SUM" {
                        $Sheet.Cells[$FormulaAddress].Formula = "=SUM($FormulaRangeAddress)"
                    }
                    "SUMIF" {
                        $Sheet.Cells[$FormulaAddress].Formula = "=SUMIF($FormulaRangeAddress,`"$($AssociatedFormulaInformation.FormulaObject.Criteria)`")"
                    }
                    "AVERAGE" {
                        $Sheet.Cells[$FormulaAddress].Formula = "=AVERAGE($FormulaRangeAddress)"
                    }
                    "AVERAGEIF" {
                        $Sheet.Cells[$FormulaAddress].Formula = "=AVERAGEIF($FormulaRangeAddress,`"$($AssociatedFormulaInformation.FormulaObject.Criteria)`")"
                    }
                    "COUNT" {
                        $Sheet.Cells[$FormulaAddress].Formula = "=COUNT($FormulaRangeAddress)"
                    }
                    "COUNTIF" {
                        $Sheet.Cells[$FormulaAddress].Formula = "=COUNTIF($FormulaRangeAddress,`"$($AssociatedFormulaInformation.FormulaObject.Criteria)`")"
                    }
                    "MAX" {
                        $Sheet.Cells[$FormulaAddress].Formula = "=MAX($FormulaRangeAddress)"
                    }
                    "MIN" {
                        $Sheet.Cells[$FormulaAddress].Formula = "=MAX($FormulaRangeAddress)"
                    }
                }

                $FormulaObjectStyle = $StyleCollection[$AssociatedFormulaInformation.FormulaObject.Style]
                $sheet.Cells[$FormulaAddress].StyleName = $FormulaObjectStyle.Name
            }

            foreach($AssociatedConditionalFormatting in $AssociatedConditionalFormattings)
            {
            
                $ColumnName = Get-OOXMLColumnString -ColNumber $($AssociatedConditionalFormatting.ColumnIndex)
            
                $Address = $("$ColumnName" + "$FirstDataRowIndex" + ":" + "$ColumnName" + "$LastRow")
                $AddressWide = $("$StartColumn" + "$FirstDataRowIndex" + ":" + "$EndColumn" + "$LastRow")

                if($Precise)
                {
                    $sheet | Add-OOXMLConditionalFormatting -Addresses $Address -RuleType $($AssociatedConditionalFormatting.FormattingObject.Condition) -StyleSheet $StyleCollection[$AssociatedConditionalFormatting.FormattingObject.Style] -ConditionValue $($AssociatedConditionalFormatting.FormattingObject.Value)
                }
                else
                {
                    $Expression = [string]::Empty
                    $StringEscape = [string]::Empty

                    switch -regex ($($($AssociatedConditionalFormatting.FormattingObject.Value).GetType())){
                        "^string$"{
                            $StringEscape = "`""
                        }
                    }

                    switch($($AssociatedConditionalFormatting.FormattingObject.Condition))
                    {
                         $([ExcelPSLib.EnumConditionalFormattingRuleType]::GreaterThan)
                         {
                            $Symbol = ">"
                            $Expression = "$" + "$ColumnName" + "$FirstDataRowIndex" + $Symbol + $StringEscape + $($AssociatedConditionalFormatting.FormattingObject.Value) + $StringEscape
                         }
                         $([ExcelPSLib.EnumConditionalFormattingRuleType]::GreaterThanOrEqual)
                         {
                            $Symbol = ">="
                            $Expression = "$" + "$ColumnName" + "$FirstDataRowIndex" + $Symbol + $StringEscape + $($AssociatedConditionalFormatting.FormattingObject.Value) + $StringEscape
                         }
                         $([ExcelPSLib.EnumConditionalFormattingRuleType]::LessThan)
                         {
                            $Symbol = "<"
                            $Expression = "$" + "$ColumnName" + "$FirstDataRowIndex" + $Symbol + $StringEscape + $($AssociatedConditionalFormatting.FormattingObject.Value) + $StringEscape
                         }
                         $([ExcelPSLib.EnumConditionalFormattingRuleType]::LessThanOrEqual)
                         {
                            $Symbol = "<="
                            $Expression = "$" + "$ColumnName" + "$FirstDataRowIndex" + $Symbol + $StringEscape + $($AssociatedConditionalFormatting.FormattingObject.Value) + $StringEscape
                         }
                         $([ExcelPSLib.EnumConditionalFormattingRuleType]::NotEqual)
                         {
                            $Symbol = "<>"
                            $Expression = "$" + "$ColumnName" + "$FirstDataRowIndex" + $Symbol + $StringEscape + $($AssociatedConditionalFormatting.FormattingObject.Value) + $StringEscape
                         }
                         $([ExcelPSLib.EnumConditionalFormattingRuleType]::Equal)
                         {
                            $Symbol = "="
                            $Expression = "$" + "$ColumnName" + "$FirstDataRowIndex" + $Symbol + $StringEscape + $($AssociatedConditionalFormatting.FormattingObject.Value) + $StringEscape
                         }
                         $([ExcelPSLib.EnumConditionalFormattingRuleType]::ContainsText)
                         {
                            $Expression = "IF(COUNTIF($" + "$ColumnName" + "$FirstDataRowIndex" + "," + $StringEscape + "*" + $($AssociatedConditionalFormatting.FormattingObject.Value) + "*" + $StringEscape + ") > 0,TRUE,FALSE)"
                         }
                         $([ExcelPSLib.EnumConditionalFormattingRuleType]::BeginsWith)
                         {
                            $Expression = "IF(COUNTIF($" + "$ColumnName" + "$FirstDataRowIndex" + "," + $StringEscape + $($AssociatedConditionalFormatting.FormattingObject.Value) + "*" + $StringEscape + ") > 0,TRUE,FALSE)"
                         }
                         $([ExcelPSLib.EnumConditionalFormattingRuleType]::EndsWith)
                         {
                            $Expression = "IF(COUNTIF($" + "$ColumnName" + "$FirstDataRowIndex" + "," + $StringEscape + "*" + $($AssociatedConditionalFormatting.FormattingObject.Value) + $StringEscape + ") > 0,TRUE,FALSE)"
                         }
                    }

                    Write-Host $Expression

                    if($Expression.Length -gt 0)
                    {
                        $sheet | Add-OOXMLConditionalFormatting -Addresses $AddressWide -RuleType $([ExcelPSLib.EnumConditionalFormattingRuleType]::Expression) -StyleSheet $StyleCollection[$AssociatedConditionalFormatting.FormattingObject.Style] -ConditionValue $Expression
                    }
                }     
            }

            if($AutoFit){
                $EndColumn = Get-OOXMLColumnString -ColNumber $($ReferencePropertySet.Length)
                $FirstColumn = Get-OOXMLColumnString -ColNumber 1
                $Sheet.Cells["$FirstColumn$($Sheet.Dimension.Start.Row):$EndColumn$LastRow"].AutoFitColumns()
            }

            <#
            if($FreezedColumnName)
            {
                $ColIdx = 1
                foreach($Property in $ReferencePropertySet)
                {
                    if($FreezedColumnName -eq $Property)
                    {
                        $Sheet.View.FreezePanes($FirstDataRowIndex,$ColIdx)
                    }
                    $ColIdx++
                }
            }
            #>

            if($FreezedColumnName)
            {
                $ColIdx = $ReferencePropertySet.IndexOf($FreezedColumnName) + 1
                if($ColIdx -gt 0){
                    $Sheet.View.FreezePanes($FirstDataRowIndex,$ColIdx)
                }
            }

            foreach($CustomFormating in $CustomFormatings)
            {
                #TODO
            }

            if($DataValidationLists)
            {
                
                $excel | Add-OOXMLWorksheet -WorkSheetName "REF_DATA"
                $DataWorkSheet = Select-OOXMLWorkSheet -WorkBook $book -WorkSheetName "REF_DATA"
                foreach($DataValidationList in $DataValidationLists)
                {
                    $DataColumnIndex = ($book.Names.Count + 1)

                    Write-Host "Named Range Count : $($book.Names.Count)" 

                    $ValueIndex = 1;

                    foreach($Value in $DataValidationList.Values)
                    {
                        $DataWorkSheet | Set-OOXMLRangeValue -Row $ValueIndex -Col $DataColumnIndex -Value $Value
                        $ValueIndex++
                    }

                    $DataRange = Convert-OOXMLCellsCoordinates -StartRow 1 -StartCol $DataColumnIndex -EndRow $ValueIndex -EndCol $DataColumnIndex
                    $book.Names.Add($DataValidationList.Name,$DataWorkSheet.Cells[$DataRange])
                }

                $FromCol = Get-OOXMLColumnString -ColNumber $($DataWorkSheet.Dimension.Start.Column)
                $ToCol = Get-OOXMLColumnString -ColNumber $($DataWorkSheet.Dimension.End.Column)
                $FromRow = $DataWorkSheet.Dimension.Start.Row
                $ToRow = $DataWorkSheet.Dimension.End.Row

                $DataWorkSheet.Cells[$("$FromCol$FromRow" + ":" + "$ToCol$ToRow")].AutoFitColumns()
 
                if($DataValidationAssignements)
                {
                    foreach($DataValidationAssignement in $DataValidationAssignements)
                    {
                        $Name = $DataValidationAssignement.Name
                        $NamedRange = $book.Names.Item($Name)
                        $ColumnNames = $($DataValidationAssignement.ColumnNames)

                        foreach($ColumnName in $ColumnNames)
                        {
                            $SelectedColumn = $($($ReferencePropertySet.IndexOf($ColumnName)) + 1)
                            $ViewRangeAddress = Convert-OOXMLCellsCoordinates -StartRow $($Sheet.Dimension.Start.Row) -StartCol $SelectedColumn -EndRow $LastRow -EndCol $SelectedColumn
                            $sheet | Add-OOXMLDataValidation -NamedRange $NamedRange -ViewRangeAddress $ViewRangeAddress
                        }
                    }
                }
            }

            $excel | Save-OOXMLPackage -FileFullPath $FileFullPath -Dispose
            return $true
        }
        catch
        {
            return $_.Exception.Message
        }

    }
}

Function Add-OOXMLDataValidation {
    <#
    .SYNOPSIS
    Apply a data validation on a given range on a given worksheet

    .DESCRIPTION
    Apply a data validation on a given range on a given worksheet

    .PARAMETER ExcelWorksheet
    The WorkSheet object where the data range is located

    .PARAMETER ViewRangeAddress
    The targeted range where data validation will be applied

    .PARAMETER NamedRange
    This is the ExcelNamedRange containing the list of valid value

    .PARAMETER ErrorStyle
    This the style of the error message

    .PARAMETER ErrorTitle
    This is the title of the error message

    .PARAMETER Error
    This is the description of the error

    .EXAMPLE
    Add-OOXMLConditionalFormatting -WorkSheet $sheet -Addresses "A1:A23" -StyleSheet $StyleGreen -RuleType GreaterThanOrEqual

    Description
    -----------
    Calls a function that will apply a data validation on a given range on a given worksheet

    .NOTES
    
    .LINK 
    
    #> 
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [OfficeOpenXml.ExcelWorksheet]$ExcelWorksheet,
        [parameter(Mandatory=$true)]
        [string]$ViewRangeAddress,
        [parameter(Mandatory=$true)]
        [OfficeOpenXml.ExcelNamedRange]$NamedRange,
        [OfficeOpenXml.DataValidation.ExcelDataValidationWarningStyle]$ErrorStyle = [OfficeOpenXml.DataValidation.ExcelDataValidationWarningStyle]::stop,
        [string]$ErrorTitle = "Error",
        [string]$Error = "Invalid Data entered !"
    )
    process
    {
        [OfficeOpenXml.DataValidation.ExcelDataValidationList]$ExcelDataValidationList = $ExcelWorksheet.DataValidations.AddListValidation($ViewRangeAddress)
        $ExcelDataValidationList.ShowErrorMessage = $true
        $ExcelDataValidationList.ErrorStyle = $ErrorStyle
        $ExcelDataValidationList.ErrorTitle = $ErrorTitle
        $ExcelDataValidationList.Error = $Error
        $ExcelDataValidationList.Formula.ExcelFormula = "=" + $NamedRange.FullAddressAbsolute
    }

}

Function Add-OOXMLConditionalFormatting {
    <#
    .SYNOPSIS
    Apply a stylesheet based on a conditional rule on a given range

    .DESCRIPTION
    Apply a stylesheet based on a conditional rule on a given range

    .PARAMETER WorkSheet
    The WorkSheet object where the cell is located

    .PARAMETER Addresses
    The targeted adresses where conditional formatting will be applied

    .PARAMETER RuleType
    The contitional formating rule type (Reduced set)

    .PARAMETER StyleSheet
    The style sheet you want to apply to the cell

    .EXAMPLE
    Add-OOXMLConditionalFormatting -WorkSheet $sheet -Addresses "A1:A23" -StyleSheet $StyleGreen -RuleType GreaterThanOrEqual

    Description
    -----------
    Calls a function that will apply a stylesheet based on a conditional rule on a given range

    .NOTES
    
    .LINK 
    
    #> 
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$true, ValueFromPipeline=$true)]
		[OfficeOpenXml.ExcelWorksheet]$WorkSheet,
        [parameter(Mandatory=$true)]
        [string[]]$Addresses,
        [parameter(Mandatory=$true)]
        [ExcelPSLib.EnumConditionalFormattingRuleType]$RuleType,
        [parameter(Mandatory=$true)]
        [OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml]$StyleSheet,
        [parameter(Mandatory=$true)]
        [string]$ConditionValue
    )
    process{
        try
        {
            $AddressString = ""
            $First = $true
            foreach($Address in $Addresses){
                if(-not $First){
                    $AddressString += ","
                    $First = $false
                }
                $AddressString += "$Address"
            }

            $ExcelAddress = New-Object OfficeOpenXml.ExcelAddress($AddressString)

            Switch($RuleType){

                $([ExcelPSLib.EnumConditionalFormattingRuleType]::BeginsWith) {
                    $ConditionalFormatted = $WorkSheet.ConditionalFormatting.AddBeginsWith($ExcelAddress)
                    $ConditionalFormatted.Text = $ConditionValue
                }

                $([ExcelPSLib.EnumConditionalFormattingRuleType]::ContainsBlanks) {
                    $ConditionalFormatted = $WorkSheet.ConditionalFormatting.AddContainsBlanks($ExcelAddress)
                }

                $([ExcelPSLib.EnumConditionalFormattingRuleType]::ContainsErrors) {
                    $ConditionalFormatted = $WorkSheet.ConditionalFormatting.AddContainsErrors($ExcelAddress)
                }

                $([ExcelPSLib.EnumConditionalFormattingRuleType]::ContainsText) {
                    $ConditionalFormatted = $WorkSheet.ConditionalFormatting.AddContainsText($ExcelAddress)
                    $ConditionalFormatted.Text = $ConditionValue
                }

                $([ExcelPSLib.EnumConditionalFormattingRuleType]::EndsWith) {
                    $ConditionalFormatted = $WorkSheet.ConditionalFormatting.AddEndsWith($ExcelAddress)
                    $ConditionalFormatted.Text = $ConditionValue
                }

                $([ExcelPSLib.EnumConditionalFormattingRuleType]::Equal) {
                    $ConditionalFormatted = $WorkSheet.ConditionalFormatting.AddEqual($ExcelAddress)
                    $ConditionalFormatted.Formula = $ConditionValue
                }

                $([ExcelPSLib.EnumConditionalFormattingRuleType]::Expression) {
                    $ConditionalFormatted = $WorkSheet.ConditionalFormatting.AddExpression($ExcelAddress)
                    $ConditionalFormatted.Formula = $ConditionValue
                }

                $([ExcelPSLib.EnumConditionalFormattingRuleType]::GreaterThan) {
                    $ConditionalFormatted = $WorkSheet.ConditionalFormatting.AddGreaterThan($ExcelAddress)
                    $ConditionalFormatted.Formula = $ConditionValue
                }

                $([ExcelPSLib.EnumConditionalFormattingRuleType]::GreaterThanOrEqual) {
                    $ConditionalFormatted = $WorkSheet.ConditionalFormatting.AddGreaterThanOrEqual($ExcelAddress)
                    $ConditionalFormatted.Formula = $ConditionValue
                }

                $([ExcelPSLib.EnumConditionalFormattingRuleType]::LessThan) {
                    $ConditionalFormatted = $WorkSheet.ConditionalFormatting.AddLessThan($ExcelAddress)
                    $ConditionalFormatted.Formula = $ConditionValue
                }

                $([ExcelPSLib.EnumConditionalFormattingRuleType]::LessThanOrEqual) {
                    $ConditionalFormatted = $WorkSheet.ConditionalFormatting.AddLessThanOrEqual($ExcelAddress)
                    $ConditionalFormatted.Formula = $ConditionValue
                }

                $([ExcelPSLib.EnumConditionalFormattingRuleType]::NotContainsBlanks) {
                    $ConditionalFormatted = $WorkSheet.ConditionalFormatting.AddNotContainsBlanks($ExcelAddress)
                }

                $([ExcelPSLib.EnumConditionalFormattingRuleType]::NotContainsErrors) {
                    $ConditionalFormatted = $WorkSheet.ConditionalFormatting.AddNotContainsErrors($ExcelAddress)
                }

                $([ExcelPSLib.EnumConditionalFormattingRuleType]::NotContainsText) {
                    $ConditionalFormatted = $WorkSheet.ConditionalFormatting.AddNotContainsText($ExcelAddress)
                    $ConditionalFormatted.Text = $ConditionValue
                }

                $([ExcelPSLib.EnumConditionalFormattingRuleType]::NotEqual) {
                    $ConditionalFormatted = $WorkSheet.ConditionalFormatting.AddNotEqual($ExcelAddress)
                    $ConditionalFormatted.Formula = $ConditionValue
                }
            }

            $ConditionalFormatted.Style.Fill.PatternType = $StyleSheet.Style.Fill.PatternType

            $ConvertedBackgroundColor = [System.Convert]::ToInt32($($StyleSheet.Style.Fill.BackgroundColor.Rgb),16)
            $ConditionalFormatted.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::FromArgb($ConvertedBackgroundColor)
        
            if($StyleSheet.Style.Border.Left.Style){$ConditionalFormatted.Style.Border.Left.Style = $StyleSheet.Style.Border.Left.Style}
      
            if($StyleSheet.Style.Border.Right.Style){$ConditionalFormatted.Style.Border.Right.Style = $StyleSheet.Style.Border.Right.Style}

            if($StyleSheet.Style.Border.Top.Style){$ConditionalFormatted.Style.Border.Top.Style = $StyleSheet.Style.Border.Top.Style}
        
            if($StyleSheet.Style.Border.Bottom.Style){$ConditionalFormatted.Style.Border.Bottom.Style = $StyleSheet.Style.Border.Bottom.Style}
        
            $ConvertedLeftBorderColor = [System.Convert]::ToInt32($($StyleSheet.Style.Border.Left.Color.Rgb),16)
            $ConditionalFormatted.Style.Border.Left.Color.Color = [System.Drawing.Color]::FromArgb($ConvertedLeftBorderColor)

            $ConvertedRightBorderColor = [System.Convert]::ToInt32($($StyleSheet.Style.Border.Right.Color.Rgb),16)
            $ConditionalFormatted.Style.Border.Right.Color.Color = [System.Drawing.Color]::FromArgb($ConvertedRightBorderColor)

            $ConvertedTopBorderColor = [System.Convert]::ToInt32($($StyleSheet.Style.Border.Top.Color.Rgb),16)
            $ConditionalFormatted.Style.Border.Top.Color.Color = [System.Drawing.Color]::FromArgb($ConvertedTopBorderColor)

            $ConvertedBottomBorderColor = [System.Convert]::ToInt32($($StyleSheet.Style.Border.Bottom.Color.Rgb),16)
            $ConditionalFormatted.Style.Border.Bottom.Color.Color = [System.Drawing.Color]::FromArgb($ConvertedBottomBorderColor)

            $ConvertedFontColor = [System.Convert]::ToInt32($($StyleSheet.Style.Font.Color.Rgb),16)
            $ConditionalFormatted.Style.Font.Color.Color = [System.Drawing.Color]::FromArgb($ConvertedFontColor)

            $ConditionalFormatted.Style.Font.Italic = $StyleSheet.Style.Font.Italic
            $ConditionalFormatted.Style.Font.Bold = $StyleSheet.Style.Font.Bold

            $ConditionalFormatted.Style.NumberFormat.Format = $StyleSheet.Style.Numberformat.Format
        }
        catch
        {
            return $_.Exception.Message
        }
    }
}

Function Get-OOXMLFormulaObject {
    <#
    .SYNOPSIS
    This function is just an helper that will return a pscustomobject compliant with the Export-OOXML cmdlet (-FormulaObject)
    
    .DESCRIPTION
    This function is just an helper that will return a pscustomobject compliant with the Export-OOXML cmdlet (-FormulaObject)

    .PARAMETER Name
    This is the property targeted by the conditional formatting

    .PARAMETER Style
    Is one of the 141 style available that will be applied

    .PARAMETER Operation
    The operation you want to perform on the column "SUM","AVERAGE","COUNT","MAX","MIN","SUMIF","AVERAGEIF","COUNTIF"

    .PARAMETER Criteria
    The criteria for the following conditional operations "SUMIF","AVERAGEIF","COUNTIF"

    .EXAMPLE
    Get-OOXMLFormulaObject -Name Size -Style DarkGray -Operation "COUNTIF" -Criteria ">5"

    Description
    -----------
    Calls a function which will return a pscustomobject compliant with the Export-OOXML cmdlet (-FormulaObject)
    
    .NOTES
    
    .LINK 
    
    #>
    [CmdletBinding()]
    param(
        [alias("N")]
        [parameter(Mandatory=$true)]
        [string]$Name,
        [alias("S")]
        [parameter(Mandatory=$true)]
        [ExcelPSLib.EnumColors]$Style,
        [alias("O")]
        [parameter(Mandatory=$true)]
        [ExcelPSLib.EnumOperations]$Operation,
        [alias("C")]
        [string]$Criteria = [String]::Empty
    )
    process{
        return [PSCustomObject]@{Name=$Name;Style=$($Style.ToString());Operation=$($Operation.ToString());Criteria=$Criteria}
    }
}

Function Get-OOXMLConditonalFormattingCustomObject{
    <#
    .SYNOPSIS
    This function is just an helper that will return a pscustomobject compliant with the Export-OOXML cmdlet (-ConditionalFormatings)
    
    .DESCRIPTION
    This function is just an helper that will return a pscustomobject compliant with the Export-OOXML cmdlet (-ConditionalFormatings)

    .PARAMETER Name
    This is the property targeted by the conditional formatting

    .PARAMETER Style
    Is one of the four style available that will be applied if the condition is true

    .PARAMETER Condition
    Condition is one of the condition present in the following enum EnumConditionalFormattingRuleType

    .PARAMETER Value
    Is the value that will be used on the propoertie according to the choosen condition

    .EXAMPLE
    $ConditionalObject = Get-OOXMLConditonalFormattingCustomObject -Name "__PROPERTY_COUNT" -Style Red -Condition GreaterThan -Value 30

    Description
    -----------
    Calls a function which will return a pscustomobject compliant with the Export-OOXML cmdlet (-ConditionalFormatings)
    
    .NOTES
    
    .LINK 
    
    #>
    [CmdletBinding()]
    param(
        [alias("N")]
        [parameter(Mandatory=$true)]
        [string]$Name,
        [alias("S")]
        [parameter(Mandatory=$true)]
        [ExcelPSLib.EnumColors]$Style,
        [alias("C")]
        [parameter(Mandatory=$true)]
        [ExcelPSLib.EnumConditionalFormattingRuleType]$Condition,
        [alias("V")]
        [parameter(Mandatory=$true)]
        $Value
    )
    process{
        return [PSCustomObject]@{Name=$Name;Style=$($Style.ToString());Condition=$Condition;Value=$Value}
    }
}

Function Test-DataTypeIntegrity{
    <#
    .SYNOPSIS
    This function was made to test that all data under the header of a column are of the same data type
    
    .DESCRIPTION
    This function was made to test that all data under the header of a column are of the same data type

    .PARAMETER Worksheet
    This is the worksheet targeted by this test function

    .PARAMETER Column
    This is the column targeted by this test function

    .EXAMPLE
    Test-DataTypeIntegrity -Worksheet $Worksheet -Column 4

    Description
    -----------
    Calls a function which will test that all data under the header of a column are of the same data type
    
    .NOTES
    
    .LINK 
    
    #>
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true)]
        [OfficeOpenXml.ExcelWorksheet]$Worksheet,
        [parameter(Mandatory=$true)]
        [int]$Column
    )
    process{
        $Top = $Worksheet.Dimension.Start.Row
        $Left = $Worksheet.Dimension.Start.Column
        $Bottom = $Worksheet.Dimension.End.Row
        $Right = $Worksheet.Dimension.End.Column

        for($i=$Top+1; $i -lt ($Bottom+1); $i++){
           [string]$CurrentDataType = $($Worksheet.GetValue($i,$Column).GetType().FullName)
           if($i -eq ($Top+1)){[string]$DataType = $CurrentDataType}
           if($DataType -inotmatch $CurrentDataType){
                return "string"
           }
        }
        return $DataType
    }
}

Function Import-OOXML{
    <#
    .SYNOPSIS
    Import an XLSX File an convert it to an array of objects
    
    .DESCRIPTION
    Import an XLSX File an convert it to an array of objects

    .PARAMETER FileFullPath
    The full path of the XLSX File
    
	.PARAMETER WorksheetName
    The name of the worksheet in the XLSX Document
    
    .PARAMETER WorksheetID
    The id of the worksheet in the XLSX Document

    .PARAMETER KeepDataType
    This is a switch parameter that when set will indicate to the import function that it should try to detect and keep data type per column

	.PARAMETER Range
    This is an optional parameter which allow one to specify a range to use for the input.  The range is specified in the normal Excel format e.g. C10:J22

    .EXAMPLE
    Import-OOXML -FileFullPath C:\Temp\DevBook.xlsx -WorksheetNumber 1 

    .EXAMPLE
    Import-OOXML -FileFullPath C:\Temp\DevBook.xlsx -WorksheetName Sheet1
    
    .EXAMPLE
    Import-OOXML -FileFullPath C:\Temp\DevBook.xlsx -WorksheetNumber 1  -Range "C4:J10"

    Description
    -----------
    Calls a function that will import an XLSX File an convert it to an array of objects
    
    .NOTES
    
    .LINK 
    
    #>
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true)]
        [ValidateScript({Test-Path -Path $_ -PathType Leaf})]
        [String]$FileFullPath,
        [parameter(ParameterSetName="WorksheetIndex", Mandatory=$true)]
		[int]$WorkSheetNumber,
        [parameter(ParameterSetName="WorksheetName", Mandatory=$true)]
		[string]$WorkSheetName,
        [string]$Range,
        [switch]$KeepDataType=$false
    )
    begin{
        Function Get-ColumnHeaders{
            $ColumHeaders = @()
        }
    }
    process{
        try
        {
            [System.IO.FileInfo]$XLSXFile = New-Object System.IO.FileInfo($FileFullPath)
		    $ExcelInstance = New-Object OfficeOpenXml.ExcelPackage($XLSXFile)
            [OfficeOpenXml.ExcelWorkbook]$Workbook = $ExcelInstance | Get-OOXMLWorkbook
            if($Workbook -ne $null){
                if($Workbook.Worksheets.Count -ge $WorksheetId){
                    if($WorkSheetName){
                        [OfficeOpenXml.ExcelWorksheet]$Worksheet = $Workbook | Select-OOXMLWorkSheet -WorksheetName $WorksheetName
                    }else{
                        [OfficeOpenXml.ExcelWorksheet]$Worksheet = $Workbook | Select-OOXMLWorkSheet -WorkSheetNumber $WorkSheetNumber
                    }

                    $Top = $Worksheet.Dimension.Start.Row
                    $Left = $Worksheet.Dimension.Start.Column
                    $Bottom = $Worksheet.Dimension.End.Row
                    $Right = $Worksheet.Dimension.End.Column

					If($Range){
                       $addrHash = Convert-OOXMLFromExcelCoordinates -Address $Range
                       $Top = $addrHash["TopRow"]
                       $Left = $addrHash["LeftCol"]
                       $Bottom = $addrHash["BottomRow"]
                       $Right = $addrHash["RightCol"]
                    }

                    $ClassGuidName = "Custom_" + [System.Guid]::NewGuid().ToString().Replace("-","")
                
                    $ClassDeclaration = "public class $ClassGuidName"
                    $ClassDeclaration += [System.Environment]::NewLine
                    $ClassDeclaration += "{"
                    $TestString = ""

                    $PropertyList = @()

                    ######### If column header has no data, Stop !!!
                    $i=$left
                    While($i -lt ($Right+1))
                    {
                        [string]$Data = $($Worksheet.GetValue($top,$i))
                        if(!$Data)
                        {
                            $i = $Right
                        }
                        else
                        {
                            $Data = [Regex]::Replace($Data, "[^0-9a-zA-Z_]", "")
                            $PropertyList += $Data
                            $ClassDeclaration += $([System.Environment]::NewLine)

                            if($KeepDataType){
                                $ClassDeclaration += "public $(Test-DataTypeIntegrity -Worksheet $Worksheet -Column $i) @$($Data);"
                            }else{
                                $ClassDeclaration += "public string @$($Data);"
                            }

                        }
                        $i++;
                    }
                
                    $ClassDeclaration += $([System.Environment]::NewLine)
                    $ClassDeclaration += "}"

                    $FinalClassDefinition = @"
                        $ClassDeclaration
"@

                    try {Add-Type -Language CSharp -TypeDefinition $FinalClassDefinition;} catch { return $_.Exception.Message }
                
                    $FullArray = @()
                    for($i=$Top+1; $i -lt ($Bottom+1); $i++){
                        $TempObject = New-Object $ClassGuidName
                        $idx=0
                        foreach($Prop in $PropertyList){
                            $TempObject.$Prop = $($Worksheet.GetValue($i,$($Left+$idx)))
                            $idx++
                        }
                        $FullArray += $TempObject
                    }
                    $FullArray

                }else{
                    Write-Error "This worksheet doesn't exist !"
                }
            }else{
                Write-Error "There is no workbook in this document !"
            }
        }
        catch
        {
            return $_.Exception.Message
        }
    }
}

Function Save-OOXMLPackage {
    <#
    .SYNOPSIS
    Save the Excel Instance to a definited XLSX File
    
    .DESCRIPTION
    Save the Excel Instance to a definited XLSX File

    .PARAMETER FileFullPath
    The full path of the XLSX File

    .PARAMETER ExcelInstance
    The Current ExcelPackage instance

    .PARAMETER Dispose
    Free the memory by closing the Excel Instance

    .EXAMPLE
    $excel | Save-OOXMLPackage -FileFullPath $OutputFileName -Dispose

    Description
    -----------
    Calls a function which will save the Current ExcelPackage instance to an XLSX file
    
    .NOTES
    
    .LINK 
    
    #>
    [CmdletBinding()]
	param (
		[string]$FileFullPath,
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]
		[OfficeOpenXml.ExcelPackage]$ExcelInstance,
        [switch]$Dispose
	)
	process{
        try
        {
            if($FileFullPath){
		        $bin = $ExcelInstance.GetAsByteArray();
		        [io.file]::WriteAllBytes($FileFullPath,$bin)
            }else{
                $ExcelInstance.Save()
            }
		
            if($Dispose){
                $ExcelInstance.Dispose()
            }
        }
        catch
        {
            return $_.Exception.Message
        }
	}
}

Function Get-OOXMLDataValidationCustomObject {
    <#
    .SYNOPSIS
    Return a Data Validation Object of the correct format
    
    .DESCRIPTION
    Return a Data Validation Object of the correct format

    .PARAMETER Name
    The name that will be given to the Data Range

    .PARAMETER Values
    The values that will be inserted into the Data Range


    .EXAMPLE
    Get-OOXMLDataValidationCustomObject -Name "Data Name" -Values @("Value_01","Value_02","Value_03")

    Description
    -----------
    Calls a function that will return a Data Validation Object of the correct format :

        [pscustomobject]@{
            Name = "Name"
            Values = @("Value_01","Value_02","Value_03")
        }
    
    .NOTES
    
    .LINK 
    
    #>
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true)]
        [string]$Name,
        [parameter(Mandatory=$true)]
        [object[]]$Values
    )
    process
    {
        return [pscustomobject]@{
            Name = $Name
            Values = $Values
        }
    }
}

Function Get-OOXMLDataValidationAssignementCustomObject {
    <#
    .SYNOPSIS
    Return a Data Validation Assignement Object of the correct format
    
    .DESCRIPTION
    Return a Data Validation Assignement Object of the correct format

    .PARAMETER Name
    The name of the named list containing the alowed value

    .PARAMETER ColumnNames
    The name of the columns that should receive data validation

    .EXAMPLE
    Get-OOXMLDataValidationAssignementCustomObject -DataValidationName "FirstList" -ColumnNames @("Name","Handles")

    Description
    -----------
    Calls a function that will return a Data Validation Object of the correct format :

        [pscustomobject]@{
            Name = "Name"
            ColumnNames = @("Col_01","Col_02","Col_03")
        }
    
    .NOTES
    
    .LINK 
    
    #>
    [CmdletBinding()]
    param(
        [string]$DataValidationName,
        [string[]]$ColumnNames
    )
    process{
        return [pscustomobject]@{
            Name = $DataValidationName
            ColumnNames = $ColumnNames
        }
    }
}

Function Open-OOXML{
<#
    .SYNOPSIS
    Opens an existing Excel file
    
    .DESCRIPTION
    Opens an existing Excel file

    .PARAMETER FileFullPath
    The location of the Excel file

    .EXAMPLE
    open-OOXML -FileFullPath C:\Temp\DevBook.xlsx

    Description
    -----------
    Returns the "OfficeOpenXml.ExcelPackage" object for the file
    
    .NOTES
    
    .LINK 
    
    #>
   [CmdletBinding()]
   param(
      [parameter(Mandatory=$true)]
      [ValidateScript({Test-Path -Path $_ -PathType Leaf})]
      [String]$FileFullPath
   )
   begin{

   }
   process{
      $retval=$null
      try{
         [System.IO.FileInfo]$XLSXFile = New-Object System.IO.FileInfo($FileFullPath)
	     $ExcelInstance = New-Object OfficeOpenXml.ExcelPackage($XLSXFile)
      }catch{
         $retval = $_.Exception.Message
      }

      if($retval -ne $null){
         return $retval
      }else{
         return $ExcelInstance
      }
   }
}

Function Convert-OOXMLFromExcelCoordinates{
<#
    .SYNOPSIS
    Utility function to convert an Excel location reference to row and column numbers
    
    .DESCRIPTION
    Converts an Excel range reference to row and column numbers.  It returns a hash containing the values

    .PARAMETER Address
    The Excel format of the cell address e.g. C22 or B22:G99

    .EXAMPLE
    Convert-OOXMLFromExcelCoordinates -address "B3".  This returns a hash containing:
    Name                           Value                                                                                                                                              
    ----                           -----                                                                                                                                              
    RightCol                       2                                                                                                                                                  
    BottomRow                      3                                                                                                                                                  
    TopRow                         3                                                                                                                                                  
    LeftCol                        2  

    .EXAMPLE
    Convert-OOXMLFromExcelCoordinates -address "B3:H10".  This returns a hash containing:
    Name                           Value                                                                                                                                              
    ----                           -----                                                                                                                                              
    RightCol                       8                                                                                                                                                  
    BottomRow                      10                                                                                                                                                 
    TopRow                         3                                                                                                                                                  
    LeftCol                        2

    Description
    -----------
    Returns a hash containing the numeric location of the cell.
    
    .NOTES
    
    .LINK 
    
    #>
   param(
      [parameter(Mandatory=$true)][string]$Address
   )


   $coordinates=@{}
   $ExcelAddress = New-Object OfficeOpenXml.ExcelAddress($Address)

   $coordinates["TopRow"]=$ExcelAddress.Start.Row
   $coordinates["LeftCol"]=$ExcelAddress.Start.Column
   $coordinates["BottomRow"]=$ExcelAddress.end.Row
   $coordinates["RightCol"]=$ExcelAddress.end.Column

   return $coordinates
}

function Read-OOXMLcell{
<#
    .SYNOPSIS
    Reads a single Excel cel
    
    .DESCRIPTION
    Retrieves the value of a given cell

    .PARAMETER Row
    The numeric row

    .PARAMETER Col
    The numeric culumn

    .PARAMETER Address
    The Excel format of the cell address e.g. C22

    .PARAMETER WorkSheet
    The worksheet holding the information

    .EXAMPLE
    read-oosmlcell -Worksheet $sheet -Row 3 -Col 2

    .EXAMPLE
    read-oosmlcell -Worksheet $sheet -Address "B3"

    Description
    -----------
    Returns the value in a specified cell
    
    .NOTES
    
    .LINK 
    
    #>
   param(
      [parameter(ParameterSetName="RowCol",Mandatory=$true)][string]$Row,
      [parameter(ParameterSetName="RowCol",Mandatory=$true)][string]$Col,
      [parameter(ParameterSetName="Address",Mandatory=$true)][string]$Address,
      [parameter(Mandatory=$true, ValueFromPipeline=$true)][OfficeOpenXml.ExcelWorksheet]$WorkSheet
   )
   if($Address){
      $range = Convert-OOXMLFromExcelCoordinates($Address)
      $Row = $range["TopRow"]
      $col = $range["LeftCol"]
   }
   $retVal = $WorkSheet.GetValue($Row,$Col)

   return $retVal
}
