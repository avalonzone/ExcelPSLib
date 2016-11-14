#Require ExcelPSLib 0.5.7
Import-Module ExcelPSLib -Force

$ComputerList = @("LOCALHOST")
$RowPosition = 2
$OutputFileName = "c:\temp\EXCELPSLIB_Demo.xlsx"

[OfficeOpenXml.ExcelPackage]$excel = New-OOXMLPackage -author "Avalon77" -title "ComputerInfos"
[OfficeOpenXml.ExcelWorkbook]$book = $excel | Get-OOXMLWorkbook

$excel | Add-OOXMLWorksheet -WorkSheetName "Local HDD" -AutofilterRange "A2:E2"
$sheet = $book | Select-OOXMLWorkSheet -WorkSheetNumber 1

$StyleGreen = New-OOXMLStyleSheet -WorkBook $book -Name "GirlStyle" -Bold -ForeGroundColor Black -FillType Solid -BackGroundColor Green -borderStyle Thin -BorderColor Black -NFormat "#,##0.00"
$StyleRed = New-OOXMLStyleSheet -WorkBook $book -Name "BoyStyle" -Bold -ForeGroundColor Black -FillType Solid -BackGroundColor Red -borderStyle Thin -BorderColor Black -NFormat "#,##0.00"
$StyleHeader = New-OOXMLStyleSheet -WorkBook $book -Name "HeaderStyle" -Bold -ForeGroundColor White -BackGroundColor Black -Size 14 -HAlign Center -VAlign Center -FillType Solid
$StyleNormal = New-OOXMLStyleSheet -WorkBook $book -Name "NormalStyle" -borderStyle Thin -BorderColor Black
$StyleNumber = New-OOXMLStyleSheet -WorkBook $book -Name "Float" -NFormat "#,##0.00"
$StyleConditionalFormatting = New-OOXMLStyleSheet -WorkBook $book -Name "ConditionalF" -Bold -ForeGroundColor Black -FillType Solid -BackGroundColor Orange -borderStyle Double -BorderColor Blue -NFormat "#,##0.0000" -Italic

$sheet | Set-OOXMLRangeValue -row $RowPosition -col 1 -value "Computer" -StyleSheet $StyleHeader | Out-Null
$sheet.Column(1).Width = 22
$sheet | Set-OOXMLRangeValue -row $RowPosition -col 2 -value "Drive" -StyleSheet $StyleHeader | Out-Null
$sheet.Column(2).Width = 16
$sheet | Set-OOXMLRangeValue -row $RowPosition -col 3 -value "Space" -StyleSheet $StyleHeader | Out-Null
$sheet.Column(3).Width = 22
$sheet | Set-OOXMLRangeValue -row $RowPosition -col 4 -value "Freespace" -StyleSheet $StyleHeader | Out-Null
$sheet.Column(4).Width = 22
$sheet | Set-OOXMLRangeValue -row $RowPosition -col 5 -value "SpaceRatio" -StyleSheet $StyleHeader | Out-Null
$sheet.Column(5).Width = 22

$RowPosition++

foreach($Computer in $ComputerList){

    if(Test-Connection -ComputerName $Computer -Count 1 -BufferSize 16){
        $LocaHardDrive = Get-WmiObject -query "Select * FROM win32_logicaldisk where DriveType=3" | Select-Object -Property *
        $VolumeSerialNumbers = @()

        foreach($Disk in $LocaHardDrive){
    
            if($Disk.size -gt 0){

                $VolumeSerialNumbers += $Disk.VolumeSerialNumber

                $FreeSpace = $Disk.freespace
                $TotalSpace = $Disk.size
                $Caption = $Disk.caption
                $FreeSpaceRatio = $FreeSpace / $TotalSpace * 100

                $sheet | Set-OOXMLRangeValue -Row $RowPosition -Col 1 -Value $Computer -StyleSheet $StyleGreen | Out-Null
                $sheet | Set-OOXMLRangeValue -row $RowPosition -col 2 -value $Caption -StyleSheet $StyleGreen | Out-Null
                $sheet | Set-OOXMLRangeValue -row $RowPosition -col 3 -value $($TotalSpace / 1GB) -StyleSheet $StyleGreen | Out-Null
                $sheet | Set-OOXMLRangeValue -row $RowPosition -col 4 -value $($FreeSpace / 1GB) -StyleSheet $StyleGreen | Out-Null

                if($FreeSpaceRatio -lt 10){
                    $sheet | Set-OOXMLRangeValue -row $RowPosition -col 5 -value $FreeSpaceRatio -StyleSheet $StyleRed | Out-Null
                }else{
                    $sheet | Set-OOXMLRangeValue -row $RowPosition -col 5 -value $FreeSpaceRatio -StyleSheet $StyleGreen | Out-Null
                }

                $RowPosition++
            }

        }

        Export-OOXML -InputObject $LocaHardDrive -FileFullPath "C:\Temp\$computer.xlsx" -ConditionalFormating @([PSCustomObject]@{Name="DeviceID";Style="Red";Condition="BeginsWith";Value="L"},[PSCustomObject]@{Name="DeviceID";Style="Green";Condition="BeginsWith";Value="B"})

    }else{

        $sheet | Set-OOXMLRangeValue -Row $RowPosition -Col 1 -Value $Computer -StyleSheet $StyleRed | Out-Null
        $sheet | Set-OOXMLRangeValue -row $RowPosition -col 2 -value "N/A" -StyleSheet $StyleRed | Out-Null
        $sheet | Set-OOXMLRangeValue -row $RowPosition -col 3 -value 0 -StyleSheet $StyleRed | Out-Null
        $sheet | Set-OOXMLRangeValue -row $RowPosition -col 4 -value 0 -StyleSheet $StyleRed | Out-Null

        $RowPosition++
    }

    $sheet | Add-OOXMLConditionalFormatting -Addresses "E3:E$($RowPosition-1)" -StyleSheet $StyleConditionalFormatting -RuleType GreaterThanOrEqual -ConditionValue "50"
    $sheet | Add-OOXMLConditionalFormatting -Addresses "B3:B$($RowPosition-1)" -StyleSheet $StyleConditionalFormatting -RuleType BeginsWith -ConditionValue "L"
    
}

$excel | Save-OOXMLPackage -FileFullPath $OutputFileName -Dispose

$ImportedOOXMLData = Import-OOXML -FileFullPath $OutputFileName -WorksheetNumber 1

$ImportedOOXMLData

