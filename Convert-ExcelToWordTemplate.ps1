<#

.SYNOPSIS
This is a simple Powershell script to get data from an Excel file and populate a word template

.DESCRIPTION
The script will use Get-ExcelWorkSheet to convert an excel worksheet to a powershell object

.EXAMPLE



.NOTES
  Author:   Leigh Butterworth
  Version:  1.0

.LINK
https://github.com/L37hal/Get-ExcelWorkSheet

#>

Param(
    [parameter(Mandatory=$false)][string]$File,
    [parameter(Mandatory=$false)][string]$WorkSheet,
    [parameter(Mandatory=$false)][string]$Header,
    [parameter(Mandatory=$false)][string]$Template,
    [parameter(Mandatory=$false)][string]$OutMappings = "yes",
    [parameter(Mandatory=$false)][string]$MappingsFile = ".\Mappings.csv",
    [parameter(Mandatory=$false)][string]$OutFileFolder,
    [parameter(Mandatory=$false)][string]$OutFilePrefix,
    [parameter(Mandatory=$false)][string]$OutFileHeadersuffix,
    [parameter(Mandatory=$false)][string]$initialDirectory = "C:\"
) # End Param()

# Get Scripts

if (!(!".\Get-ExcelWorkSheet.ps1"))
{
 Invoke-WebRequest -Uri "https://raw.githubusercontent.com/L37hal/Get-ExcelWorkSheet/master/Get-ExcelWorkSheet.ps1" -OutFile ".\Get-ExcelWorkSheet.ps1"
}

if (!(!".\Replace-WordTemplate.ps1"))
{
 Invoke-WebRequest -Uri "https://raw.githubusercontent.com/L37hal/Replace-WordTemplate/master/Replace-WordTemplate.ps1" -OutFile ".\Replace-WordTemplate.ps1"
}

# *** Entry Point to Functions ***

Function Get-File($initialDirectory)
{  
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
} # end function Get-File

Function Get-Folder($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    # $foldername.rootfolder = $initialDirectory
    $foldername.ShowDialog() | Out-Null
    $foldername.SelectedPath
} # end function Get-Folder

Function Get-OutputFilename($item)
{
    $OutFileSuffix = $excelData.$OutFileHeadersuffix[$item]
    $OutputFilename = "$OutFileFolder\$OutFilePrefix$OutFileSuffix.docx"
    $OutputFilename
} # end function Get-OutputFilename

# *** Entry Point to Script ***

# Initialize the excelData array
$excelData = @()
# Get the data for the excelData array
$excelData = . .\Get-ExcelWorkSheet -File $File -WorkSheet $WorkSheet -Header $Header
# Get the Headers
$Headers = $excelData | Get-member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name'


if (!(test-path $MappingsFile))
{
    $Mappings = new-object PSObject
    For ($i = 0; $i -le $Headers.count-1; $i++)
    {
        $uniqueID = -join ((65..90) + (97..122) | Get-Random -Count 7 | % {[char]$_})
        $uniqueID = "a$uniqueID"
        $Mappings | add-member -membertype NoteProperty -name $Headers[$i] -Value  $uniqueID
    }
    if ($OutMappings -eq "yes")
    {
        $Mappings | export-csv -Path $MappingsFile -NoTypeInformation
    }
}
Else
{
    $Mappings = import-csv $MappingsFile
}

$text = "`n"
$text += "Please use these Values in the word template now`n"
$text += "`n"
ForEach ($Header in $Headers){
    $item = "$Header= "
    $Header =  $Mappings.$Header | Out-String
    $item += $Header
    $text += $item
}

Clear-Host
Write-Host $text
pause

if (!$Template)
{
    Clear-Host
    $text = "`n"
    $text += "Please select the word template"
    Write-Host $text
    $Template = Get-File -initialDirectory "C:\"
}

if (!$OutFilePrefix)
{
    $text = "`n"
    $text += "What will be the filename prefix?`n"
    $text += "`n"
    Clear-Host
    $OutFilePrefix = Read-Host $text
}

if (!$OutFileHeadersuffix)
{
    $text = "`n"
    $text += "What will be the filename suffix?`n"
    $text += "`n"
    $count = 0
    ForEach ($Header in $Headers)
    {
        $count += 1
        $text += "$count) $Header`n"
    }
    Clear-Host
    $OutFileHeadersuffix = Read-Host $text
    $OutFileHeadersuffix = $Headers[($OutFileHeadersuffix-1)]
}

if (!$OutFileFolder)
{
    $text = "`n"
    $text += "Where will the documents be output (folder)?`n"
    $text += "`n"
    Clear-Host
    Write-Host $text
    $OutFileFolder = Get-Folder($initialDirectory)
}

$text = "`n"

$text += "`n"
$text += "All files will be output to: `n $OutFileFolder\`n"
$text += "`n"
$ExampleItem = 1
$OutputFilename = Get-OutputFilename($ExampleItem)
$text += "An example file will be: `n $OutputFilename"
$text += "`n"
$text += "`n"
$text += "Press Ctrl+C to quit..."
Clear-Host
Write-Host $text
pause

For ($i = 0; $i -le $excelData.count-1; $i++)
{
    $OutputFilename = Get-OutputFilename($i)
    [array]$Dataset = $excelData[$i]
    . .\Replace-WordTemplate -Template $Template -OutPath $OutputFilename -Mappings $Mappings -Dataset $Dataset
    Write-Host "Generated: $OutputFilename`n"
}

Remove-Item ".\Get-ExcelWorkSheet.ps1"
Remove-Item ".\Replace-WordTemplate.ps1"