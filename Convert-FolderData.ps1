# Script Version
# $sScriptVersion = "0.30"

# Set Error Action to Silently Continue
$ErrorActionPreference =  "SilentlyContinue"

# add WPF for message boxes
Add-Type -AssemblyName PresentationFramework

# define script folder
$scriptFolder = $PSScriptRoot

# define functions 
$fFunction1 = join-Path $scriptFolder "openFolder.ps1"

#load powershell functions
."$fFunction1"

# The html source template file
$sHtmlTemplate = 'HTMLTemplate.html'
$sourceHTMLFile = Join-Path -Path $scriptFolder -ChildPath $sHtmlTemplate

# select folder
#The starting folder to analyze
$startFolder = Get-SHDOpenFolderDialog -Title "Select the root folder for the Table of Contents"

# define the output file
# The final html file that will be produced, #does not need to exist
$OutputFileName ="Inhaltsverzeichnis.html"
$OutputfolderTableofContent ="00 Inhaltsverzeichnis"
$destinationHTMLFolder = Join-Path -Path $startFolder -ChildPath $OutputfolderTableofContent
$destinationHTMLFile = Join-Path -Path $destinationHTMLFolder  -ChildPath $OutputFileName

# define image folder
$imageSourceFolder = Join-Path -Path $scriptFolder -ChildPath '.images'
$imageDestinationFolder = Join-Path -Path $destinationHTMLFolder -ChildPath '.images'

# message box if calculate the hash 
$msgBoxInputSize =  [System.Windows.MessageBox]::Show('Soll die Größe der Dateien ermittelt werden?', 'Mit Dateigröße', 'YesNo','Info')

[bool]$SizeOff = $false
switch  ($msgBoxInputSize) {
    'Yes' {
    	$SizeOff = $false
    }
    'No' {
    	$SizeOff = $true
    }
}


# message box if calculate the hash 
$msgBoxInputHash =  [System.Windows.MessageBox]::Show('Soll der Hash Wert ermittelt werden?', 'Mit Hash Wert?', 'YesNo','Info')

[bool]$HashOff = $false
switch  ($msgBoxInputHash) {
    'Yes' {
    	$HashOff = $false
    }
    'No' {
    	$HashOff = $true
    }
}




$htmlLines = @()

function CreateFileDetailRecord{
    param(
        [string]$FilePath
    )
    
    process{
        # read file data
        $files = Get-ChildItem -Path $FilePath -File | Select-Object Name,LastWriteTime,Fullname,Length  
        # get hash and print the result to logfile
        $newFIleRecord = New-Object -TypeName PSObject 
        if (!$HashOff){ 
            $hash = Get-FileHash -Path $FilePath | Select-Object Hash
            $shash = $hash.Hash
        }
        else {
            $shash = " "
        }
        $Filename = $files.Name
        $WriteTime = $files.LastWriteTime
        $FullFileName = $files.Fullname
        $Size = $files.Length

        #Determine units for a more friendly output
    if(($Size / 1GB) -ge 1){
        [string]$units = "GB"
        $fileSize = [math]::Round(($Size / 1GB),2)
    }
    else
    {
        if(($Size / 1MB) -ge 1){
            $units = "MB"
            $fileSize = [math]::Round(($Size / 1MB),2)
        }
        else{
            $units = "KB"
            $fileSize = [math]::Round(($Size / 1KB),2)
        }
    }
        $newFIleRecord | Add-Member -MemberType NoteProperty -Name FileName -Value $Filename
        $newFIleRecord | Add-Member -MemberType NoteProperty -Name LastWriteTime -Value $WriteTime
        $newFIleRecord | Add-Member -MemberType NoteProperty -Name Hash -Value $shash
        $newFIleRecord | Add-Member -MemberType NoteProperty -Name Fullname -Value $FullFileName
        $newFIleRecord | Add-Member -MemberType NoteProperty -Name Size -Value $fileSize
        $newFIleRecord | Add-Member -MemberType NoteProperty -Name Units -Value $units   
    }
    end{
        return $newFIleRecord;
    }
}

#Function that creates a folder detail record
function CreateFolderDetailRecord{
    param(
        [string]$FolderPath
    )
    
    process{
    #Get the total size of the folder by recursively summing its children
    $subFolderItems = Get-ChildItem $FolderPath -recurse -force | Where-Object {$_.PSIsContainer -eq $false} | Measure-Object -property Length -sum | Select-Object Sum
    $folderSizeRaw = 0
    $folderSize = 0
    $units = ""

    #Account for no children
    if($subFolderItems.sum -gt 0){
        $folderSizeRaw = $subFolderItems.sum     
    }    

    #Determine units for a more friendly output
    if(($subFolderItems.sum / 1GB) -ge 1){
        $units = "GB"
        $folderSize = [math]::Round(($subFolderItems.sum / 1GB),2)
    }
    else
    {
        if(($subFolderItems.sum / 1MB) -ge 1){
            $units = "MB"
            $folderSize = [math]::Round(($subFolderItems.sum / 1MB),2)
        }
        else{
            $units = "KB"
            $folderSize = [math]::Round(($subFolderItems.sum / 1KB),2)
        }
    }

    #Create an object with the given properties
    $newFolderRecord = New-Object -TypeName PSObject
    $newFolderRecord | Add-Member -MemberType NoteProperty -Name FolderPath -Value $FolderPath
    $newFolderRecord | Add-Member -MemberType NoteProperty -Name FolderSizeRaw -Value $folderSizeRaw
    $newFolderRecord | Add-Member -MemberType NoteProperty -Name FolderSizeInUnits -Value $folderSize
    $newFolderRecord | Add-Member -MemberType NoteProperty -Name Units -Value $units
    }
    end{
        return $newFolderRecord;
    }
}

function Convert-FileDataToHTML {
    [CmdletBinding()]
    param (
        [string]$FilePath
    )
    
    begin {
        # Read file data
        $dataFile = CreateFileDetailRecord -FilePath $FilePath
    }
    
    process {
        [DateTime]$date = $dataFile.LastWriteTime 
        [string]$dateFormat = $date.tostring("dd-MMM-yyyy hh.mm.ss")
        $FileLink = Resolve-Path -Path $dataFile.Fullname -Relative
        # convert to HTML link
        $FileLink = $FileLink.Replace("\","/").Replace(" ", "%20")
        $HTMLOutput = '<li><a href="../' + $FileLink +'">' + $($dataFile.FileName) + '</a> <span class= FileStyle>  </span>' + "`n" +
        '<ul>' + "`n" +
                '<li> &ensp; [<span style="color:grey"> Änderungsdatum: </span>] &ensp; (<span style="color:blue">' + $dateFormat + '</span>)</li>' + "`n" 
               if (!$SizeOff){
                  $HTMLOutput = $HTMLOutput + '<li> &ensp; [<span style="color:grey"> Dateigröße: </span>] &emsp; &emsp; &emsp;(<span style="color:blue">' + $($dataFile.Size) +' ' + $($dataFile.Units) + '</span>)</li>' + "`n"  
                }
               if (!$HashOff){ 
                    $HTMLOutput = $HTMLOutput + ' <li> &ensp; [<span style="color:grey"> File Hash: </span>] &emsp; &emsp; &ensp; (<span style="color:blue">' + $($dataFile.Hash) + '</span>)</li>' + "`n"
               }
                $HTMLOutput = $HTMLOutput +'</ul>' + "`n" + '</li>' + "`n"    
    }
    
    end {
        Return $HTMLOutput
    }
}

#Function that recursively creates the html for the output, given a starting location
function GetAllFolderDetails{
    param([string]$FolderPath)    

    $recursiveHTML = @()
    [int]$cont = 0
    #Get properties used for processing
    $folderItem = Get-Item -Path $FolderPath
    $folderDetails = CreateFolderDetailRecord -FolderPath $FolderPath
    $subFolders = Get-ChildItem $FolderPath | Where-Object {$_.PSIsContainer -eq $true} | Sort-Object

    #If has subfolders, create hmtl drilldown. 
    if($subFolders.Count -gt 0)
    {
        $recursiveHTML += '<li><span class="caret">' + $folderItem.Name
          if (!$SizeOff){
            $recursiveHTML += '(<span style="color:red">' + $folderDetails.FolderSizeInUnits + " " + $folderDetails.Units + '</span>)</span>' + "`n"
          }
          else{
            $recursiveHTML += '</span>' + "`n"
          }
        $recursiveHTML += '<ul class="nested">'
        #Get all file data in subfolder
        $files = Get-ChildItem -Path $folderItem.FullName -File| Select-Object Name,FullName
        $cont = ($files | Measure-Object  | Select-Object Count).Count  
        if ($cont -gt 0){
            If ($cont -eq 1){
                $recursiveHTML += Convert-FileDataToHTML -FilePath $files.Fullname
            }
            else {
                foreach($file in $files.GetEnumerator()){
                    $recursiveHTML += Convert-FileDataToHTML -FilePath $file.Fullname
                }
            }
        }
    }
    else
    {
        $recursiveHTML += '<li><span class="caret">' + $folderItem.Name 
        if (!$SizeOff){
          $recursiveHTML += ' (<span style="color:red">' + $folderDetails.FolderSizeInUnits + " " + $folderDetails.Units + '</span>)</span>' + "`n"
         }
         else{
         $recursiveHTML += '</span>' + "`n"
         }
        $recursiveHTML += '<ul class="nested">'
        # Get all file data in subfolder
        $files = Get-ChildItem -Path $folderItem -File| Select-Object Name,FullName 
        $cont = ($files | Measure-Object  | Select-Object Count).Count 
        if ($cont -gt 0){
            If ($cont -eq 1){
                $recursiveHTML += Convert-FileDataToHTML -FilePath $files.Fullname
            }
        
            else{
                foreach($file in $files.GetEnumerator()){
                    $recursiveHTML += Convert-FileDataToHTML -FilePath $file.Fullname
                }
            }
        }
        $recursiveHTML += '</ul>'+ "`n"
    }

    #Recursively call this function for all subfolders
    foreach($subFolder in $subFolders)
    {   
        $recursiveHTML += GetAllFolderDetails -FolderPath $subFolder.FullName;
    }

    #Close up all tags
    if($subFolders.Count -gt 0)
    { 
        $recursiveHTML += '</ul>' + "`n"
    } 
    return $recursiveHTML
}

#Processing Starts Here

Set-Location -Path $startFolder

# delete old files
 Remove-Item -Path $destinationHTMLFile -Force
 New-Item $destinationHTMLFile -Type File
 Remove-Item -Path $imageDestinationFolder -Force -Recurse
#Opening html
$htmlLines += '<ul id="myUL">'+ "`n"

#This function call will return all of the recursive html for the starting folder and below
$htmlLines += GetAllFolderDetails -FolderPath $startFolder

#Closing html
$htmlLines += '</ul>'

# get date
$DateNow = Get-Date
$creationDate = $DateNow.tostring("dd-MMM-yyyy")

# delete the dummy html file
Remove-Item -Path $destinationHTMLFile -Force

#Get the html template, replace the template with generated code and write to the final html file
$sourceHTML = Get-Content -Path $sourceHTMLFile;
$destinationHTML = $sourceHTML.Replace('[FinalHTML]', $htmlLines).Replace('[CreationDateHTML]', $creationDate) 
$destinationHTML | Set-Content $destinationHTMLFile -encoding utf8

# copy image folder to destination
Copy-Item -Path $imageSourceFolder -Destination $destinationHTMLFolder -Recurse -Force
Set-ItemProperty $imageDestinationFolder Attributes -Value "Hidden"

Start-Process ((Resolve-Path "$destinationHTMLFile").Path)