write-host -f Cyan "`n`nPockEthernet Report Scanner Plus (Renames Files) -  By: Anthony Tippy`n`n"

if (Get-Module -ListAvailable -Name PSWritePDF) {
    Write-host -f Green "PSWritePDF Module Found!"
} 
else {
    Write-Host -f red "PSWritePDF Module Not Installed. Installing PSWritePDF Now!"
    Install-Module -Name PSWritePDF -Force
    exit
}

$path = "$home\scans\"
If(!(test-path $path))
{ Write-host "Scans Folder not found, creating folder now!"
      New-Item -ItemType Directory -Force -Path $path
}

#delete old report
Remove-Item "$home\Port Scan PDF Scrape Results.csv" -ErrorAction silentlyContinue

$ContractCode = Read-Host "`n`nEnter Project Title for scans (Will be added to each report filename):"

$duplicates = 0
$count = 0
$direct = get-childitem -Path "$home\scans\" | where {$_.extension -eq ".pdf"}| % { Write-Output $_.FullName } 

foreach ($d in $direct){
$DeviceResultItem = New-Object PSObject

write-host -f green "Reading: $d`n`n"

$PDF = Convert-PDFToText -FilePath $d
write-host "`n`n$PDF`n`n"

$DeviceResultItem | Add-Member -Name "Old Filename" -MemberType NoteProperty -Value $d

$date = [regex]::Match($PDF, "(?<=Date:\s)[\w-]+").value
$date = $date -replace " Address: Tag:",""
$author = [regex]::Match($PDF, "(?<=User:\s)(.*)(?=\sLocation:)").value
$Address = [regex]::Match($PDF, "(?<=Address:\s)(.*)(?=\sTag:)").value
$Room1 = [regex]::Match($PDF, "(?<=Location:\s)(.*)").value
$Room2 = [regex]::Match($PDF, "(?<=Location\sID:\s)(.*)(?=\sAddress:)").value
$Room = "$Room1 $Room2"
#$RoomOld = [regex]::Match($PDF, "(?<Room / Corridor:\s)[\w-]+").value
$building =[regex]::Match($PDF, "(?<=Building:\s)(.*)").value
$device = [regex]::Match($PDF, "(?<=Comment: )(.*)").value
$device = $device -replace "Port ID:",""
$DeviceID = [regex]::Match($PDF, "(?<=Device\sID: )(.*)").value
$VLAN1 = [regex]::Match($PDF, "(?<=Native\sVLAN: )(.*)").value
$VLAN2 = [regex]::Match($PDF, "(?<=Native\sVLAN\sID: )(.*)").value
$VLAN = "$VLAN1 $VLAN2"
$Port1 = [regex]::Match($PDF, "(?<=Port\sID:\s)(.*)(?=\sUntrust)").value
$Port2 = [regex]::Match($PDF, "(?<=Port\sID:\s)(.*)(?=\sCapabilities:)").value
$Port = "$Port1 $Port2"


#Rename pdfs
try{
$newfile = "$date-$ContractCode-$Room-$building-$device-Port Scan Report.pdf"
Rename-Item -Path $d -NewName "$newfile" -erroraction stop

$DeviceResultItem | Add-Member -Name "New Filename" -MemberType NoteProperty -Value $newfile
}
catch{
write-host -f red "Error: $d is a duplicate of another port scan!`n`n"
$newfile = "(Duplicate)-$date-$ContractCode-$Room-$building-$device-Port Scan Report.pdf"
Rename-Item -Path $d -NewName "$newfile" -force -erroraction continue
$duplicates = $duplicates +1
}
#>

# Table Coulumn 1 - Time
$DeviceResultItem | Add-Member -Name "Scan Date" -MemberType NoteProperty -Value $date
# Table Coulumn 1 - Time
$DeviceResultItem | Add-Member -Name "Author" -MemberType NoteProperty -Value $author
# Table Coulumn 1 - Time
$DeviceResultItem | Add-Member -Name "Address" -MemberType NoteProperty -Value $Address
# Table Coulumn 1 - Time
$DeviceResultItem | Add-Member -Name "Room" -MemberType NoteProperty -Value $Room
# Table Coulumn 1 - Time
$DeviceResultItem | Add-Member -Name "Device Scanned" -MemberType NoteProperty -Value $device
# Table Coulumn 1 - Time
$DeviceResultItem | Add-Member -Name "Switch Info" -MemberType NoteProperty -Value $DeviceID
# Table Coulumn 1 - Time
$DeviceResultItem | Add-Member -Name "Switch Port" -MemberType NoteProperty -Value $Port
# Table Coulumn 1 - Time
$DeviceResultItem | Add-Member -Name "VLAN" -MemberType NoteProperty -Value $VLAN

#Add line to the report
$DeviceResultsData = $DeviceResultItem


$DeviceResultsData | export-csv -Path "$home\Port Scan PDF Scrape Results.csv" -NoTypeInformation -Append -FORCE

$count = $count + 1
}
write-host "DONE!"
write-host "$count : Total PDF's Indexed`n`n"
write-host -f yellow "$duplicates Duplicate Reports Found`n`n"

invoke-item "$home\Port Scan PDF Scrape Results.csv"
import-csv "$home\Port Scan PDF Scrape Results.csv" | export-csv "$home\Port Scan PDF Scrape Results Log.csv" -Append -Force

write-host -f green "`nResults can be found at $home\Port Scan PDF Scrape Results.csv`n`n"

Read-Host -Prompt “Press Enter to exit”
