# MIT License
# cscheunert 2021
# based on https://gist.github.com/milolav/3e296ed6a9f8a6c431a8553060f7514b

param (
  [string]$Imagefile,
  [string]$Disknumber,
  [string]$Indexnumber
)

#dism and diskpart require elevated privileges
if(-not [bool]([Security.Principal.WindowsIdentity]::GetCurrent().Groups -contains 'S-1-5-32-544')) {
  Write-Error "Script must be started as an Administrator"
}

# Warn the user
Write-Host ""
Write-Host -ForegroundColor Red "WARNING: This script can delete your files or operating system if you're not careful. I'm not responsible for any damage or lost files. Proceed at your own risk."
Write-Host ""
# If no parameters are supplied the script will ask for the information
if(!$ImageFile){
$ImageFile = Read-Host "Path to the image file (iso, wim or esd)"
$ImageFile = $ImageFile -replace('"','')	# the quotation marks break Get-ChildItem because Powershell is trying to read a drive with quotation mark before the drive letter but the quotation marks are unnecessary so they get removed
}

if(!$Disknumber){
	# External hard disk media
$DiskList = Get-WmiObject Win32_DiskDrive -Property * -Filter "MediaType = 'Fixed hard disk media' OR MediaType = 'External hard disk media'" | Select-Object Index,Model,@{Name="Size (GB)";Expression={[math]::truncate($_.Size/1GB)}} | Sort-Object Index 
$DiskOut = $DiskList | Out-String
$DiskChoice = $DiskList.Index | Out-String
Write-Host $DiskOut
do {
	try {
		$err = $false
		$Disknumber = Read-Host "Index number of the drive"
		if (-not $DiskChoice.Contains($DiskNumber)) { $err = $true ; "nicht enthalten"}
	}
	catch {
		$err = $true;
		}
	} while ($err)
}


$FileExtension = (Get-ChildItem $ImageFile).Extension

# Mount the iso file and get the right install file
if($FileExtension -eq ".iso"){
	$IsIso = "True"
	$WinIso = Resolve-Path $ImageFile
	Write-Progress "Mounting $WinIso ..."
	$MountResult = Mount-DiskImage -ImagePath $WinIso -StorageType ISO -PassThru
	$DriveLetter = ($MountResult | Get-Volume).DriveLetter
	if (-not $DriveLetter) {
		Write-Error "ISO file not loaded correctly" -ErrorAction Continue
		Dismount-DiskImage -ImagePath $WinIso | Out-Null
		return
	}
	
	$SourcesPath = $DriveLetter + ":\sources"
	
	# Check for the kind of install file in the iso
	if(Test-Path $SourcesPath\install.wim){
		$WinImage  = "$SourcesPath\install.wim"
	}
	elseif(Test-Path $SourcesPath\install.esd){
		$WinImage  = "$SourcesPath\install.esd"
	}
	else{
		Write-Host -ForegroundColor Yellow -BackgroundColor Red "ERROR: No install.wim or install.esd file located in $SourcesPath"
		return
	}
	
}
elseif($FileExtension -eq ".wim" -or $FileExtension -eq ".esd"){
	$IsIso = "False"
	$WinImage = Resolve-Path $ImageFile
	$WinImage = $WinImage.Path
}
else{
	Write-Host -ForegroundColor Yellow -BackgroundColor Red "ERROR: File must be an iso, wim or esd file"
	return
}


# Enumerating installation images
if(!$Indexnumber){
	Write-Host "Inspecting $WinImage"
	$WimOutput = dism /get-wiminfo /wimfile:`"$WinImage`" | Out-String
	$WimInfo = $WimOutput | Select-String "(?smi)Index : (?<Id>\d+).*?Name : (?<Name>[^`r`n]+)" -AllMatches
	if (!$WimInfo.Matches) {
	Write-Error "Images not found in install.wim`r`n$WimOutput" -ErrorAction Continue
	Dismount-DiskImage -ImagePath $WinIso | Out-Null
	return
	}
	
	$Items = @{ }
	$Menu = ""
	$DefaultIndex = 1
	$WimInfo.Matches | ForEach-Object { 
	$Items.Add([int]$_.Groups["Id"].Value, $_.Groups["Name"].Value)
	$Menu += $_.Groups["Id"].Value + ") " + $_.Groups["Name"].Value + "`r`n"
	if ($_.Groups["Name"].Value -eq $WinEdition) {
		$DefaultIndex = [int]$_.Groups["Id"].Value
	}
	}
	
	Write-Output $Menu
	do {
	try {
		$err = $false
		$WimIdx = if (([int]$val = Read-Host "Please select version [$DefaultIndex]") -eq "") { $DefaultIndex } else { $val }
		if (-not $Items.ContainsKey($WimIdx)) { $err = $true }
	}
	catch {
		$err = $true;
	}
	} while ($err)
	Write-Output $Items[$WimIdx]
}
else{
	$WimIdx = $Indexnumber
}


# find free letter for the efi drive
$drvlist = (Get-PSDrive -PSProvider filesystem).Name
Foreach ($drvletter in "ZYXWVUTSRQPONMLKJIHGFED".ToCharArray()) {
    If ($drvlist -notcontains $drvletter) {
        $EfiLetter = $drvletter
        break
    }
}

# find free letter for the windows drive
$drvlist = (Get-PSDrive -PSProvider filesystem).Name
$drvlist = $drvlist + $EfiLetter
Foreach ($drvletter in "ZYXWVUTSRQPONMLKJIHGFED".ToCharArray()) {
    If ($drvlist -notcontains $drvletter) {
        $WinLetter = $drvletter
        break
    }
}

$opt = $host.UI.PromptForChoice("Please confirm your choice" , "" , [System.Management.Automation.Host.ChoiceDescription[]] @("&Continue", "&Quit"), 1)
if ($opt -eq 1) {
  Dismount-DiskImage -ImagePath $WinIso | Out-Null -ErrorAction Continue
  return
}


#partition and format the drive
Write-Host "Partitioning and formating drive"
@"
select disk $DiskNumber
clean
convert gpt
select partition 1
delete partition override
create partition primary size=300
format quick fs=ntfs
create partition efi size=100
format quick fs=fat32
assign letter="$EfiLetter"
create partition msr size=128
create partition primary
format quick fs=ntfs
assign letter="$WinLetter"
exit
"@ | diskpart | Out-Null

#apply image using dism
Write-Host "Installing windows"
Invoke-Expression "dism /apply-image /imagefile:`"$WinImage`" /index:$WimIdx /applydir:$($WinLetter):\"
Write-Host ""
Invoke-Expression "$($WinLetter):\windows\system32\bcdboot $($WinLetter):\windows /f uefi /s $($EfiLetter):"
Write-Host ""
Invoke-Expression "bcdedit /store $($EfiLetter):\EFI\Microsoft\Boot\BCD"
Write-Host ""

# Cleanup
Write-Host "`r`nCleaning up..."
@"
select disk $disknumber
select partition 2
remove letter="$EfiLetter"
select partition 4
remove letter="$VirtualWinLetter"
exit
"@ | diskpart | Out-Null
if($IsIso -eq "True"){
	Dismount-DiskImage -ImagePath $WinIso | Out-Null
}

Write-Host "`r`nDone"

