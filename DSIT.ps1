############ Settings ##############
$InventoryPath = "E:\DSIT"
## TODO: $InventoryDevices = "E:\DSIT\hostnames.txt"
######### // Settings // ###########

############ Preinit ###############
Set-Location $InventoryPath
$computername = $env:COMPUTERNAME  ## later changes to the current computer in the $InventoryDevices list
mkdir "$InventoryPath\$computername"
######### // Preinit // ############

# Network Drives
Get-CimInstance -ClassName Win32_MappedLogicalDisk | Select-Object Name,ProviderName | Export-Csv -Encoding ASCII -Force -Path "$InventoryPath\$computername\Network_Drives.csv"

# Powershell Version
$PSVersionTable.PSVersion | Export-Csv -Encoding ASCII -Force -Path "$InventoryPath\$computername\Powershell_Version.csv"

# Get all Printer
Get-CimInstance -ClassName  Win32_Printer | Select-Object -Property Name,PortName,Default | Sort-Object Name |Sort-Object Default -Descending | Export-Csv -Encoding ASCII -Force -Path "$InventoryPath\$computername\Printer.csv"

# Get Office
Get-CimInstance -ClassName Win32_Product -Filter "name like '%office%'" | Select-Object -Property Vendor,Name,Version | Export-Csv -Encoding ASCII -Force -Path "$InventoryPath\$computername\Installed_Office_Products.csv"

# Get all Applications
Get-CimInstance -ClassName Win32_Product | Select-Object -Property Vendor,Name,Version | Export-Csv -Encoding ASCII -Force -Path "$InventoryPath\$computername\Installed_Applications.csv"

# Local Drives (with ntfs):
Get-CimInstance -ClassName win32_volume -Filter "filesystem like 'ntfs'" | Sort-Object Name | Select-Object Name,Label | Export-Csv -Encoding ASCII -Force -Path "$InventoryPath\$computername\Local_NTFS_Drives.csv"

# Find all local PST-Files
# Get-ChildItem -Path C:\ -Filter *.pst -Recurse -ErrorAction SilentlyContinue
Get-CimInstance -ClassName win32_volume -Filter "filesystem like 'ntfs' and DriveLetter like '%'" | foreach-object {Get-ChildItem -Path $_.name -Filter *.pst -Recurse -Force -ErrorAction SilentlyContinue | Select-Object FullName} | Export-Csv -Encoding ASCII -Force -Path "$InventoryPath\$computername\PST_Files.csv"

# Processor and Windows Architecture
Get-CimInstance -ClassName Win32_processor | Select-Object Name,Caption,AddressWidth,DataWidth | Export-Csv -Encoding ASCII -Force -Path "$InventoryPath\$computername\Processor.csv"
Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object OSArchitecture | Export-Csv -Encoding ASCII -Force -Path "$InventoryPath\$computername\OperatingSystem.csv"

# Get Bios Information
Get-CimInstance -ClassName win32_bios | Select-Object SerialNumber,Manufacturer,BiosVersion,ReleaseDate,SMBIOSBIOSVersion,SMBIOSMajorVersion,SMBIOSMinorVersion | Export-Csv -Encoding ASCII -Force -Path "$InventoryPath\$computername\Bios.csv"
Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object Manufacturer,Model,NumberOfProcessors,NumberOfLogicalProcessors | Export-Csv -Encoding ASCII -Force -Path "$InventoryPath\$computername\ComputerSystem.csv"

# Environment Variables
Get-ChildItem Env: | Export-Csv -Encoding ASCII -Force -Path "$InventoryPath\$computername\Environment.csv"

# # Get EventLog
mkdir $InventoryPath\$computername\Eventlog
Get-EventLog -LogName * | ForEach-Object {$p = $_; $p.Entries |Export-Csv -path $InventoryPath\$computername\Eventlog\$($p.LogDisplayName).csv}




# The MIT License (MIT)
# 
# Copyright (c) 2016 agowa338
# 
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
# 
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
# 
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
