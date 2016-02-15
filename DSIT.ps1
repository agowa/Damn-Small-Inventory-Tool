# Network Drives
Get-WmiObject -Class Win32_MappedLogicalDisk | Select-Object Name,ProviderName | Format-Table | Out-String -width 4096

# Powershell Version
$PSVersionTable.PSVersion

# Get all Printer
Get-WMIObject -Class Win32_Printer | Select-Object -Property Name,PortName,Default | Sort-Object Name |Sort-Object Default -Descending | Format-Table | Out-String -width 4096

# Get Office
Get-WmiObject -Class Win32_Product -Filter "name like '%office%'" | Select-Object -Property Vendor,Name,Version | Format-Table | Out-String -width 4096

# Get all Applications
Get-WmiObject -Class Win32_Product | Select-Object -Property Vendor,Name,Version | Format-Table | Out-String -width 4096

# Local Drives (with ntfs):
Get-WmiObject win32_volume -Filter "filesystem like 'ntfs'" | Sort-Object Name | Select-Object Name,Label | Format-Table | Out-String -width 4096

# Find all local PST-Files
# Get-ChildItem -Path C:\ -Filter *.pst -Recurse -ErrorAction SilentlyContinue
Get-WmiObject win32_volume -Filter "filesystem like 'ntfs' and DriveLetter like '%'" | foreach-object {Get-ChildItem -Path $_.name -Filter *.pst -Recurse -ErrorAction SilentlyContinue | Select-Object FullName | Format-Table | Out-String -width 4096}

# Processor and Windows Architecture
Get-WmiObject Win32_processor | Select-Object Name,Caption,AddressWidth,DataWidth | Format-Table | Out-String -width 4096
Get-WmiObject Win32_OperatingSystem | Select-Object OSArchitecture | Format-Table | Out-String -width 4096

# Get Bios Information
Get-WmiObject win32_bios | Select-Object SerialNumber,Manufacturer,BiosVersion,ReleaseDate,SMBIOSBIOSVersion,SMBIOSMajorVersion,SMBIOSMinorVersion | Format-Table | Out-String -width 4096
Get-WmiObject Win32_ComputerSystem | Select-Object Manufacturer,Model,NumberOfProcessors,NumberOfLogicalProcessors | Format-Table | Out-String -width 4096

# Environment Variables
Get-ChildItem Env: | Format-List | Out-String -width 4096

# # Get EventLog
# Get-EventLog -LogName * | Format-List
