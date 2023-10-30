<#
The Psycho Kinectic Energy Meter (P.K.E. Meter) was one
of the Ghostbusters tools invented by Dr. Egon Spengler
that enabled them to track ghosts and other entities.
#>
mode con:cols=90 lines=62
Import-Module activedirectory
$user = $env:UserName
function ResetAccount($dn){
	write-host 'Removing AdminCount attribute...  ' -NoNewline
	try{
		Get-ADObject $dn  | set-adobject -Remove @{admincount=1}
		Write-Host 'OK' -ForegroundColor Green
	}
	catch{
		Write-Host 'Failed.' -ForegroundColor Red
		Write-Host $_ -ForegroundColor DarkRed
	}
	write-host 'Resettings ACL... ' -NoNewline
	try{
		$acl = Get-ACL -Path "AD:\$dn"
		$acl.SetAccessRuleProtection($False,$True)
		Set-Acl -Path "AD:\$dn" -AclObject $acl
		Write-Host 'OK' -ForegroundColor Green
	}
	catch{
		Write-Host 'Failed.' -ForegroundColor Red
		Write-Host $_ -ForegroundColor DarkRed
	}
}
<#----------------------------------------------------------------------------------------------------

Object type selection

----------------------------------------------------------------------------------------------------#>
do {
  [string]$object_type = 0
  while ( $object_type -lt 1 -or $object_type -gt 5) {
Write-Host '                                        _.-,'
Write-Host -ForegroundColor Red '                                   _' -NoNewline; Write-Host '.-`  / ' -NoNewline; Write-Host -ForegroundColor Red '.._'
Write-Host -ForegroundColor Red '                                .-:`' -NoNewline; Write-Host '/ - - \' -NoNewline; Write-Host -ForegroundColor Red ':::::-.'
Write-Host -ForegroundColor Red '                              .:::' -NoNewline; Write-Host ' `  e e  ` ' -NoNewline; Write-Host -ForegroundColor Red '`-::::.'
Write-Host -ForegroundColor Red '                             :::: ' -NoNewline; Write-Host '(    ^    )_' -NoNewline; Write-Host -ForegroundColor Red '.::::::'
Write-Host -ForegroundColor Red '                            ::::' -NoNewline; Write-Host '.` `.  o   `' -NoNewline; Write-Host -ForegroundColor Red '.::::`' -NoNewline; Write-Host '.`/_'
Write-Host '                        .  ' -NoNewline; Write-Host -ForegroundColor Red ':::' -NoNewline; Write-Host '.`       -  ' -NoNewline; Write-Host -ForegroundColor Red '.::::`' -NoNewline; Write-Host '_   _.:'
Write-Host '                      .-``---` .`|      ' -NoNewline; Write-Host -ForegroundColor Red '.::::`' -NoNewline; Write-Host '   ```' -NoNewline; Write-Host -ForegroundColor Red '::::'
Write-Host '                     `. ..-' -NoNewline; Write-Host -ForegroundColor Red ':::' -NoNewline; Write-Host '`  |    ' -NoNewline; Write-Host -ForegroundColor Red '.::::`        ::::'
Write-Host '                      `.` ' -NoNewline; Write-Host -ForegroundColor Red '::::' -NoNewline; Write-Host '    \ ' -NoNewline; Write-Host -ForegroundColor Red '.::::`          ::::'
Write-Host -ForegroundColor Red '                           ::::   .::::`           ::::'
Write-Host -ForegroundColor Red '                            ::::.::::`' -NoNewline; Write-Host '._' -NoNewline; Write-Host -ForegroundColor Red '          ::::'
Write-Host -ForegroundColor Red '                             ::::::`' -NoNewline; Write-Host ' /  `-      ' -NoNewline; Write-Host -ForegroundColor Red '.::::'
Write-Host -ForegroundColor Red '                              `::::' -NoNewline; Write-Host '-/__    __.-' -NoNewline; Write-Host -ForegroundColor Red ':::``'
Write-Host -ForegroundColor Red '                                `-:::::::::::::-``'
Write-Host -ForegroundColor Red '                                    ```::::``'
Write-Host '__________  ____  __.    ___________       _____          __                '
Write-Host '\______   \|    |/ _|    \_   _____/      /     \   _____/  |_  ___________ '
Write-Host ' |     ___/|      <       |    __)_      /  \ /  \_/ __ \   __\/ __ \_  __ \'
Write-Host ' |    |    |    |  \      |        \    /    Y    \  ___/|  | \  ___/|  | \/'
Write-Host ' |____| /\ |____|__ \ /\ /_______  /    \____|__  /\___  >__|  \___  >__|   '
Write-Host '        \/         \/ \/         \/             \/     \/          \/       '
Write-Host 'PKE Meter - Automated script to search "ghosts" and extract objects from Active Directory'
Write-Host '                                                                             Version 1.16'
Write-Host '                                                                Release date : 2023/10/29'
Write-Host '                                                                      Charles BLANC ROLIN'
Write-Host '                                                             https://github.com/woundride'
Write-Host '                                                                   https://www.apssis.com'
Write-Host '-----------------------------------------------------------------------------------------'
Write-Host '                                                  Licence Crative Commons CC BY-NC-SA 4.0'
Write-Host '                                       https://creativecommons.org/licenses/by-nc-sa/4.0/'
Write-Host '-----------------------------------------------------------------------------------------'
Write-Host ''
Write-Host 'Select object type :'
Write-Host ''
Write-Host '1 - Computer'
Write-Host '2 - User'
Write-Host '3 - Group'
Write-Host '4 - Configuration'
Write-Host ''
Write-Host -ForegroundColor Yellow '5 - Quit'
Write-Host ''
Write-Host '-----------------------------------------------------------------------------------------'
Write-Host ''
[string]$object_type = read-host 'Enter your selection'
Write-Host ''
Write-Host '-----------------------------------------------------------------------------------------'
Write-Host ''
switch ($object_type) {
<#----------------------------------------------------------------------------------------------------

Computer Selection

----------------------------------------------------------------------------------------------------#>
1{
	do {
		[string]$os_version = 0
		while ( $os_version -lt 1 -or $os_version -gt 20) {
	Write-Host '       ___________'
	Write-Host '      |.---------.|'
	Write-Host '      ||         ||'
	Write-Host '      ||         ||'
	Write-Host '      ||         ||		   _____                            _                '
	Write-Host '      | --------- |		  / ____|                          | |               '
	Write-Host '        )__ ____(		 | |     ___  _ __ ___  _ __  _   _| |_ ___ _ __ ___ '
	Write-Host '       [=== -- o ]--.		 | |    / _ \|  _   _ \|  _ \| | | | __/ _ \  __/ __|'
	Write-Host '     __|---------|__ \		 | |___| (_) | | | | | | |_) | |_| | ||  __/ |  \__ \'
	Write-Host '    [::::::::::: :::] )		  \_____\___/|_| |_| |_| .__/ \__,_|\__\___|_|  |___/'
	Write-Host '                    /T\		                       | |                           '
	Write-Host '                    \_/		                       |_|                           '
	Write-Host ''
	Write-Host '-----------------------------------------------------------------------------------------'
	Write-Host ''
	Write-Host 'Choose option :'
	Write-Host ''
	Write-Host ''
	Write-Host -ForegroundColor Red '1 - Windows 2000 + 2000 Server'
	Write-Host -ForegroundColor Red '2 - Windows XP (+ Embedded)'
	Write-Host -ForegroundColor Red '3 - Windows Vista'
	Write-Host -ForegroundColor Red '4 - Windows 7 (+ Embedded)'
	Write-Host -ForegroundColor Red '5 - Windows 8'
	Write-Host -ForegroundColor Red '6 - Windows 8.1'
	Write-Host -ForegroundColor Magenta '7 - Windows 10 / 11 --> New options menu'
	Write-Host ''
	Write-Host '-----------------------------------------------------------------------------------------'
	Write-Host ''
	Write-Host -ForegroundColor Red '8 - Windows Server 2003'
	Write-Host -ForegroundColor Red '9 - Windows Server 2008 + 2008 R2'
	Write-Host -ForegroundColor Red '10 - Windows Server 2012 + 2012 R2'
	Write-Host '11 - Windows Server 2016'
	Write-Host '12 - Windows Server 2019'
	Write-Host '13 - Windows Server 2022'
	Write-Host ''
	Write-Host -ForegroundColor Cyan '14 - All Systems in 1 file'
	Write-Host -ForegroundColor Red '15 - Inactive Computer(s) (not been connected for more than 3 months)'
	Write-Host -ForegroundColor Cyan '16 - Disabled Computer(s)'
	Write-Host -ForegroundColor Red '17 - Duplicated computers accounts'
	Write-Host -ForegroundColor Red '18 - All computers with unsupported OS'
	Write-Host ''
	Write-Host ''
	Write-Host -ForegroundColor Cyan '19 - Computers Statistics'
	Write-Host ''
	Write-Host -ForegroundColor Yellow '20 - Return'
	Write-Host
	Write-Host '-----------------------------------------------------------------------------------------'
	Write-Host ''
	[string]$os_version = read-host 'Enter your selection'
	Write-Host ''
	Write-Host '-----------------------------------------------------------------------------------------'
	Write-Host ''
	switch ($os_version) {
	1{
		Write-Host -ForegroundColor Red 'Windows 2000 + 2000 Server - End of life : 2010/07/13'
		Write-Host ''
		Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*2000*' } -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_2000.csv
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_2000.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	2{
		Write-Host -ForegroundColor Red 'Windows XP - End of life : 2014/04/08'
		Write-Host ''
		Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*XP*' } -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_xp.csv
		Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*Embedded*' -and OperatingSystemVersion -Like '*5.1*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_xp_embedded.csv
		Write-Host 'Exported files :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_xp.csv'
		Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_xp_embedded.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	3{
		Write-Host -ForegroundColor Red 'Windows Vista - End of life : 2017/04/11'
		Write-Host ''
		Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*vista*' } -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_vista.csv
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_vista.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	4{
		Write-Host -ForegroundColor Red 'Windows 7 - End of life : 2020/01/14'
		Write-Host ''
		Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*windows*7*' } -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_7.csv
		Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*Embedded*' -and OperatingSystemVersion -Like '*6.1*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_7_embedded.csv
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_7.csv'
		Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_7_embedded.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	5{
		Write-Host -ForegroundColor Red 'Windows 8 - End of life : 2016/01/12'
		Write-Host ''
		Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*windows*8*' -and OperatingSystemVersion -Like '*6.2*' } -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_8.csv
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_8.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	6{
		Write-Host -ForegroundColor Red 'Windows 8.1 - End of life : 2023/01/10'
		Write-Host ''
		Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*windows*8*' -and OperatingSystemVersion -Like '*6.3*' } -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_8-1.csv
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_8-1.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	7{
			do {
			[string]$10_version = 0
			while ( $10_version -lt 1 -or $10_version -gt 31) {
		Write-Host "    ___"
		Write-Host "'::|_|_|"
		Write-Host "'.:|_|_| Choose Windows 10 / 11 version :"
		Write-Host ''
		Write-Host ''
		Write-Host -ForegroundColor Red '1 - Windows 10 1507'
		Write-Host '2 - Windows 10 1507 LTSC'
		Write-Host -ForegroundColor Red '3 - Windows 10 1511'
		Write-Host -ForegroundColor Red '4 - Windows 10 1607'
		Write-Host '5 - Windows 10 1607 LTSC'
		Write-Host -ForegroundColor Red '6 - Windows 10 1703'
		Write-Host -ForegroundColor Red '7 - Windows 10 1709'
		Write-Host -ForegroundColor Red '8 - Windows 10 1803'
		Write-Host -ForegroundColor Red '9 - Windows 10 1803 Enterprise / Education'
		Write-Host -ForegroundColor Red '10 - Windows 10 1809'
		Write-Host -ForegroundColor Red '11 - Windows 10 1809 Enterprise / Education'
		Write-Host -ForegroundColor Green '12 - Windows 10 1809 LTSC'
		Write-Host -ForegroundColor Red '13 - Windows 10 1903 (19H1)'
		Write-Host -ForegroundColor Red '14 - Windows 10 1909 (19H2)'
		Write-Host -ForegroundColor Red '15 - Windows 10 1909 (19H2) Enterprise / Education'
		Write-Host -ForegroundColor Red '16 - Windows 10 2004 (20H1)'
		Write-Host -ForegroundColor Red '17 - Windows 10 20H2'
		Write-Host -ForegroundColor Red '18 - Windows 10 20H2 Enterprise / Education'
		Write-Host -ForegroundColor Red '19 - Windows 10 21H1'
		Write-Host -ForegroundColor Red '20 - Windows 11 21H2'
		Write-Host '21 - Windows 11 21H2 Enterprise / Education'
		Write-Host -ForegroundColor Red '22 - Windows 10 21H2'
		Write-Host '23 - Windows 10 21H2 Enterprise / Education'
		Write-Host -ForegroundColor Green '24 - Windows 10 21H2 LTSC'
		Write-Host '25 - Windows 11 22H2'
		Write-Host '26 - Windows 11 22H2 Enterprise / Education'
		Write-Host '27 - Windows 10 22H2'
		Write-Host '28 - Windows 10 22H2 Enterprise / Education'
		Write-Host ''
		Write-Host -ForegroundColor Cyan '29 - All Windows 10 versions in 1 file'
		Write-Host -ForegroundColor Cyan '30 - All Windows 10 versions in separated files'
		Write-Host ''
		Write-Host -ForegroundColor Yellow '31 - Return'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
		[string]$10_version = read-host 'Enter your selection'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
		switch ($10_version) {
			1{
				Write-Host -ForegroundColor Red 'Windows 10 1507 - End of life : 2017/05/09'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*LTSC*' -and OperatingSystemVersion -Like '*10240*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_10_1507.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_10_1507.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			2{
				Write-Host -ForegroundColor Green 'Windows 10 1507 LTSC - End of life : 2025/10/14'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*LTSC*' -and OperatingSystemVersion -Like '*10240*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\os_windows_10_1507_ltsc.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\os_windows_10_1507_ltsc.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			3{
				Write-Host -ForegroundColor Red 'Windows 10 1511 - End of life : 2017/10/10'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystemVersion -Like '*10586*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_10_1511.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_10_1511.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			4{
				Write-Host -ForegroundColor Red 'Windows 10 1607 - End of life : 2019/04/09'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*LTSC*' -and OperatingSystemVersion -Like '*14393*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_10_1607.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_10_1607.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			5{
				Write-Host -ForegroundColor Green 'Windows 10 1607 LTSC - End of life : 2026/10/13'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*LTSC*' -and OperatingSystemVersion -Like '*14393*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\os_windows_10_1607_ltsc.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\os_windows_10_1607_ltsc.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			6{
				Write-Host -ForegroundColor Red 'Windows 10 1703 - End of life : 2019/10/08'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystemVersion -Like '*15063*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_10_1703.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_10_1703.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			7{
				Write-Host -ForegroundColor Red 'Windows 10 1709 - End of life : 2020/10/13'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystemVersion -Like '*16299*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_10_1709.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_10_1709.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			8{
				Write-Host -ForegroundColor Red 'Windows 10 1803 - End of life : 2019/11/12'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -Like '*pro*' -and OperatingSystemVersion -Like '*17134*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_10_1803.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_10_1803.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			9{
				Write-Host -ForegroundColor Red 'Windows 10 1803 Enterprise / Education - End of life : 2021/05/11'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystemVersion -Like '*17134*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_10_1803_e.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_10_1803_e.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			10{
				Write-Host -ForegroundColor Red 'Windows 10 1809 - End of life : 2020/11/10'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -Like '*pro*' -and OperatingSystem -NotLike '*LTSC*' -and OperatingSystemVersion -Like '*17763*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_10_1809.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_10_1809.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			11{
				Write-Host -ForegroundColor Red 'Windows 10 1809 Enterprise / Education - End of life : 2021/05/11'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystem -NotLike '*LTSC*' -and OperatingSystemVersion -Like '*17763*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_10_1809_e.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_10_1809_e.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			12{
				Write-Host -ForegroundColor Green 'Windows 10 1809 LTSC - End of life : 2029/01/09'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*LTSC*' -and OperatingSystemVersion -Like '*17763*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\os_windows_10_1809_ltsc.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\os_windows_10_1809_ltsc.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			13{
				Write-Host -ForegroundColor Red 'Windows 10 1903 (19H1) - End of life : 2020/12/08'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystemVersion -Like '*18362*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_10_1903.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_10_1903.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			14{
				Write-Host -ForegroundColor Red 'Windows 10 1909 (19H2) - End of life : 2021/05/11'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -Like '*pro*' -and OperatingSystemVersion -Like '*18363*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_10_1909.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_10_1909.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			15{
				Write-Host -ForegroundColor Red 'Windows 10 1909 (19H2) Enterprise / Education - End of life : 2022/05/10'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystemVersion -Like '*18363*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_10_1909_e.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_10_1909_e.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			16{
				Write-Host -ForegroundColor Red 'Windows 10 2004 (20H1) - End of life : 2021/12/14'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystemVersion -Like '*19041*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_10_2004.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_10_2004.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			17{
				Write-Host -ForegroundColor Red 'Windows 10 20H2 - End of life : 2022/05/10'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -Like '*pro*' -and OperatingSystemVersion -Like '*19042*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_10_20H2.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_10_20H2.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			18{
				Write-Host -ForegroundColor Red 'Windows 10 20H2 Enterprise / Education - End of life : 2023/05/09'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystemVersion -Like '*19042*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_10_20H2_e.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_10_20H2_e.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			19{
				Write-Host -ForegroundColor Red 'Windows 10 21H1 - End of life : 2022/12/13'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystemVersion -Like '*19043*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_10_21H1.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_10_21H1.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			20{
				Write-Host -ForegroundColor Red 'Windows 11 21H2 - End of life : 2023/10/10'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*11*' -and OperatingSystem -Like '*pro*' -and OperatingSystemVersion -Like '*22000*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_11_21H2.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_11_21H2.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			21{
				Write-Host -ForegroundColor Green 'Windows 11 21H2 Enterprise / Education - End of life : 2024/10/08'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*11*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystemVersion -Like '*22000*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\os_windows_11_21H2_e.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\os_windows_11_21H2_e.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			22{
				Write-Host -ForegroundColor Red 'Windows 10 21H2 - End of life : 2023/06/13'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -Like '*pro*' -and OperatingSystem -NotLike '*LTSC*' -and OperatingSystemVersion -Like '*19044*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_10_21H2.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_10_21H2.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			23{
				Write-Host -ForegroundColor Green 'Windows 10 21H2 Enterprise / Education - End of life : 2024/06/11'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystem -NotLike '*LTSC*' -and OperatingSystemVersion -Like '*19044*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\os_windows_10_21H2_e.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\os_windows_10_21H2_e.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			24{
				Write-Host -ForegroundColor Green 'Windows 10 21H2 LTSC - End of life : 2027/01/12'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*LTSC*' -and OperatingSystemVersion -Like '*19044*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\os_windows_10_21H2_ltsc.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\os_windows_10_21H2_ltsc.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			25{
				Write-Host -ForegroundColor Green 'Windows 11 22H2 - End of life : 2024/10/08'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*11*' -and OperatingSystemVersion -Like '*22621*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\os_windows_11_22H2.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\os_windows_11_22H2.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			26{
				Write-Host -ForegroundColor Green 'Windows 11 22H2 Enterprise / Education - End of life : 2025/10/14'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*11*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystemVersion -Like '*22621*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\os_windows_11_22H2_e.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\os_windows_11_22H2_e.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			27{
				Write-Host -ForegroundColor Green 'Windows 10 22H2 - End of life : 2024/05/14'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*11*' -and OperatingSystemVersion -Like '*19045*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\os_windows_10_22H2.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\os_windows_10_22H2.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			28{
				Write-Host -ForegroundColor Green 'Windows 10 22H2 Enterprise / Education - End of life : 2025/05/13'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*11*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystemVersion -Like '*19045*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\os_windows_10_22H2_e.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\os_windows_10_22H2_e.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			29{
				Write-Host -ForegroundColor Cyan 'All Windows 10 versions in 1 file'
				Write-Host ''
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\os_windows_10.csv
				Write-Host 'Exported file :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\os_windows_10.csv'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
			30{
				Write-Host -ForegroundColor Cyan 'All Windows 10 versions in separated files'
				Write-Host ''
				Mkdir C:\Users\$user\Desktop\Windows_10
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*LTSC*' -and OperatingSystemVersion -Like '*10240*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Windows_10\unsupported_os_windows_10_1507.csv
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*LTSC*' -and OperatingSystemVersion -Like '*10240*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Windows_10\os_windows_10_1507_ltsc.csv
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystemVersion -Like '*10586*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Windows_10\unsupported_os_windows_10_1511.csv
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*LTSC*' -and OperatingSystemVersion -Like '*14393*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Windows_10\unsupported_os_windows_10_1607.csv
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*LTSC*' -and OperatingSystemVersion -Like '*14393*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Windows_10\os_windows_10_1607_ltsc.csv
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystemVersion -Like '*15063*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Windows_10\unsupported_os_windows_10_1703.csv
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystemVersion -Like '*16299*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Windows_10\unsupported_os_windows_10_1709.csv
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -Like '*pro*' -and OperatingSystemVersion -Like '*17134*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Windows_10\unsupported_os_windows_10_1803.csv
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystemVersion -Like '*17134*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Windows_10\unsupported_os_windows_10_1803_e.csv
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -Like '*pro*' -and OperatingSystem -NotLike '*LTSC*' -and OperatingSystemVersion -Like '*17763*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Windows_10\unsupported_os_windows_10_1809.csv
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystem -NotLike '*LTSC*' -and OperatingSystemVersion -Like '*17763*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Windows_10\unsupported_os_windows_10_1809_e.csv
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*LTSC*' -and OperatingSystemVersion -Like '*17763*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Windows_10\os_windows_10_1809_ltsc.csv
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystemVersion -Like '*18362*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Windows_10\unsupported_os_windows_10_1903.csv
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -Like '*pro*' -and OperatingSystemVersion -Like '*18363*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Windows_10\unsupported_os_windows_10_1909.csv
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystemVersion -Like '*18363*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Windows_10\unsupported_os_windows_10_1909_e.csv
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystemVersion -Like '*19041*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Windows_10\unsupported_os_windows_10_2004.csv
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -Like '*pro*' -and OperatingSystemVersion -Like '*19042*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Windows_10\unsupported_os_windows_10_20H2.csv
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystemVersion -Like '*19042*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Windows_10\unsupported_os_windows_10_20H2_e.csv
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystemVersion -Like '*19043*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Windows_10\unsupported_os_windows_10_21H1.csv
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -Like '*pro*' -and OperatingSystem -NotLike '*LTSC*' -and OperatingSystemVersion -Like '*19044*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Windows_10\unsupported_os_windows_10_21H2.csv
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystem -NotLike '*LTSC*' -and OperatingSystemVersion -Like '*19044*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Windows_10\os_windows_10_21H2_e.csv
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*LTSC*' -and OperatingSystemVersion -Like '*19044*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Windows_10\os_windows_10_21H2_ltsc.csv
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*11*' -and OperatingSystemVersion -Like '*19045*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Windows_10\os_windows_10_22H2.csv
				Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*11*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystemVersion -Like '*19045*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Windows_10\os_windows_10_22H2_e.csv
				Write-Host ''
				Write-Host 'All files are exported in :'
				Write-Host ''
				Write-Host 'C:\Users\'$user'\Desktop\Windows_10'
				Write-Host ''
				Write-Host '-----------------------------------------------------------------------------------------'
				Write-Host ''
			}
	}
	}
	} while ( $10_version -ne 31 ) }
	8{
		Write-Host -ForegroundColor Red 'Windows Server 2003 - End of life : 2016/10/09'
		Write-Host ''
		Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*2003*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_server_2003.csv
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_server_2003.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	9{
		Write-Host -ForegroundColor Red 'Windows Server 2008 + 2008 R2 - End of life : 2020/01/14'
		Write-Host ''
		Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*2008*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_server_2008.csv
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_server_2008.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	10{
		Write-Host -ForegroundColor Red 'Windows Server 2012 + 2012 R2 - End of life : 2023/10/10'
		Write-Host ''
		Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*2012*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\unsupported_os_windows_server_2012.csv
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\unsupported_os_windows_server_2012.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	11{
		Write-Host -ForegroundColor Green 'Windows Server 2016 - End of life : 2027/01/12'
		Write-Host ''
		Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*2016*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\os_windows_server_2016.csv
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\os_windows_server_2016.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	12{
		Write-Host -ForegroundColor Green 'Windows Server 2019 - End of life : 2029/01/09'
		Write-Host ''
		Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*2019*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\os_windows_server_2019.csv
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\os_windows_server_2019.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	13{
		Write-Host -ForegroundColor Green 'Windows Server 2022 - End of life : 2031/10/14'
		Write-Host ''
		Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*2022*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\os_windows_server_2022.csv
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\os_windows_server_2022.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	14{
		Write-Host -ForegroundColor Cyan 'All Systems in 1 file (enabled)'
		Write-Host ''
		Get-ADComputer -Filter {enabled -eq 'true'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\all_computers.csv
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\all_computers.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	15{
		Write-Host -ForegroundColor Cyan 'All computers who have not been connected for more than 3 months'
		Search-ADAccount -ComputersOnly -AccountInActive -TimeSpan 90:00:00:00 -ResultPageSize 2000 -ResultSetSize $null | ?{$_.Enabled -eq $True} | Select-Object Name, SamAccountName, LastLogonDate, DistinguishedName | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\inactive_computers.csv -NoTypeInformation
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\inactive_computers.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
		if (( Read-Host 'Do you want to disable inactive accounts ? - Administrator rights needed - Answer "sure" to validate' ) -eq 'sure' ){
				Search-ADAccount -ComputersOnly -AccountInActive -TimeSpan 90:00:00:00 -ResultPageSize 2000 -ResultSetSize $null | ?{$_.Enabled -eq $True} | Select-Object SamAccountName | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\inactive_computers_to_disable.csv -NoTypeInformation
				$file_comput_dis=import-csv C:\Users\$user\Desktop\inactive_computers_to_disable.csv 
				foreach ($entry in $file_comput_dis) 
				{
					Set-ADComputer -Identity $($entry.SamAccountName) -Enabled $false 
				}
		}
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	16{
		Write-Host -ForegroundColor Cyan 'All disabled computers'
		Write-Host ''
		Get-ADComputer -Filter {enabled -eq 'false'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\all_disabled_computers.csv
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\all_disabled_computers.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	17{
		Write-Host -ForegroundColor Red 'Duplicated computers accounts'
		Write-Host ''
		Get-adobject -ldapfilter '(cn=*cnf:*)' ; Get-adobject -ldapfilter '(sAMAccountName=$duplicate)' | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\duplicated_computers_accounts.csv
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\duplicated_computers_accounts.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
		18{
				Write-Host -ForegroundColor Red 'All computers with unsupported OS in separated files'
			Write-Host ''
			Mkdir C:\Users\$user\Desktop\Unsupported_OS
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*2000*' } -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_2000.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*XP*' } -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_xp.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*Embedded*' -and OperatingSystemVersion -Like '*5.1*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_xp_embedded.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*vista*' } -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_vista.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*windows*7*' } -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_7.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*Embedded*' -and OperatingSystemVersion -Like '*6.1*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_7_embedded.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*windows*8*' -and OperatingSystemVersion -Like '*6.2*' } -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_8.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*windows*8*' -and OperatingSystemVersion -Like '*6.3*' } -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_8-1.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*LTSC*' -and OperatingSystemVersion -Like '*10240*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_10_1507.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystemVersion -Like '*10586*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_10_1511.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*LTSC*' -and OperatingSystemVersion -Like '*14393*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_10_1607.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystemVersion -Like '*15063*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_10_1703.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystemVersion -Like '*16299*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_10_1709.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -Like '*pro*' -and OperatingSystemVersion -Like '*17134*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_10_1803.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystemVersion -Like '*17134*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_10_1803_e.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -Like '*pro*' -and OperatingSystem -NotLike '*LTSC*' -and OperatingSystemVersion -Like '*17763*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_10_1809.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystem -NotLike '*LTSC*' -and OperatingSystemVersion -Like '*17763*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_10_1809_e.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystemVersion -Like '*18362*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_10_1903.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -Like '*pro*' -and OperatingSystemVersion -Like '*18363*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_10_1909.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystemVersion -Like '*18363*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_10_1909_e.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystemVersion -Like '*19041*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_10_2004.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -Like '*pro*' -and OperatingSystemVersion -Like '*19042*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_10_20H2.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystemVersion -Like '*19042*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_10_20H2_e.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystemVersion -Like '*19043*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_10_21H1.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*2003*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_server_2003.csv
			Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*2008*'} -Properties SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Select SamAccountName, LastLogonDate, OperatingSystem, OperatingSystemVersion, CanonicalName | Export-CSV -Encoding UTF8 C:\Users\$user\Desktop\Unsupported_OS\unsupported_os_windows_server_2008.csv
			Write-Host ''
			Write-Host 'All files are exported in :'
			Write-Host ''
			Write-Host 'C:\Users\'$user'\Desktop\Unsupported_OS'
			Write-Host ''
			Write-Host '-----------------------------------------------------------------------------------------'
			Write-Host ''
		}
	19{
		Write-Host -ForegroundColor Cyan 'Computers Statistics (Number of computers by OS)'
		$all_computers = Get-ADComputer -Filter {enabled -eq 'true'}
		$innactive_computers = Search-ADAccount -ComputersOnly -AccountInActive -TimeSpan 90:00:00:00 -ResultPageSize 2000 -ResultSetSize $null | ?{$_.Enabled -eq $True}
		$disabled_computers = Get-ADComputer -Filter {enabled -eq 'false'}
		$windows2000 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*2000*' }
		$windowsxp = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*XP*' }
		$windowsxpembedded = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*Embedded*' -and OperatingSystemVersion -Like '*5.1*'}
		$windowsvista = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*vista*' }
		$windows7 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*windows*7*' }
		$windows7embedded = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*Embedded*' -and OperatingSystemVersion -Like '*6.1*'}
		$windows8 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*windows*8*' -and OperatingSystemVersion -Like '*6.2*' }
		$windows8_1 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*windows*8*' -and OperatingSystemVersion -Like '*6.3*' }
		$windows10_1507 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*LTSC*' -and OperatingSystemVersion -Like '*10240*'}
		$windows10_1507_ltsc = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*LTSC*' -and OperatingSystemVersion -Like '*10240*'}
		$windows10_1511 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystemVersion -Like '*10586*'}
		$windows10_1607 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*LTSC*' -and OperatingSystemVersion -Like '*14393*'}
		$windows10_1607_ltsc = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*LTSC*' -and OperatingSystemVersion -Like '*14393*'}
		$windows10_1703 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystemVersion -Like '*15063*'}
		$windows10_1709 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystemVersion -Like '*16299*'}
		$windows10_1803 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -Like '*pro*' -and OperatingSystemVersion -Like '*17134*'}
		$windows10_1803e = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystemVersion -Like '*17134*'}
		$windows10_1809 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -Like '*pro*' -and OperatingSystem -NotLike '*LTSC*' -and OperatingSystemVersion -Like '*17763*'}
		$windows10_1809e = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystem -NotLike '*LTSC*' -and OperatingSystemVersion -Like '*17763*'}
		$windows10_1809_ltsc = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*LTSC*' -and OperatingSystemVersion -Like '*17763*'}
		$windows10_1903 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystemVersion -Like '*18362*'}
		$windows10_1909 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -Like '*pro*' -and OperatingSystemVersion -Like '*18363*'}
		$windows10_1909_e = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystemVersion -Like '*18363*'}
		$windows10_2004 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystemVersion -Like '*19041*'}
		$windows10_20h2 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -Like '*pro*' -and OperatingSystemVersion -Like '*19042*'}
		$windows10_20h2_e = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystemVersion -Like '*19042*'}
		$windows10_21h1 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystemVersion -Like '*19043*'}
		$windows10_21h2 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -Like '*pro*' -and OperatingSystem -NotLike '*LTSC*' -and OperatingSystemVersion -Like '*19044*'}
		$windows10_21h2_e = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystem -NotLike '*LTSC*' -and OperatingSystemVersion -Like '*19044*'}
		$windows10_21h2_ltsc = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*10*LTSC*' -and OperatingSystemVersion -Like '*19044*'}
		$windows11_21h2 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*11*' -and OperatingSystem -Like '*pro*' -and OperatingSystemVersion -Like '*22000*'}
		$windows11_21h2_e = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*11*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystemVersion -Like '*22000*'}
		$windows11_22h2 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*11*' -and OperatingSystem -Like '*pro*' -and OperatingSystemVersion -Like '*22621*'}
		$windows11_22h2_e = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*11*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystemVersion -Like '*22621*'}
		$windows10_22h2 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*11*' -and OperatingSystem -Like '*pro*' -and OperatingSystemVersion -Like '*19045*'}
		$windows10_22h2_e = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*11*' -and OperatingSystem -NotLike '*pro*' -and OperatingSystemVersion -Like '*19045*'}
		$windows2003 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*2003*'}
		$windows2008 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*2008*'}
		$windows2012 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*2012*'}
		$windows2016 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*2016*'}
		$windows2019 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*2019*'}
		$windows2022 = Get-ADComputer -Filter {enabled -eq 'true' -and OperatingSystem -Like '*2022*'}
		
		Write-Output '-----------------------------------------------------------------------------------------' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt
		Write-Output 'Computers Statistics' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '-----------------------------------------------------------------------------------------' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Number of computers :' $all_computers.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'All computers who have not been connected for more than 3 months :' $innactive_computers.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'All disabled computers :' $disabled_computers.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '-----------------------------------------------------------------------------------------' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Client OS' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '-----------------------------------------------------------------------------------------' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows 2000 + 2000 Server - End of life : 2010/07/13 :' $windows2000.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows XP - End of life : 2014/04/08 :' $windowsxp.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows XP Embedded - End of life : 2014/04/08 :' $windowsxpembedded.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows Vista - End of life : 2017/04/11 :' $windowsvista.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows 7 - End of life : 2020/01/14 :' $windows7.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows 7 Embedded - End of life : 2020/01/14 :' $windows7embedded.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows 8 - End of life : 2016/01/12 :' $windows8.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows 8.1 - End of life : 2023/01/10 :' $windows8_1.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows 10 1507 - End of life : 2017/05/09 :' $windows10_1507.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Windows 10 1507 LTSC - End of life : 2025/10/14 :' $windows10_1507_ltsc.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows 10 1511 - End of life : 2017/10/10 :' $windows10_1511.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows 10 1607 - End of life : 2019/04/09 :' $windows10_1607.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Windows 10 1607 LTSC - End of life : 2026/10/13 :' $windows10_1607_ltsc.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows 10 1703 - End of life : 2019/10/08 :' $windows10_1703.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows 10 1709 - End of life : 2020/10/13 :' $windows10_1709.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows 10 1803 - End of life : 2019/11/12 :' $windows10_1803.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows 10 1803 Enterprise / Education - End of life : 2021/05/11 :' $windows10_1809e.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows 10 1809 - End of life : 2020/11/10 :' $windows10_1809.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows 10 1809 Enterprise / Education - End of life : 2021/05/11 :' $windows10_1809e.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Windows 10 1809 LTSC - End of life : 2029/01/09 :' $windows10_1809_ltsc.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows 10 1903 (19H1) - End of life : 2020/12/08 :' $windows10_1903.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows 10 1909 (19H2) - End of life : 2021/05/11 :' $windows10_1909.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported Windows : 10 1909 (19H2) Enterprise / Education - End of life : 2022/05/10 :' $windows10_1909_e.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows 10 2004 (20H1) - End of life : 2021/12/14 :' $windows10_2004.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows 10 20H2 - End of life : 2022/05/10 :' $windows10_20h2.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows 10 20H2 Enterprise / Education - End of life : 2023/05/09 :' $windows10_20h2_e.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows 10 21H1 - End of life : 2022/12/13 :' $windows10_21h1.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows 11 21H2 - End of life : 2023/10/10 :' $windows11_21h2.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Windows 11 21H2 Enterprise / Education - End of life : 2024/10/08 :' $windows11_21h2_e.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows 10 21H2 - End of life : 2023/06/13 :' $windows10_21h2.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Windows 10 21H2 Enterprise / Education - End of life : 2024/06/11 :' $windows10_21h2_e.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Windows 10 21H2 LTSC - End of life : 2027/01/12 :' $windows10_21h2_ltsc.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Windows 11 22H2 - End of life : 2024/10/08 :' $windows11_22h2.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Windows 11 22H2 Enterprise / Education - End of life : 2025/10/14 :' $windows11_22h2_e.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Windows 10 22H2 - End of life : 2024/05/14 :' $windows10_22h2.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Windows 10 22H2 Enterprise / Education - End of life : 2025/05/13 :' $windows10_22h2_e.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '-----------------------------------------------------------------------------------------' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Server OS' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '-----------------------------------------------------------------------------------------' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows Server 2003 - End of life : 2016/10/09 :' $windows2003.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows Server 2008 + 2008 R2 - End of life : 2020/01/14 :' $windows2008.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Unsupported OS : Windows Server 2012 + 2012 R2 - End of life : 2023/10/10 :' $windows2012.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Windows Server 2016 - End of life : 2027/01/12 :' $windows2016.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Windows Server 2019 - End of life : 2029/01/09 :' $windows2019.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Output 'Windows Server 2022 - End of life : 2031/10/14 :' $windows2022.count | Out-File -FilePath c:\Users\$user\Desktop\computers_statistics.txt -Append
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\computers_statistics.txt'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	}
	}
	} while ( $os_version -ne 20 ) }
<#----------------------------------------------------------------------------------------------------

User Selection

----------------------------------------------------------------------------------------------------#>
2{
	do {
  [string]$user_extract = 0
  while ( $user_extract -lt 1 -or $user_extract -gt 23) {
	Write-Host '      ////^\\\\'
	Write-Host '      | ^   ^ |'
	Write-Host '     @ (o) (o) @'
	Write-Host '      |   <   |'
	Write-Host '      |  ___  |'
	Write-Host '       \_____/'
	Write-Host '     ____|  |____'
	Write-Host '    /    \__/    \'
	Write-Host '   /              \		  _    _                   '
	Write-Host '  /\_/|        |\_/\		 | |  | |                  '
	Write-Host ' / /  |        |  \ \		 | |  | |___  ___ _ __ ___ '
	Write-Host '( <   |        |   > )		 | |  | / __|/ _ \  __/ __|'
	Write-Host ' \ \  |        |  / /		 | |__| \__ \  __/ |  \__ \'
	Write-Host '  \ \ |________| / /		  \____/|___/\___|_|  |___/'
	Write-Host '   \ \|						'
	Write-Host ''
	Write-Host '-----------------------------------------------------------------------------------------'
	Write-Host ''
	Write-Host 'Choose user extract type :'
	Write-Host ''
	Write-Host ''
	Write-Host -ForegroundColor Red '1 - AdminSDHolder Account(s) [previously belonged to an administrator group]'
	Write-Host '2 - Administrator Account(s) [belonging to an administrator group + adminCount = 1]'
	Write-Host -ForegroundColor Red '3 - Administrator Account(s) [in administrator group] without adminCount = 1'
	Write-Host -ForegroundColor Red '4 - Administrator Account(s) with mailbox'
	Write-Host '5 - Administrator Group(s)'
	Write-Host -ForegroundColor Red '6 - Inactive user(s) (not been connected for more than 3 months)'
	Write-Host -ForegroundColor Red '7 - Inactive user(s) (not been connected for more than 6 months)'
	Write-Host -ForegroundColor Red '8 - Never connected user(s)'
	Write-Host -ForegroundColor Cyan '9 - All user accounts enabled and not expired'
	Write-Host -ForegroundColor Red '10 - All enabled user accounts that never exire'
	Write-Host -ForegroundColor Cyan '11 - Disabled user(s)'
	Write-Host -ForegroundColor Red '12 - All expired user accounts (not disabled)'
	Write-Host -ForegroundColor Red '13 - All users whose password never expire'
	Write-Host -ForegroundColor Red '14 - Users who cannot change their password'
	Write-Host '15 - Users with mailbox'
	Write-Host '16 - Users without mailbox'
	Write-Host -ForegroundColor Cyan '17 - Password last set for all users accounts (enabled)'
	Write-Host -ForegroundColor Cyan '18 - Password last set for a specific user account'
	Write-Host -ForegroundColor Cyan '19 - Last logon date for a specific user account'
	Write-Host -ForegroundColor Cyan '20 - Search user by EmployeeID'
	Write-Host -ForegroundColor Cyan '21 - Connections history for a specific user'
	Write-Host -ForegroundColor Cyan '22 - Check inheritance ACL for a specific user'
	Write-Host ''
	Write-Host -ForegroundColor Yellow '23 - Return'
	Write-Host ''
	Write-Host '-----------------------------------------------------------------------------------------'
	Write-Host ''
	[string]$user_extract = read-host 'Enter your selection'
	Write-Host ''
	Write-Host '-----------------------------------------------------------------------------------------'
	Write-Host ''
	switch ($user_extract) {
	1{
		$accList = Get-ADObject -filter 'AdminCount -eq 1 -and isCriticalsystemObject -notlike "*"' -properties * 
		$adminGroupList = get-adgroup -filter 'admincount -eq 1 -and iscriticalsystemobject -like "*"'` | select -ExpandProperty distinguishedName
		$counter=0
		$orphanList = @()
		foreach($acc in $accList ){
			$counter++
			$dn = $acc.DistinguishedName
			$memberOf = Get-ADPrincipalGroupMembership -Identity $dn
			foreach( $group in $memberOf ){
				$isAdmin = $adminGroupList.Contains($group.DistinguishedName)
				if ( $isAdmin ){ break }
			}
			if ( $isAdmin ){
			}
			else{
				$orphanList += $dn
			}
		}
		$orphanList | Out-File c:\Users\$user\Desktop\adminsdholder.txt
		Write-Host -ForegroundColor Cyan 'Found '$orphanList.count' AdminSDHolder account(s)'
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\adminsdholder.txt'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
		foreach( $acc in $orphanList ){
			if (( Read-Host 'Reset security descriptor for account '$($acc)'? - Administrator rights needed - Answer y' ) -eq 'y' ){
				ResetAccount $dn
			}
		}

	}
	2{
		$accList = Get-ADObject -filter 'AdminCount -eq 1 -and isCriticalsystemObject -notlike "*"' -properties * 
		$adminGroupList = get-adgroup -filter 'admincount -eq 1 -and iscriticalsystemobject -like "*"'` | select -ExpandProperty distinguishedName
		$counter=0
		$admin_list=@()
		foreach($acc in $accList ){
			$counter++
			$dnr = $acc.samaccountname
			$memberOf = Get-ADPrincipalGroupMembership -Identity $dnr
			$data = foreach( $group in $memberOf ){
				$isAdmin = $adminGroupList.Contains($group.DistinguishedName)
				if ( $isAdmin ){ break }
			}
		if ( $isAdmin ){
		$admin_list += $dnr
		}
		}
		$admin_list | Out-File c:\Users\$user\Desktop\administrator_users.txt
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\administrator_users.txt'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	3{
		$accList = Get-ADUser -Filter {-not (AdminCount -like "1")}
		$adminGroupList = get-adgroup -filter 'admincount -eq 1 -and iscriticalsystemobject -like "*"'` | select -ExpandProperty distinguishedName
		$counter=0
		$orphanList = @()
		foreach($acc in $accList ){
			$counter++
			$dn = $acc.DistinguishedName
			$memberOf = Get-ADPrincipalGroupMembership -Identity $dn
			foreach( $group in $memberOf ){
				$isAdmin = $adminGroupList.Contains($group.DistinguishedName)
				if ( $isAdmin ){ break }
			}
			if ( $isAdmin ){
				$orphanList += $dn
			}
		}
		$orphanList | Out-File c:\Users\$user\Desktop\admin_without_admincount_equal_to_1.txt
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\admin_without_admincount_equal_to_1.txt'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	4{
		$accList = Get-ADObject -filter 'AdminCount -eq 1 -and isCriticalsystemObject -notlike "*" -and mail -like "*"' -properties * 
		$adminGroupList = get-adgroup -filter 'admincount -eq 1 -and iscriticalsystemobject -like "*"'` | select -ExpandProperty distinguishedName
		$admin_list=@()
		foreach($acc in $accList ){
			$dn = $acc.DistinguishedName
			$memberOf = Get-ADgroup -Filter {member -RecursiveMatch $dn} 
			$data = foreach( $group in $memberOf ){
				$isAdmin = $adminGroupList.Contains($group.DistinguishedName)
				if ( $isAdmin ){ break }
			}
		if ( $isAdmin ){
		$admin_list += $acc.sAMAccountName
		}
		}
		$admin_list | Out-File c:\Users\$user\Desktop\administrator_users_with_mailbox.txt
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\administrator_users_with_mailbox.txt'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	5{
		Write-Host 'Administrator Group(s)'
		$adminGroupList = get-adgroup -filter 'admincount -eq 1 -and iscriticalsystemobject -like "*"'` | select distinguishedName | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\admin_groups.csv
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\admin_groups.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	6{
		Write-Host -ForegroundColor Cyan 'All enabled and not expired user accounts who have not been connected for more than 3 months'
		$innactive_time_3_m = (Get-Date).Adddays(-(90))
		Get-ADUser -filter {LastLogonTimeStamp -lt $innactive_time_3_m -and enabled -eq $true} -properties * | where { $_.AccountExpirationDate -gt (Get-Date) -or $_.AccountExpirationDate -eq $null } | Select-Object Name, SamAccountName, LastLogonDate, DistinguishedName, EmployeeId | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\inactive_users_3_months.csv -NoTypeInformation
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\inactive_users_3_months.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
		if (( Read-Host 'Do you want to disable inactive accounts ? - Administrator rights needed - Answer "sure" to validate' ) -eq 'sure' ){
				Get-ADUser -filter {LastLogonTimeStamp -lt $innactive_time_3_m -and enabled -eq $true} -properties * | where { $_.AccountExpirationDate -gt (Get-Date) -or $_.AccountExpirationDate -eq $null } | Select-Object SamAccountName | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\inactive_users_to_disable.csv -NoTypeInformation
				$file_us_dis=import-csv C:\Users\$user\Desktop\inactive_users_to_disable.csv 
				foreach ($entry in $file_us_dis) 
				{
					Set-ADUser -Identity $($entry.SamAccountName) -Enabled $false 
				}
		}
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	7{
		Write-Host -ForegroundColor Cyan 'All enabled and not expired user accounts who have not been connected for more than 6 months'
		$innactive_time_6_m = (Get-Date).Adddays(-(180))
		Get-ADUser -filter {LastLogonTimeStamp -lt $innactive_time_6_m -and enabled -eq $true} -properties * | where { $_.AccountExpirationDate -gt (Get-Date) -or $_.AccountExpirationDate -eq $null } | Select-Object Name, SamAccountName, LastLogonDate, DistinguishedName, EmployeeId | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\inactive_users_6_months.csv -NoTypeInformation
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\inactive_users_6_months.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
		if (( Read-Host 'Do you want to disable inactive accounts ? - Administrator rights needed - Answer "sure" to validate' ) -eq 'sure' ){
				Get-ADUser -filter {LastLogonTimeStamp -lt $innactive_time_6_m -and enabled -eq $true} -properties * | where { $_.AccountExpirationDate -gt (Get-Date) -or $_.AccountExpirationDate -eq $null } | Select-Object SamAccountName | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\inactive_users_to_disable.csv -NoTypeInformation
				$file_us_dis=import-csv C:\Users\$user\Desktop\inactive_users_to_disable.csv 
				foreach ($entry in $file_us_dis) 
				{
					Set-ADUser -Identity $($entry.SamAccountName) -Enabled $false 
				}
		}
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	8{
		Write-Host 'Never connected user(s)'
		Get-ADUser -Filter {-not (lastlogontimestamp -like "*") -and -not (iscriticalsystemobject -eq $true)} | Select-Object Name, SamAccountName, LastLogonDate, DistinguishedName, EmployeeId | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\never_connected_users.csv -NoTypeInformation
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\never_connected_users.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	9{
		Write-Host -ForegroundColor Cyan 'All user accounts enabled and not expired'
		Get-ADUser -filter {ObjectClass -like 'user' -and Enabled -eq 'true'} -properties * | where { $_.AccountExpirationDate -gt (Get-Date) -or $_.AccountExpirationDate -eq $null } | Select-Object Name, SamAccountName, LastLogonDate, DistinguishedName, EmailAddress, EmployeeID, Title, Department, AccountExpirationDate | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\all_user_accounts_enabled_and_not_expired.csv -NoTypeInformation
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\all_user_accounts_enabled_and_not_expired.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	10{
		Write-Host -ForegroundColor Cyan 'All enabled user accounts that never exire'
		Get-ADUser -filter {ObjectClass -like 'user' -and Enabled -eq 'true'} -properties * | Where {$_.AccountExpirationDate -eq $null } | Select-Object Name, SamAccountName, LastLogonDate, DistinguishedName, EmployeeId | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\all_user_accounts_never_expire.csv -NoTypeInformation
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\all_user_accounts_never_expire.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	11{
		Write-Host -ForegroundColor Cyan 'All disabled users'
		Search-ADAccount -UsersOnly -AccountDisabled -ResultPageSize 2000 -ResultSetSize $null | Select-Object Name, SamAccountName, LastLogonDate, DistinguishedName, EmployeeId | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\all_disabled_users.csv -NoTypeInformation
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\all_disabled_users.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	12{
		Write-Host -ForegroundColor Cyan 'All expired user accounts and not disabled'
		Search-ADAccount -UsersOnly -AccountExpired -ResultPageSize 2000 -ResultSetSize $null | Select-Object Name, SamAccountName, LastLogonDate, DistinguishedName, AccountExpirationDate, EmployeeId | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\all_expired_user_accounts.csv -NoTypeInformation
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\all_expired_user_accounts.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
		if (( Read-Host 'Do you want to disable expired accounts ? - Administrator rights needed - Answer "sure" to validate' ) -eq 'sure' ){
				Search-ADAccount -UsersOnly -AccountExpired -ResultPageSize 2000 -ResultSetSize $null | Select-Object SamAccountName | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\expired_users_to_disable.csv -NoTypeInformation
				$file_us_dis=import-csv C:\Users\$user\Desktop\expired_users_to_disable.csv 
				foreach ($entry in $file_us_dis) 
				{
					Set-ADUser -Identity $($entry.SamAccountName) -Enabled $false 
				}
		}
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	13{
		Write-Host -ForegroundColor Red 'All users whose password never expire'
		Search-ADAccount -PasswordNeverExpires -UsersOnly -ResultPageSize 2000 -resultSetSize $null | Select-Object Name, SamAccountName, DistinguishedName, EmployeeId | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\all_users_whose_password_never_expire.csv -NoTypeInformation
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\all_users_whose_password_never_expire.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	14{
		Write-Host -ForegroundColor Red 'Users who cannot change their password'
		Get-ADUser -filter * -properties * | where { $_.CannotChangePassword -eq 'true' } | Select-Object Name, SamAccountName, LastLogonDate, DistinguishedName, EmployeeId | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\users_who_cannot_change_their_password.csv -NoTypeInformation
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\users_who_cannot_change_their_password.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	15{
		Write-Host 'Users with mailbox'
		Get-ADUser -filter 'mail -like "*"' -properties * | Select-Object Name, SamAccountName, LastLogonDate, DistinguishedName, EmployeeId, Mail | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\users_with_mailbox.csv -NoTypeInformation
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\users_whith_mailbox.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	16{
		Write-Host 'Users without mailbox'
		Get-ADUser -filter 'mail -notlike "*"' -properties * | Select-Object Name, SamAccountName, LastLogonDate, DistinguishedName, EmployeeId | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\users_without_mailbox.csv -NoTypeInformation
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\users_whithout_mailbox.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	17{
		Write-Host -ForegroundColor Cyan 'Password last set for all users accounts (enabled)'
		Get-ADUser -filter {ObjectClass -like 'user' -and Enabled -eq 'true'} -properties * | Select-Object Name, SamAccountName, PasswordLastSet, DistinguishedName, EmployeeId | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\password_last_set_for_all_users.csv -NoTypeInformation
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\password_last_set_for_all_users.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	18{
		Write-Host -ForegroundColor Cyan 'Password last set for a specific user account'
		Write-Host ''
		[string]$user_name = read-host 'Enter user name (SamAccountName)'
		Get-ADUser -Identity $user_name -properties * | Select-Object Name, SamAccountName, PasswordLastSet, DistinguishedName, EmployeeId | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\password_last_set_for_$user_name.csv -NoTypeInformation
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\password_last_set_for_'$user_name'.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	19{
		Write-Host -ForegroundColor Cyan 'Last logon date for a specific user account'
		Write-Host ''
		[string]$user_name = read-host 'Enter user name (SamAccountName)'
		Get-ADUser -Identity $user_name -properties * | Select-Object Name, SamAccountName, LastLogonDate, DistinguishedName, EmployeeId | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\last_logon_date_for_$user_name.csv -NoTypeInformation
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\last_logon_date_for_'$user_name'.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	20{
		Write-Host -ForegroundColor Cyan 'Search user by EmployeeID'
		Write-Host ''
		[string]$user_id = read-host 'Enter user id (EmployeeID)'
		Get-ADUser -filter {EmployeeID -like $user_id} -Property * | Select Name, SamAccountName, EmployeeId | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\user_id_$user_id.csv -NoTypeInformation
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\user_id_'$user_id'.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	21{
		Write-Host -ForegroundColor Cyan 'Connections history for a specific user'
		Write-Host ''
		$DCs = Get-ADDomainController -Filter *
		[string]$nbod = read-host 'Enter number of days prior in history'
		$startDate = (get-date).AddDays(-$nbod )
		foreach ($DC in $DCs){
		$slogonevents = Get-Eventlog -LogName Security -ComputerName $DC.Hostname -after $startDate | where {$_.eventID -eq 4624 }}
		[string]$user_login= read-host 'Enter user name (SamAccountName)'
		Write-Output "Connections for : "$user_login" from "$startDate"`n" | Out-File -FilePath c:\Users\$user\Desktop\connections_history.txt -Append
		foreach ($e in $slogonevents){
			if (($e.EventID -eq 4624 ) -and ($e.ReplacementStrings[8] -eq 2) -and ($e.ReplacementStrings[5] -eq "$user_login")){
			Write-Output "Type: Local Logon`nDate: "$e.TimeGenerated "Workstation: "$e.ReplacementStrings[11] | Out-File -FilePath c:\Users\$user\Desktop\connections_history.txt -Append
			Write-Host "Type: Local Logon | Date: "$e.TimeGenerated " | Status: Success | User: "$e.ReplacementStrings[5] " | Workstation: "$e.ReplacementStrings[11]
			}
			if (($e.EventID -eq 4624 ) -and ($e.ReplacementStrings[8] -eq 10) -and ($e.ReplacementStrings[5] -eq "$user_login")){
			Write-Output "Type: Remote Logon`nDate: "$e.TimeGenerated "Workstation: "$e.ReplacementStrings[11] "`nIP Address: "$e.ReplacementStrings[18] | Out-File -FilePath c:\Users\$user\Desktop\connections_history.txt -Append
			Write-Host "Type: Remote Logon | Date: "$e.TimeGenerated " | Status: Success | User: "$e.ReplacementStrings[5] " | Workstation: "$e.ReplacementStrings[11] " | IP Address: "$e.ReplacementStrings[18]
			}}
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\connections_history.txt'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	22{
		Write-Host -ForegroundColor Cyan 'Check inheritance ACL for a specific user'
		Write-Host ''
		[string]$user_name = read-host 'Enter user name (SamAccountName)'
		$dn = Get-ADUser -Identity $user_name -properties *
		Get-Acl -Path "AD:\$dn" | select -ExpandProperty Access | Out-File c:\Users\$user\Desktop\inheritance_check_for_$user_name.txt
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\inheritance_check_for_'$user_name'.txt'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
}
}
} while ( $user_extract -ne 23 ) }
<#----------------------------------------------------------------------------------------------------

Group Selection

----------------------------------------------------------------------------------------------------#>
3{
	do {
  [string]$group_extract = 0
  while ( $group_extract -lt 1 -or $group_extract -gt 7) {
Write-Host '                                 ___                           '
Write-Host '        _      _      _         / _ \_ __ ___  _   _ _ __  ___ '
Write-Host '     \/.|.\/\/.|.\/\/.|.\/     / /_\/  __/ _ \| | | |  _ \/ __|'
Write-Host '      \_;_/  \_;_/  \_;_/     / /_\\| | | (_) | |_| | |_) \__ \'
Write-Host '       / \    / \    / \      \____/|_|  \___/ \__,_| .__/|___/'
Write-Host '                                                    |_|        '
	Write-Host ''
	Write-Host '-----------------------------------------------------------------------------------------'
	Write-Host ''
	Write-Host 'Choose group extract type :'
	Write-Host ''
	Write-Host ''
	Write-Host -ForegroundColor Cyan '1 - All groups'
	Write-Host '2 - All objects in a specific group'
	Write-Host '3 - All groups to which an object belongs'
	Write-Host '4 - All objects in "Protected Users" group'
	Write-Host '5 - All objects in "Organization Management" group (Exchange Golbal Administrators)'
	Write-Host '6 - All objects in "Recipient Management" group (Exchange Recipient Administrators)'
	Write-Host ''
	Write-Host -ForegroundColor Yellow '7 - Return'
	Write-Host ''
	Write-Host '-----------------------------------------------------------------------------------------'
	Write-Host ''
	[string]$group_extract = read-host 'Enter your selection'
	Write-Host ''
	Write-Host '-----------------------------------------------------------------------------------------'
	Write-Host ''
	switch ($group_extract) {
	1{
		Write-Host -ForegroundColor Cyan 'All groups'
		Get-ADGroup -Filter * | Select-Object samAccountName, DistinguishedName | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\all_groups.csv -NoTypeInformation
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\all_groups.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	2{
		Write-Host 'All objects in a specific group'
		Write-Host ''
		[string]$group_name = read-host 'Enter group name (SamAccountName)'
		Get-AdGroupMember -Identity $group_name | Select-Object ObjectClass, SamAccountName, Name, DistinguishedName | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\$group_name.csv -NoTypeInformation
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\'$group_name'.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	3{
		Write-Host 'All groups to which an object belongs'
		Write-Host ''
		[string]$object_name = read-host 'Enter object name (SamAccountName)'
		Get-ADPrincipalGroupMembership -Identity $object_name | Select-Object SamAccountName, Name, DistinguishedName | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\$object_name.csv -NoTypeInformation
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\'$object_name'.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	4{
		Write-Host 'All objects in "Protected Users" group'
		Write-Host ''
		Get-AdGroupMember -Identity 'Protected Users' | Select-Object ObjectClass, SamAccountName, Name, DistinguishedName | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\protected_users.csv -NoTypeInformation
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\protected_users.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	5{
		Write-Host 'All objects in "Organization Management" group (Exchange Golbal Administrators)'
		Write-Host ''
		Get-AdGroupMember -Identity 'Organization Management' | Select-Object ObjectClass, SamAccountName, Name, DistinguishedName | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\exchange_administrators.csv -NoTypeInformation
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\exchange_administrators.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
	6{
		Write-Host 'All objects in "Recipient Management" group (Exchange Recipient Administrators)'
		Write-Host ''
		Get-AdGroupMember -Identity 'Recipient Management' | Select-Object ObjectClass, SamAccountName, Name, DistinguishedName | Export-CSV -Encoding UTF8 c:\Users\$user\Desktop\exchange_recipient_administrators.csv -NoTypeInformation
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\exchange_recipient_administrators.csv'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
}
}
} while ( $group_extract -ne 7 ) }
<#----------------------------------------------------------------------------------------------------

Configuration Selection

----------------------------------------------------------------------------------------------------#>
4{
	do {
  [string]$config_extract = 0
  while ( $config_extract -lt 1 -or $config_extract -gt 2) {
Write-Host ' _________________'
Write-Host '|# :           : #|'
Write-Host '|  :           :  |    _________                _____.__ '
Write-Host '|  :           :  |    \_   ___ \  ____   _____/ ____\__| ____'
Write-Host '|  :           :  |    /    \  \/ /  _ \ /    \   __\|  |/ ___\'
Write-Host '|  :___________:  |    \     \___(  <_> )   |  \  |  |  / /_/  >'
Write-Host '|     _________   |     \______  /\____/|___|  /__|  |__\___  /'
Write-Host '|    | __      |  |            \/            \/        /_____/'
Write-Host '|    ||  |     |  |'
Write-Host '\____||__|_____|__|'
	Write-Host ''
	Write-Host '-----------------------------------------------------------------------------------------'
	Write-Host ''
	Write-Host 'Choose group extract type :'
	Write-Host ''
	Write-Host ''
	Write-Host -ForegroundColor Red '1 - Check last modification date of AdminSDHolder object'
	Write-Host ''
	Write-Host -ForegroundColor Yellow '2 - Return'
	Write-Host ''
	Write-Host '-----------------------------------------------------------------------------------------'
	Write-Host ''
	[string]$config_extract = read-host 'Enter your selection'
	Write-Host ''
	Write-Host '-----------------------------------------------------------------------------------------'
	Write-Host ''
	switch ($config_extract) {
	1{
		Write-Host -ForegroundColor Red 'Check last modification date of AdminSDHolder object'
		$adminsdholder_creation = Get-adobject -ldapfilter '(cn=AdminSDHolder)' -Properties * | Select WhenCreated
		$adminsdholder_last_modification = Get-adobject -ldapfilter '(cn=AdminSDHolder)' -Properties * | Select Modified
		Write-Host ''
		Write-Host 'Creation date : '$adminsdholder_creation
		Write-Host ''
		Write-Host 'Last modification : '$adminsdholder_last_modification
		Write-Output '-----------------------------------------------------------------------------------------' | Out-File -FilePath c:\Users\$user\Desktop\adminsdholder_last_modification.txt
		Write-Output 'AdminSDHolder object' | Out-File -FilePath c:\Users\$user\Desktop\adminsdholder_last_modification.txt -Append
		Write-Output '-----------------------------------------------------------------------------------------' | Out-File -FilePath c:\Users\$user\Desktop\adminsdholder_last_modification.txt -Append
		Write-Output 'Creation date : '$adminsdholder_creation | Out-File -FilePath c:\Users\$user\Desktop\adminsdholder_last_modification.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\adminsdholder_last_modification.txt -Append
		Write-Output 'Last modification : '$adminsdholder_last_modification | Out-File -FilePath c:\Users\$user\Desktop\adminsdholder_last_modification.txt -Append
		Write-Output '' | Out-File -FilePath c:\Users\$user\Desktop\adminsdholder_last_modification.txt -Append
		Write-Output 'If the difference between the 2 values is more than 24 hours, it s suspicious. Check AdminSDHolder object permissions' | Out-File -FilePath c:\Users\$user\Desktop\adminsdholder_last_modification.txt -Append
		Write-Host ''
		Write-Host 'Exported file :'
		Write-Host ''
		Write-Host 'C:\Users\'$user'\Desktop\adminsdholder_last_modification.txt'
		Write-Host ''
		Write-Host '-----------------------------------------------------------------------------------------'
		Write-Host ''
	}
}
}
} while ( $config_extract -ne 2 ) }
}
}
} while ( $object_type -ne 5 )
