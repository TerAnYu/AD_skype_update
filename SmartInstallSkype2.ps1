$ip = Get-WmiObject -Class Win32_NetworkAdapterConfiguration | foreach{If($_.defaultipgateway){$_.defaultipgateway}}
# $fname = "skype.exe"
# $args = "/VERYSILENT /NOLAUNCH /NOGOOGLE /NOSTARTUP /NOPLUGINS /LANG=ru"
$fname = "skype.msi"
$args = "/quiet /norestart"

# откуда берём установщик
	if ($ip -eq "10.1.48.1") {$source = "\\share1\share2\distrib\skype\"+$fname}
	if ($ip -eq "10.1.76.1") {$source = "\\share1\share2\distrib\skype\"+$fname}
	if ($ip -eq "10.1.47.1") {$source = "\\share2\distrib\skype\"+$fname}
	if ($ip -eq "192.168.4") {$source = "\\share2\distrib\skype\"+$fname}
	if ($ip -eq "10.1.45.1") {$source = "\\share3\Distr\skype\"+$fname}
	if ($ip -eq "10.1.46.1") {$source = "\\share4\distrib\skype\"+$fname}
	if ($ip -eq "10.1.44.1") {$source = "\\share5\distrib\skype\"+$fname}
	if ($ip -eq "10.1.49.1") {$source = "\\share6\share\!Distr\skype\"+$fname}
Write-Host -NoNewline "Устанавливаю отсюда:" $source `n

$fpath = "C:\Program Files (x86)\Microsoft\Skype for Desktop\Skype.exe"
$fpaths = "C:\Program Files\Microsoft\Skype for Desktop\Skype.exe"
$isfile = Test-Path $fpath
$isfiles = Test-Path $fpaths
$yearlog=$(Get-Date).ToUniversalTime().ToString("yyyy")
$logfname = $yearlog+"_log.txt"

# Здесь происходит магия, функция по определению номера версий
function Get-EXEFileVersion {
    param (
        [IO.FileInfo] $FLE
    )
		try {
			$version = (get-item $FLE).VersionInfo.FileVersion
			$version = $version.Trim()
			return $version
		} catch {
			throw "Failed to get file version: {0}." -f $_
		}
}

function Get-MsiDatabaseVersion {
    param (
        [IO.FileInfo] $FLE
    )
		try {
			$windowsInstaller = New-Object -com WindowsInstaller.Installer
			$database = $windowsInstaller.GetType().InvokeMember(
				"OpenDatabase", "InvokeMethod", $Null,
				$windowsInstaller, @($FLE.FullName, 0)
			)
		 
			$q = "SELECT Value FROM Property WHERE Property = 'ProductVersion'"
			$View = $database.GetType().InvokeMember(
				"OpenView", "InvokeMethod", $Null, $database, ($q)
			)
		 
			$View.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $View, $Null)
			$record = $View.GetType().InvokeMember( "Fetch", "InvokeMethod", $Null, $View, $Null )
			$version = $record.GetType().InvokeMember( "StringData", "GetProperty", $Null, $record, 1 )
			$version = $version.Trim()
			return $version
		} catch {
			throw "Failed to get file version: {0}." -f $_
		}
}

function Get-FileVersionRes {
    param (
        [IO.FileInfo] $FLE
    )
		try {
			if ([System.IO.Path]::GetExtension($source).Split(".")[1] -eq "msi") {
				$skpinst = (Get-MsiDatabaseVersion $source)[1]
			}
			else {
				$skpinst = (Get-EXEFileVersion $source)
			}
			return $skpinst
		} catch {
			throw "Failed to get file version: {0}." -f $_
		}
}

# Здесь магия сравнений установленной версии и новой, принимаем решение обновляем, устанавливаем или ничего не делаем
if ($isfile -eq "True") {
	Write-Host -NoNewline "Skype установлен x86!" `n
	$skp = (Get-EXEFileVersion $fpath)
	$skpinst = (Get-FileVersionRes $source)
	Write-Host -NoNewline "Установленная версия:" $skp `n
	Write-Host -NoNewline "Устанавливаемая версия:" $skpinst `n
	if ($skp -ne $skpinst) {
		Write-Host -NoNewline "Обновляю с $skp до $skpinst" `n
		$app = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -match "Skype" }
		$app.Uninstall()
		$exitcode=(Start-Process $source -ArgumentList "$args" -Wait -Passthru).ExitCode
		Add-Content -Encoding UTF8 -Path "$(Split-Path -Path $source)\$logfname" -Value "$(Get-Date -Format "yyyy\/MM\/dd HH:mm:ss K") - $env:computername - Обновляю с $skp до $skpinst - код завершения: $exitcode"
	}
	else {
	Write-Host -NoNewline "Обновление не требуется" `n
	Add-Content -Encoding UTF8 -Path "$(Split-Path -Path $source)\$logfname" -Value "$(Get-Date -Format "yyyy\/MM\/dd HH:mm:ss K") - $env:computername - Обновление не требуется: $skp"
	}
}

if ($isfiles -eq "True") {
	Write-Host -NoNewline "Skype установлен x64!" `n
	$skp = (Get-EXEFileVersion $fpaths)
	$skpinst = (Get-FileVersionRes $source)
	Write-Host -NoNewline "Установленная версия: " $skp `n
	Write-Host -NoNewline "Устанавливаемая версия:" $skpinst `n
	if ($skp -ne $skpinst) {
		Write-Host -NoNewline "Обновляю с $skp до $skpinst" `n
		$app = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -match "Skype" }
		$app.Uninstall()
		$exitcode=(Start-Process $source -ArgumentList "$args" -Wait -Passthru).ExitCode
		Add-Content -Encoding UTF8 -Path "$(Split-Path -Path $source)\$logfname" -Value "$(Get-Date -Format "yyyy\/MM\/dd HH:mm:ss K") - $env:computername - Обновляю с $skp до $skpinst - код завершения: $exitcode"
	}
	else {
	Write-Host -NoNewline "Обновление не требуется" `n
	Add-Content -Encoding UTF8 -Path "$(Split-Path -Path $source)\$logfname" -Value "$(Get-Date -Format "yyyy\/MM\/dd HH:mm:ss K") - $env:computername - Обновление не требуется: $skp"
	}
}

if ($isfile -ne "True") { if ($isfiles -ne "True"){
		$skpinst = (Get-FileVersionRes $source)
		Write-Host -NoNewline "Skype не найден! Устанавливаю $skpinst!" `n
		$exitcode=(Start-Process $source -ArgumentList "$args" -Wait -Passthru).ExitCode
		Add-Content -Encoding UTF8 -Path "$(Split-Path -Path $source)\$logfname" -Value "$(Get-Date -Format "yyyy\/MM\/dd HH:mm:ss K") - $env:computername - не было, усатнавливаю: $skpinst - код завершения: $exitcode"
	}
}
