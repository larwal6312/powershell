###Script to uninstall old adobe reader package and install an update version"

# List of servers/computers
$Computers = get-content "D:\Scripts\adobe_upgrade\serverlist.txt"

# Path to updated adobe executable
$adobeEXE = "\\network_location\adobeReader\AcroRdrDC2001320064_en_US.exe"

# Loop through servers/computers
foreach ($Computer in $Computers)
{
    Write-Host "`r`nTesting connection to $Computer"
    if (Test-Connection -ComputerName "$Computer" -Quiet -Count 4)
        {
        Write-Host "I can ping $Computer"
        if (Test-Path "\\$Computer\c$")
            {
            Write-Host "I have access to $Computer"
            $arrAdobe = Get-WMIObject win32_product -ComputerName $Computer | where { $_.name -eq "Adobe Acrobat Reader DC" }

            #only uninstall if version is older than version installing
            if (($arrAdobe.Name -eq "Adobe Acrobat Reader DC") -and ([version]$arrAdobe.Version -eq [version]"20.013.20064"))
                {
                Write-Host "Uninstalling adobe from $Computer"
                Invoke-Command -ComputerName $Computer -ScriptBlock {
                    $Adobe = Get-WMIObject win32_product | where { $_.name -eq "Adobe Acrobat Reader DC" }
                    $Adobe.Uninstall()
                    }
                write-host "Adobe was uninstalled"
                Write-Host "Copying executable to $Computer"
                Copy-Item $adobeEXE "\\$Computer\d$\"

                Write-Host "Installing new version of Adobe"
                Invoke-Command -ComputerName $Computer -ScriptBlock {
                    Start-Process -FilePath "D:\AcroRdrDC2001320064_en_US.exe" -ArgumentList "/sPB /rs" -Wait
                    }
                }
            else
                {
                Write-Host "Did not uninstall adobe from $Computer"
                Write-Host $arrAdobe.Name "version" $arrAdobe.Version "found"
                }
            }
        else
            {
            Write-Host "Cannot access $Computer"
            }
        }
    else
        {
        Write-Host "$Computer did not ping"
        }
}
