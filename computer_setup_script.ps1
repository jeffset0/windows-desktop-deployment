# Import variables from file
$jsonVars = Get-Content $PSScriptRoot\computer_setup_variables.json | ConvertFrom-Json

# Iterate through $jsonVars to create new variables
for ($var=0; $var -lt $jsonVars.variables.count; $var++) {
    New-Variable -Name $jsonVars.variables.key[$var] -Value $jsonVars.variables.value[$var]
}

# Get new computer name
$newComputerName = Read-Host -Prompt 'Enter new computer or laptop name'

# Get admin credential
$adminCredential = Get-Credential "$domain\" -Message 'Enter domain credentials (domain\username)'

# Check if computer or laptop name exists
Try {
        Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" # Install and import Active Directory module
        $exist = Get-ADComputer $newComputerName -ErrorAction Stop -Credential $adminCredential # Check AD if computer name exist
        Write-Host $newComputerName "exist or credentials incorrect"
        Start-Sleep -S 10 # Wait for 10 seconds
    }
    Catch {
        # If computer does not already exist

        # Add drive to use in Powershell
        New-PSDrive -Name V -PSProvider filesystem -Root $folder_path -Credential $adminCredential

        # Change power settings for desktops and laptops
        powercfg -x -monitor-timeout-ac 0 # display setting plugged in
        powercfg -x -monitor-timeout-dc 0 # display setting on battery

        powercfg -x -disk-timeout-ac 0 # hard disk setting plugged in
        powercfg -x -disk-timeout-dc 0 # hard disk setting on battery

        powercfg -x -standby-timeout-ac 0 # sleep setting plugged in
        powercfg -x -standby-timeout-dc 0 # sleep setting on battery

        powercfg -setacvalueindex SCHEME_CURRENT 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 0 # disable usb selective suspend plugged in
        powercfg -setdcvalueindex SCHEME_CURRENT 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 0 # disable usb selective suspend on battery

        powercfg -setacvalueindex SCHEME_CURRENT 4f971e89-eebd-4455-a8de-9e59040e7347 7648efa3-dd9c-4e3e-b566-50f929386280 0 # power button do nothing plugged in
        powercfg -setdcvalueindex SCHEME_CURRENT 4f971e89-eebd-4455-a8de-9e59040e7347 7648efa3-dd9c-4e3e-b566-50f929386280 0 # power button do nothing on battery

        powercfg -setacvalueindex SCHEME_CURRENT 4f971e89-eebd-4455-a8de-9e59040e7347 96996bc0-ad50-47ec-923b-6f41874dd9eb 0 # sleep button do nothing plugged in
        powercfg -setdcvalueindex SCHEME_CURRENT 4f971e89-eebd-4455-a8de-9e59040e7347 96996bc0-ad50-47ec-923b-6f41874dd9eb 0 # sleep button do nothing on battery

        # Change power settings for laptops
        powercfg -x -hibernate-timeout-ac 0 # hibernate setting plugged in
        powercfg -x -hibernate-timeout-dc 0 # hibernate setting on battery

        powercfg -hibernate off # turn off all hibernate options including hybrid sleep

        powercfg -setacvalueindex SCHEME_CURRENT 19cbb8fa-5279-450e-9fac-8a3d5fedd0c1 12bbebe6-58d6-4636-95bb-3217ef867c1a 0 # wireless maximum performance plugged in
        powercfg -setdcvalueindex SCHEME_CURRENT 19cbb8fa-5279-450e-9fac-8a3d5fedd0c1 12bbebe6-58d6-4636-95bb-3217ef867c1a 0 # wireless maximum performance on battery

        powercfg -setacvalueindex SCHEME_CURRENT 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936 0 # close lid do nothing plugged in
        powercfg -setdcvalueindex SCHEME_CURRENT 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936 0 # close lid do nothing on battery

        # Apply power setting changes
        powercfg -SetActive SCHEME_CURRENT

        # Install chocolatey
        Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))

        ### use for loop
        # Install Zoom
        Write-Host '=========='
        Write-Host '=========='
        Write-Host '==========          Installing Zoom'
        Write-Host '=========='
        Write-Host '=========='
        choco install -y zoom --install-args="ZSSOHOST=$zoomSSO"

        # Install Adobe Acorbat Reader
        Write-Host '=========='
        Write-Host '=========='
        Write-Host '==========          Installing Acrobat Reader'
        Write-Host '=========='
        Write-Host '=========='
        choco install -y adobereader

        # Install CutePDF Writer
        Write-Host '=========='
        Write-Host '=========='
        Write-Host '==========          Installing CutePDF Writer'
        Write-Host '=========='
        Write-Host '=========='
        choco install -y cutepdf

        # Install Google Chrome
        Write-Host '=========='
        Write-Host '=========='
        Write-Host '==========          Installing Chrome'
        Write-Host '=========='
        Write-Host '=========='
        choco install -y googlechrome

        # Install Mozilla Firefox
        Write-Host '=========='
        Write-Host '=========='
        Write-Host '==========          Installing Firefox'
        Write-Host '=========='
        Write-Host '=========='
        choco install -y firefox

        # Install IrFanView
        Write-Host '=========='
        Write-Host '=========='
        Write-Host '==========          Installing IrFanView'
        Write-Host '=========='
        Write-Host '=========='
        choco install -y irfanview

        # Install 7-Zip
        Write-Host '=========='
        Write-Host '=========='
        Write-Host '==========          Installing 7-Zip'
        Write-Host '=========='
        Write-Host '=========='
        choco install -y 7zip

        # Install Java Runtime
        Write-Host '=========='
        Write-Host '=========='
        Write-Host '==========          Installing Java Runtime'
        Write-Host '=========='
        Write-Host '=========='
        choco install -y javaruntime

        # Install Microsoft .Net 4.6.2
        Write-Host '=========='
        Write-Host '=========='
        Write-Host '==========          Installing .Net 4.6.2'
        Write-Host '=========='
        Write-Host '=========='
        choco install -y dotnet-4.6.2

        # Install Microsoft Silverlight
        Write-Host '=========='
        Write-Host '=========='
        Write-Host '==========          Installing Silverlight'
        Write-Host '=========='
        Write-Host '=========='
        choco install -y silverlight

        # Install Image Resizer for Windows
        Write-Host '=========='
        Write-Host '=========='
        Write-Host '==========          Installing Image Resizer'
        Write-Host '=========='
        Write-Host '=========='
        choco install -y imageresizerapp

        # Install Kaspersky Endpoint Security
        Write-Host '=========='
        Write-Host '=========='
        Write-Host '==========          Installing Kaspersky Endpoint Security'
        Write-Host '=========='
        Write-Host '=========='
        Copy-Item "V:\kaspersky" -Destination "$PSScriptRoot\" -Recurse # Copy installer folder to temp folder
        Start-Process -Wait -FilePath msiexec.exe -ArgumentList "/i $PSScriptRoot\kaspersky\kes_win.msi EULA=1 KSN=1 PRIVACYPOLICY=1 ALLOWREBOOT=0 /qn" # Install in silent mode

        # Install Kaspersky Network Agent
        Write-Host '=========='
        Write-Host '=========='
        Write-Host '==========          Installing Kaspersky Network Agent'
        Write-Host '=========='
        Write-Host '=========='
        Start-Process -Wait -FilePath msiexec.exe -ArgumentList "/i $PSScriptRoot\kaspersky\Kaspersky_Network_Agent.msi SERVERADDRESS=$kasp_ip DONT_USE_ANSWER_FILE=1 EULA=1 PRIVACYPOLICY=1 /qn" # Install in silent mode

        # Install PrintKey Pro
        Write-Host '=========='
        Write-Host '=========='
        Write-Host '==========          Installing PrintKey Pro'
        Write-Host '=========='
        Write-Host '=========='
        Copy-Item "V:\PrintKey-Pro.exe" -Destination "$PSScriptRoot\PrintKey-Pro.exe" # Copy installer to temp folder
        Start-Process -Wait -FilePath $PSScriptRoot\PrintKey-Pro.exe -ArgumentList "/S /v/qn" # Install in silent mode

        # Rename PrintKey license file and create new file with valid key
        if ("C:\Program Files (x86)\Warecentral\PrintKey-Pro\PKey_Pro.ini") {
            Rename-Item -Path "C:\Program Files (x86)\Warecentral\PrintKey-Pro\PKey_Pro.ini" -NewName "PKey_Pro_Orig.ini"
        }
        for ($var=0; $var -lt $jsonVars.printkey.count; $var++) {
            $jsonVars.printkey.value[$var] | Out-File -FilePath "C:\Program Files (x86)\Warecentral\PrintKey-Pro\PKey_Pro.ini" -Append
        }

        # Install Microsoft Office 365
        Write-Host '=========='
        Write-Host '=========='
        Write-Host '==========          Installing Office 365'
        Write-Host '=========='
        Write-Host '=========='
        Copy-Item "V:\O365ProPlusNew.xml" -Destination "$PSScriptRoot\O365ProPlusNew.xml" # Copy xml to temp folder
        choco install -y office365proplus --params "/ConfigPath:$PSScriptRoot\O365ProPlusNew.xml"

        # Install TeamViewer
        Write-Host '=========='
        Write-Host '=========='
        Write-Host '==========          Installing TeamViewer'
        Write-Host '=========='
        Write-Host '=========='
        Copy-Item "V:\TeamViewer_Host_Setup.exe" -Destination "$PSScriptRoot\TeamViewer_Host_Setup.exe" # Copy installer to temp folder
        Start-Process -Wait -FilePath $PSScriptRoot\TeamViewer_Host_Setup.exe -ArgumentList "/S" # Install in silent modes

        # Install Mimecast for Outlook
        Write-Host '=========='
        Write-Host '=========='
        Write-Host '==========          Installing Mimecast for Outlook'
        Write-Host '=========='
        Write-Host '=========='
        Copy-Item "V:\Mimecast_for_Outlook.msi" -Destination "$PSScriptRoot\Mimecast_for_Outlook.msi" # Copy installer to temp folder
        Start-Process -Wait -FilePath msiexec.exe -ArgumentList "/i $PSScriptRoot\Mimecast_for_Outlook.msi /qn /norestart" # Install in silent mode

        # Check if workstation is a laptop
        if(Get-WmiObject -Class win32_systemenclosure | Where-Object {$_.chassistypes -eq 9 -or $_.chassistypes -eq 10 -or $_.chassistypes -eq 14}) # Check if laptop | 9=Laptop, 10=Notebook, 14=Sub-Notebook
            {
                # Install DisplayLink
                Write-Host '=========='
                Write-Host '=========='
                Write-Host '==========          Installing DisplayLink'
                Write-Host '=========='
                Write-Host '=========='
                choco install -y displaylink

                # Create VPN connection
                Write-Host '=========='
                Write-Host '=========='
                Write-Host '==========          Creating VPN profile and desktop shortcut'
                Write-Host '=========='
                Write-Host '=========='
                Add-VpnConnection -Name $vpn_name -AuthenticationMethod Pap -EncryptionLevel Optional -AllUserConnection -IdleDisconnectSeconds 0 -DnsSuffix $domain -ServerAddress $vpn_address -TunnelType L2tp -L2tpPsk $vpn_key -RememberCredential -Force # Create VPN profile
                Copy-Item "V:\$vpn_name.lnk" -Destination "C:\Users\Public\Desktop\$vpn_name.lnk" # Copy VPN shortcut to desktop
            }

        # Enable Remote Desktop
        Write-Host '=========='
        Write-Host '=========='
        Write-Host '==========          Enable Remote Desktop'
        Write-Host '=========='
        Write-Host '=========='
        Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -Name "fDenyTSConnections" -Value 0

        # Disable Windows Firewall
        Write-Host '=========='
        Write-Host '=========='
        Write-Host '==========          Disable Windows Firewall'
        Write-Host '=========='
        Write-Host '=========='
        Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled False

        # Activate Windows
        Write-Host '=========='
        Write-Host '=========='
        Write-Host '==========          Activating Windows'
        Write-Host '=========='
        Write-Host '=========='
        Start-Process -FilePath "c:\Windows\System32\changePK.exe" -ArgumentList "/ProductKey $windows_key"

        # Get and install windows update
        Write-Host '=========='
        Write-Host '=========='
        Write-Host '==========          Installing Windows Updates'
        Write-Host '=========='
        Write-Host '=========='
        $NuGet = Get-PackageProvider | Where-Object {$_.Name -eq "NuGet"}
        If ($NuGet = " ") {Install-PackageProvider -Name NuGet -Force} # Install NuGet package if missing

        If(-not(Get-InstalledModule PSWindowsUpdate -ErrorAction silentlycontinue)) # Check if PSWIndowsUpdate module is installed
            {
                Set-PSRepository PSGallery -InstallationPolicy Trusted # Trust PSGallery Repository
                Install-Module PSWindowsUpdate -Confirm:$False -Force # Install PSWindowsUpdate module
            }
        Get-WindowsUpdate -Install -AcceptAll -IgnoreReboot

        # Get and install windows update
        Write-Host '=========='
        Write-Host '=========='
        Write-Host '==========          Enable SMB1'
        Write-Host '=========='
        Write-Host '=========='
        # Get-WindowsOptionalFeature -Online -FeatureName "SMB1Protocol"
        Enable-WindowsOptionalFeature -Online -FeatureName "SMB1Protocol" -All -NoRestart

        # Rename computer and join domain
        Write-Host '=========='
        Write-Host '=========='
        Write-Host '==========          Renaming computer to' $newComputerName 'and joining domain' $domain
        Write-Host '=========='
        Write-Host '=========='

        Start-Sleep -S 3 # Wait for 3 seconds

        Add-Computer -DomainName $domain -NewName $newComputerName -Credential $adminCredential -Restart -Force

        Write-Host '=========='
        Write-Host '=========='
        Write-Host '==========          Restarting computer'
        Write-Host '=========='
        Write-Host '=========='

        Start-Sleep -S 3 # Wait for 3 seconds before restart
    }