##*===============================================
##* INSTALLATION
##*===============================================
[String]$installPhase = 'Installation'

## Handle Zero-Config MSI Installations
If ($useDefaultMsi) {
    [Hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Install'; Path = $defaultMsiFile }; If ($defaultMstFile) {
        $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile)
    }
    Execute-MSI @ExecuteDefaultMSISplat; If ($defaultMspFiles) {
        $defaultMspFiles | ForEach-Object { Execute-MSI -Action 'Patch' -Path $_ }
    }
}

## <Perform Installation tasks here>

Execute-MSI -Path "$dirFiles\STIG Viewer 3.msi" -Parameters '/qn'
Remove-Item -Path "C:\Users\Public\Desktop\STIG Viewer 3.lnk" -Force
##*===============================================
##* POST-INSTALLATION
##*===============================================
[String]$installPhase = 'Post-Installation'

## <Perform Post-Installation tasks here>

Remove-Item -Path "$env:Public\Desktop\Start ManageEngine Endpoint Central.lnk" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\ManageEngine Endpoint Central\Help.lnk" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\ManageEngine Endpoint Central\ReadMe.lnk" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\ManageEngine Endpoint Central\Update Manager.lnk" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\ManageEngine Endpoint Central\Uninstall.lnk" -Force -ErrorAction SilentlyContinue


}
ElseIf ($deploymentType -ieq 'Uninstall') {
##*===============================================
##* PRE-UNINSTALLATION
##*===============================================
[String]$installPhase = 'Pre-Uninstallation'

## Show Welcome Message, close Internet Explorer with a 60 second countdown before automatically closing


## Show Progress Message (with the default message)


## <Perform Pre-Uninstallation tasks here>


##*===============================================
##* UNINSTALLATION
##*===============================================
[String]$installPhase = 'Uninstallation'

## Handle Zero-Config MSI Uninstallations
If ($useDefaultMsi) {
    [Hashtable]$ExecuteDefaultMSISplat = @{ Action = 'Uninstall'; Path = $defaultMsiFile }; If ($defaultMstFile) {
        $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile)
    }
    Execute-MSI @ExecuteDefaultMSISplat
}

## <Perform Uninstallation tasks here>


Execute-MSI -Action Uninstall -Path "{E09F8DE0-1572-427F-889E-3BE6ED80621E}" -Parameters '/qn'


##*===============================================
##* POST-UNINSTALLATION
##*===============================================
[String]$installPhase = 'Post-Uninstallation'

## <Perform Post-Uninstallation tasks here>
}
