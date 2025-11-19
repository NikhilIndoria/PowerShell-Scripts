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

Execute-Process -Path "$dirFiles\DaxStudio_3_3_2_setup.exe" -Parameters '/ALLUSERS /VERYSILENT /SUPPRESSMSGBOXES' -WindowStyle Hidden

##*===============================================
##* POST-INSTALLATION
##*===============================================
[String]$installPhase = 'Post-Installation'

## <Perform Post-Installation tasks here>


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

# Execute-Process -path "$env:ProgramFiles\Bosch\ConfigManager\uninst_VL_ConfigManager.exe" -Parameters '/S /noreboot' -WindowStyle Hidden -NoWait

Execute-Process -Path "$env:ProgramFiles\DAX Studio\unins000.exe" -Parameters '/ALLUSERS /VERYSILENT /SUPPRESSMSGBOXES'

##*===============================================
##* POST-UNINSTALLATION
##*===============================================
[String]$installPhase = 'Post-Uninstallation'

## <Perform Post-Uninstallation tasks here>
}
