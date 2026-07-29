@{
    SchemaVersion = '1.0'
    Name = 'Example Application'
    RecipeVersion = '1.0'
    Vendor = 'Example Vendor'
    Description = 'Example showing all supported PSADT recipe sections and tokens.'

    ProcessesToKill = @('ExampleApp', 'ExampleUpdater')
    DesktopShortcutName = 'Example Application'

    Sections = @{
        PreInstall = @'
# Runs after the Unified Packager's automatic process-kill block.
Write-ADTLogEntry -Message 'Running custom pre-install tasks.' -Source $adtSession.InstallPhase
'@

        Install = @'
# A nonblank Install section replaces the automatically generated MSI/EXE command.
$installerPath = Join-Path {{FilesDirectory}} '{{InstallerFile}}'
Write-ADTLogEntry -Message "Installing $installerPath" -Source $adtSession.InstallPhase
Start-ADTProcess -FilePath $installerPath -ArgumentList '/S' -WaitForMsiExec
'@

        PostInstall = @'
# Runs after automatic desktop shortcut removal.
Write-ADTLogEntry -Message 'Running custom post-install tasks.' -Source $adtSession.InstallPhase
'@

        PreUninstall = @'
# Runs after the Unified Packager's automatic process-kill block.
Write-ADTLogEntry -Message 'Running custom pre-uninstall tasks.' -Source $adtSession.InstallPhase
'@

        Uninstall = @'
# A nonblank Uninstall section replaces the automatically generated uninstall command.
Write-ADTLogEntry -Message 'Running custom uninstall tasks.' -Source $adtSession.InstallPhase
'@

        PostUninstall = @'
Write-ADTLogEntry -Message 'Running custom post-uninstall cleanup.' -Source $adtSession.InstallPhase
'@
    }
}
