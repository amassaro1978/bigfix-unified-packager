@{
    SchemaVersion = '1.0'
    Name = 'REPLACE WITH APPLICATION NAME'
    RecipeVersion = '1.0'
    Vendor = 'REPLACE WITH VENDOR'
    Description = 'REPLACE WITH A SHORT DESCRIPTION OF THE CUSTOM PACKAGE STEPS'

    Sections = @{
        # Runs after the Unified Packager's automatic pre-install process kills.
        PreInstall = @'
'@

        # When nonblank, replaces the automatically generated MSI/EXE install command.
        Install = @'
'@

        # Runs after the Unified Packager's automatic post-install shortcut cleanup.
        PostInstall = @'
'@

        # Runs after the Unified Packager's automatic pre-uninstall process kills.
        PreUninstall = @'
'@

        # When nonblank, replaces the automatically generated uninstall command.
        Uninstall = @'
'@

        # Runs at the end of Post-Uninstallation.
        PostUninstall = @'
'@
    }
}
