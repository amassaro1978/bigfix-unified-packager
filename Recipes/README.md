# PSADT Recipe Library

Set `$RecipeSharePath` near the top of `Invoke-BigFixPackager.ps1` to this
folder or to a trusted UNC share. The Unified Packager loads `*.psd1` recipe
files from that location.

Start a new recipe by copying
`Templates\Recipe-Template.psd1` into the recipe-library root, renaming it for
the application, and filling in only the sections the package needs. The
`Templates` subfolder is not scanned by the dropdown.

## Required fields

- `SchemaVersion` must be `1.0`.
- `Name` is the friendly name shown in the dropdown.
- `Sections` must be a hashtable.

`RecipeVersion`, `Vendor`, `Description`, `FixletIconPath`,
`ProcessesToKill`, and `DesktopShortcutName` are optional metadata.

Selecting a recipe copies its `Vendor` and `Name` values into the Unified
Packager's Vendor and Application Name fields.

When `FixletIconPath` is set, selecting the recipe automatically selects and
previews that image in the Unified Packager's Fixlet Icon field. Use a trusted
PNG, JPG, JPEG, or ICO file. UNC and absolute paths are used as written;
relative paths are resolved from the recipe file's folder.

When `ProcessesToKill` is set to an array of process executable names,
selecting the recipe populates the Unified Packager's comma-separated
Processes to Kill field. Enter names without the `.exe` extension.

When `DesktopShortcutName` is set, selecting the recipe populates the Desktop
Shortcut Name field. Enter the shortcut filename without the `.lnk` extension.

Switching recipes clears a value only when it came from the previous recipe.
Values entered or changed manually remain available when the selected recipe
does not supply that metadata.

## Supported sections

- `PreInstall`
- `Install`
- `PostInstall`
- `PreUninstall`
- `Uninstall`
- `PostUninstall`

Blank or omitted sections are ignored. `Install` and `Uninstall` replace the
basic auto-generated installer commands when nonblank. Pre/Post sections run
after the Unified Packager's built-in actions.

## Supported tokens

- `{{Vendor}}`
- `{{AppName}}`
- `{{AppVersion}}`
- `{{InstallerType}}`
- `{{InstallerFile}}`
- `{{FilesDirectory}}` (expands to `$adtSession.DirFiles`)

Unknown tokens and invalid PowerShell block script generation. A copy of the
selected recipe is saved under `SupportFiles\RecipeSnapshots` in the package.
