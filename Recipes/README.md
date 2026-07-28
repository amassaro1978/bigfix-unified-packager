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

`RecipeVersion`, `Vendor`, and `Description` are optional metadata.

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
