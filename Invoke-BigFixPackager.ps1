<#
.SYNOPSIS
    Unified BigFix Packaging Tool - PSADT 4.x to Fixlets + QA Offers in one flow.
.DESCRIPTION
    Takes a PSADT package folder, generates Invoke-AppDeployToolkit.ps1,
    opens ISE for review, signs the script, creates Install/Update/Remove fixlets,
    POSTs them to BigFix, then creates QA offers automatically.
.AUTHOR
    Anthony Massaro
.VERSION
    1.0
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

$ToolVer = "1.0 - Unified Packager"

# =========================
# CONFIG (customize per environment)
# =========================
$LogFile = "C:\temp\BigFixPackager.log"

# PSADT recipe library. Replace this with a trusted UNC path, for example:
# $RecipeSharePath = "\\server\software\PSADT-Recipes"
$RecipeSharePath = Join-Path $PSScriptRoot "Recipes"

# Code Signing - update with your cert thumbprint
$SigningCertThumbprint = "YOUR-CERT-THUMBPRINT-HERE"

# BigFix servers
$ServerList = @(
    "https://dev.server:52311",
    "https://prod.server:52311"
)

# Sites
$SiteList = @(
    "Test Group Managed (Workstations)",
    "test site 1"
)

# Fixlet Action name
$FixletActionName_Default = "Action1"

# Offer group IDs
$QA_GroupIdWithPrefix = "00-12345"

# Group targeting
$UseDirectGroupMembershipRelevance = $true
$UseSitesPlural = $true

# Offer defaults
$OfferDefaults = @{
    PreActionShowUI  = $false
    RetryCount       = 3
    RetryWaitISO     = "PT1H"
    StartOffsetISO   = "PT0S"
    EndOffsetISO     = "P365DT0S"
    Reapply          = $true
    ContinueOnErr    = $true
}

# LLM Config (for offer descriptions)
$LLMConfig = @{
    EnableAuto = $true
    ApiUrl     = "https://redacted/v1/chat/completions"
    ApiKeyEnv  = "LITELLM_KEY"
    Model      = "gpt-40"
}

# HTTP settings
$IgnoreCertErrors = $true
$SaveXmlToTemp    = $true

# =========================
# UTILITY FUNCTIONS
# =========================
function LogLine($txt) {
    try {
        $line = "{0} {1}" -f (Get-Date -Format 'u'), $txt
        if ($script:LogBox) {
            $script:LogBox.AppendText($line + "`r`n")
            $script:LogBox.SelectionStart = $script:LogBox.Text.Length
            $script:LogBox.ScrollToCaret()
        }
        $dir = Split-Path $LogFile -Parent
        if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        # Use FileStream with shared read/write to avoid conflicts with CMTrace
        try {
            $fs = [System.IO.FileStream]::new($LogFile, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
            $sw = [System.IO.StreamWriter]::new($fs)
            $sw.WriteLine($line)
            $sw.Close()
            $fs.Close()
        } catch {
            # Silently skip file write if still locked
        }
    } catch {}
}

function Encode-SiteName([string]$Name) {
    $enc = [System.Web.HttpUtility]::UrlEncode($Name, [System.Text.Encoding]::UTF8)
    $enc = $enc -replace '\+','%20' -replace '\(','%28' -replace '\)','%29'
    return $enc
}

function Get-BaseUrl([string]$ServerInput) {
    if (-not $ServerInput) { throw "Server is empty." }
    $s = $ServerInput.Trim()
    if ($s -notmatch '^(?i)https?://') {
        if ($s -match ':\d+$') { $s = "https://$s" } else { $s = "https://$s`:52311" }
    }
    return $s.TrimEnd('/')
}

function Join-ApiUrl([string]$BaseUrl,[string]$RelativePath) {
    $rp = if ($RelativePath.StartsWith("/")) { $RelativePath } else { "/$RelativePath" }
    $BaseUrl.TrimEnd('/') + $rp
}

function Get-AuthHeader([string]$User,[string]$Pass) {
    $pair = "$User`:$Pass"
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    "Basic " + [Convert]::ToBase64String($bytes)
}

function SafeEscape([string]$s) {
    if ($null -eq $s) { return "" }
    [System.Security.SecurityElement]::Escape($s)
}

function To-XmlBool([bool]$b) { if ($b) { 'true' } else { 'false' } }

function Get-NumericGroupId([string]$GroupIdWithPrefix) {
    if ($GroupIdWithPrefix -match '^\d{2}-(\d+)$') { return $Matches[1] }
    return ($GroupIdWithPrefix -replace '[^\d]','')
}

function Build-GroupMembershipRelevance([string]$SiteName,[string]$GroupIdNumeric,[bool]$UseSitesPluralLocal=$UseSitesPlural) {
    if ($UseSitesPluralLocal) {
        return "(member of group $GroupIdNumeric of sites)"
    } else {
        $siteEsc = $SiteName.Replace('"','\"')
        return "(member of group $GroupIdNumeric of site whose (name of it = `"$siteEsc`"))"
    }
}

# =========================
# ICON HELPERS
# =========================
function Find-IconFiles([string]$FolderPath) {
    $filesFolder = Join-Path $FolderPath "Files"
    $icons = @()
    # Check Files/ subfolder first (most likely location for app icon)
    if (Test-Path $filesFolder) {
        $icons = @(Get-ChildItem -Path $filesFolder -File | Where-Object {
            $_.Extension -match '^\.(png|jpg|jpeg|ico)$' -and
            $_.Name -notmatch '^(AppDeployToolkit|PSADT|Setup|Deploy-Application)'
        })
    }
    # Also check root folder, excluding PSADT framework icons
    $icons += @(Get-ChildItem -Path $FolderPath -File | Where-Object {
        $_.Extension -match '^\.(png|jpg|jpeg|ico)$' -and
        $_.Name -notmatch '^(AppDeployToolkit|PSADT|Setup|Deploy-Application)'
    })
    return $icons
}

function Get-IconBase64([string]$IconPath) {
    if (-not $IconPath -or -not (Test-Path $IconPath)) { return $null }
    $bytes = [System.IO.File]::ReadAllBytes($IconPath)
    $ext = [System.IO.Path]::GetExtension($IconPath).ToLower()
    $mime = switch ($ext) {
        '.png'  { 'image/png' }
        '.jpg'  { 'image/jpeg' }
        '.jpeg' { 'image/jpeg' }
        '.ico'  { 'image/x-icon' }
        default { 'image/png' }
    }
    $b64 = [Convert]::ToBase64String($bytes)
    return @{ Base64 = $b64; MimeType = $mime; DataUri = "data:$mime;base64,$b64" }
}

function Get-IconPreview([string]$IconPath, [int]$MaxSize=64) {
    try {
        $img = [System.Drawing.Image]::FromFile($IconPath)
        $ratio = [Math]::Min($MaxSize / $img.Width, $MaxSize / $img.Height)
        $newW = [int]($img.Width * $ratio)
        $newH = [int]($img.Height * $ratio)
        $thumb = $img.GetThumbnailImage($newW, $newH, $null, [IntPtr]::Zero)
        $img.Dispose()
        return $thumb
    } catch { return $null }
}

# =========================
# CREDENTIAL DIALOG
# =========================
function Show-CredentialDialog {
    param(
        [string]$Title = "Enter Credentials",
        [string]$Message = "Enter your BigFix credentials",
        [System.Windows.Forms.IWin32Window]$Owner = $null,
        [switch]$StayOnTop
    )
    
    $credForm = New-Object System.Windows.Forms.Form
    $credForm.Text = $Title
    $credForm.Size = New-Object System.Drawing.Size(400, 220)
    $credForm.StartPosition = "CenterParent"
    $credForm.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#2D2D30")
    $credForm.ForeColor = [System.Drawing.Color]::White
    $credForm.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $credForm.FormBorderStyle = "FixedDialog"
    $credForm.MaximizeBox = $false
    $credForm.MinimizeBox = $false
    $credForm.ShowInTaskbar = $false
    $credForm.TopMost = $StayOnTop.IsPresent
    
    $lblMsg = New-Object System.Windows.Forms.Label
    $lblMsg.Text = $Message
    $lblMsg.Location = New-Object System.Drawing.Point(15, 12)
    $lblMsg.AutoSize = $true
    $credForm.Controls.Add($lblMsg)
    
    $lblUser = New-Object System.Windows.Forms.Label
    $lblUser.Text = "Username:"
    $lblUser.Location = New-Object System.Drawing.Point(15, 45)
    $lblUser.AutoSize = $true
    $credForm.Controls.Add($lblUser)
    
    $tbCredUser = New-Object System.Windows.Forms.TextBox
    $tbCredUser.Location = New-Object System.Drawing.Point(120, 42)
    $tbCredUser.Size = New-Object System.Drawing.Size(240, 20)
    $tbCredUser.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#3C3C3C")
    $tbCredUser.ForeColor = [System.Drawing.Color]::White
    $tbCredUser.BorderStyle = "FixedSingle"
    $credForm.Controls.Add($tbCredUser)
    
    $lblPwd = New-Object System.Windows.Forms.Label
    $lblPwd.Text = "Password:"
    $lblPwd.Location = New-Object System.Drawing.Point(15, 80)
    $lblPwd.AutoSize = $true
    $credForm.Controls.Add($lblPwd)
    
    $tbCredPass = New-Object System.Windows.Forms.TextBox
    $tbCredPass.Location = New-Object System.Drawing.Point(120, 77)
    $tbCredPass.Size = New-Object System.Drawing.Size(240, 20)
    $tbCredPass.PasswordChar = '*'
    $tbCredPass.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#3C3C3C")
    $tbCredPass.ForeColor = [System.Drawing.Color]::White
    $tbCredPass.BorderStyle = "FixedSingle"
    $credForm.Controls.Add($tbCredPass)
    
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Size = New-Object System.Drawing.Size(80, 30)
    $btnOK.Location = New-Object System.Drawing.Point(180, 120)
    $btnOK.FlatStyle = "Flat"
    $btnOK.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#0078D7")
    $btnOK.ForeColor = [System.Drawing.Color]::White
    $btnOK.FlatAppearance.BorderSize = 0
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $credForm.Controls.Add($btnOK)
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Size = New-Object System.Drawing.Size(80, 30)
    $btnCancel.Location = New-Object System.Drawing.Point(270, 120)
    $btnCancel.FlatStyle = "Flat"
    $btnCancel.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#555555")
    $btnCancel.ForeColor = [System.Drawing.Color]::White
    $btnCancel.FlatAppearance.BorderSize = 0
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $credForm.Controls.Add($btnCancel)
    
    $credForm.AcceptButton = $btnOK
    $credForm.CancelButton = $btnCancel

    $credForm.Add_Shown({
        $credForm.Activate()
        $credForm.BringToFront()
        $tbCredUser.Focus()
    })

    $result = if ($Owner) {
        $credForm.ShowDialog($Owner)
    } else {
        $credForm.ShowDialog()
    }
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $tbCredUser.Text.Trim() -and $tbCredPass.Text) {
        return @{ User = $tbCredUser.Text.Trim(); Pass = $tbCredPass.Text }
    }
    return $null
}

# =========================
# HTTP FUNCTIONS
# =========================
if ($IgnoreCertErrors) {
    try { [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } } catch {}
}
[System.Net.ServicePointManager]::Expect100Continue = $false

function Post-Xml {
    param([string]$Url,[string]$User,[string]$Pass,[string]$XmlBody)
    $tmpFile = Join-Path $env:TEMP ("BES_Post_{0:yyyyMMdd_HHmmss}_{1}.xml" -f (Get-Date), (Get-Random))
    [System.IO.File]::WriteAllText($tmpFile, $XmlBody, [Text.Encoding]::UTF8)
    if ($SaveXmlToTemp) { LogLine "Saved XML to: $tmpFile" }
    $pair = "$User`:$Pass"
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    $basic = "Basic " + [Convert]::ToBase64String($bytes)
    $resp = Invoke-WebRequest -Method Post -Uri $Url `
        -Headers @{ "Authorization" = $basic } `
        -ContentType "application/xml" `
        -InFile $tmpFile `
        -UseBasicParsing `
        -ErrorAction Stop
    LogLine ("POST HTTP {0}" -f [int]$resp.StatusCode)
    return $resp
}

# =========================
# PSADT PARSING
# =========================
function Parse-PsadtFolder {
    param([string]$FolderPath)
    
    $result = @{
        Vendor        = ""
        AppName       = ""
        Version       = ""
        PsadtExeName  = ""
        InstallerType = "Unknown"
        InstallerFile = ""
        FilesFolder   = ""
    }
    
    if (-not (Test-Path $FolderPath)) { throw "PSADT folder not found: $FolderPath" }
    
    # Find the renamed Invoke-AppDeployToolkit exe
    $exeFiles = Get-ChildItem -Path $FolderPath -Filter "*.exe" -File |
        Where-Object { $_.Name -notmatch '(?i)ServiceUI|Deploy-Application|Invoke-AppDeployToolkit\.exe$' }
    $psadtExe = $exeFiles | Where-Object { $_.Name -match '^[A-Za-z].*-[\d]' } | Select-Object -First 1
    
    if (-not $psadtExe) {
        # Fallback: look for the renamed Invoke-AppDeployToolkit pattern
        $psadtExe = Get-ChildItem -Path $FolderPath -Filter "*.exe" -File |
            Where-Object { $_.Name -ne "Invoke-AppDeployToolkit.exe" -and $_.Name -ne "ServiceUI.exe" } |
            Select-Object -First 1
    }
    
    if ($psadtExe) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($psadtExe.Name)
        # Pattern: VendorAppName-Version (e.g., GoogleChrome-131.0.6778.265)
        if ($baseName -match '^(.+?)-(\d+[\d\.]+.*)$') {
            $result.AppName = $Matches[1]
            $result.Version = $Matches[2]
        } else {
            $result.AppName = $baseName
        }
        $result.PsadtExeName = $psadtExe.Name
    }
    
    # Check Files folder for installer
    $filesFolder = Join-Path $FolderPath "Files"
    if (Test-Path $filesFolder) {
        $result.FilesFolder = $filesFolder
        $msiFiles = Get-ChildItem -Path $filesFolder -Filter "*.msi" -File
        $exeInstallers = Get-ChildItem -Path $filesFolder -Filter "*.exe" -File
        
        if ($msiFiles.Count -gt 0) {
            $result.InstallerType = "MSI"
            $result.InstallerFile = $msiFiles[0].Name
        } elseif ($exeInstallers.Count -gt 0) {
            $result.InstallerType = "EXE"
            $result.InstallerFile = $exeInstallers[0].Name
        }
    }
    
    return $result
}

# =========================
# PSADT RECIPE HELPERS
# =========================
$RecipeSectionNames = @(
    "PreInstall",
    "Install",
    "PostInstall",
    "PreUninstall",
    "Uninstall",
    "PostUninstall"
)
$SupportedRecipeTokens = @(
    "{{Vendor}}",
    "{{AppName}}",
    "{{AppVersion}}",
    "{{InstallerType}}",
    "{{InstallerFile}}",
    "{{FilesDirectory}}"
)

function Import-PsadtRecipe {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Recipe file not found: $Path"
    }

    try {
        $recipe = Import-PowerShellDataFile -LiteralPath $Path -ErrorAction Stop
    } catch {
        throw "Could not read recipe '$Path': $($_.Exception.Message)"
    }

    if (-not $recipe) { throw "Recipe '$Path' is empty." }
    if ([string]$recipe.SchemaVersion -ne "1.0") {
        throw "Recipe '$Path' uses unsupported SchemaVersion '$($recipe.SchemaVersion)'. Expected 1.0."
    }
    if ([string]::IsNullOrWhiteSpace([string]$recipe.Name)) {
        throw "Recipe '$Path' is missing Name."
    }
    if ($null -ne $recipe.FixletIconPath -and -not ($recipe.FixletIconPath -is [string])) {
        throw "Recipe '$Path' FixletIconPath must be a string."
    }
    if (-not $recipe.Sections -or -not ($recipe.Sections -is [System.Collections.IDictionary])) {
        throw "Recipe '$Path' must contain a Sections hashtable."
    }

    foreach ($sectionName in $recipe.Sections.Keys) {
        if ($RecipeSectionNames -notcontains [string]$sectionName) {
            throw "Recipe '$Path' contains unknown section '$sectionName'."
        }
        $sectionValue = $recipe.Sections[$sectionName]
        if ($null -ne $sectionValue -and -not ($sectionValue -is [string])) {
            throw "Recipe section '$sectionName' in '$Path' must be a string."
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$sectionValue)) {
            $unknownTokens = [regex]::Matches([string]$sectionValue, '\{\{[^{}]+\}\}') |
                ForEach-Object { $_.Value } |
                Where-Object { $SupportedRecipeTokens -notcontains $_ } |
                Select-Object -Unique
            if ($unknownTokens) {
                throw "Recipe section '$sectionName' contains unknown token(s): $($unknownTokens -join ', ')"
            }

            # Validate syntax when the selected recipe is imported. Substitute
            # harmless sample values because package metadata is not known yet.
            $syntaxSample = [string]$sectionValue
            foreach ($token in $SupportedRecipeTokens) {
                $sampleValue = if ($token -eq "{{FilesDirectory}}") {
                    '$adtSession.DirFiles'
                } else {
                    "RecipeValidationValue"
                }
                $syntaxSample = $syntaxSample.Replace($token, $sampleValue)
            }
            $parseTokens = $null
            $parseErrors = $null
            [void][System.Management.Automation.Language.Parser]::ParseInput(
                $syntaxSample,
                [ref]$parseTokens,
                [ref]$parseErrors
            )
            if ($parseErrors.Count -gt 0) {
                $messages = $parseErrors | ForEach-Object { $_.Message } | Select-Object -Unique
                throw "Recipe section '$sectionName' has invalid PowerShell: $($messages -join '; ')"
            }
        }
    }

    $recipe["_SourcePath"] = (Resolve-Path -LiteralPath $Path).Path
    $recipe["_Sha256"] = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    if (-not [string]::IsNullOrWhiteSpace([string]$recipe.FixletIconPath)) {
        $iconPath = [Environment]::ExpandEnvironmentVariables(([string]$recipe.FixletIconPath).Trim())
        if (-not [System.IO.Path]::IsPathRooted($iconPath)) {
            $iconPath = Join-Path (Split-Path -Parent $recipe["_SourcePath"]) $iconPath
        }
        if (-not (Test-Path -LiteralPath $iconPath -PathType Leaf)) {
            throw "Recipe '$Path' FixletIconPath was not found: $iconPath"
        }
        if ([System.IO.Path]::GetExtension($iconPath) -notin @(".png", ".jpg", ".jpeg", ".ico")) {
            throw "Recipe '$Path' FixletIconPath must be a PNG, JPG, JPEG, or ICO file."
        }
        $recipe["_ResolvedFixletIconPath"] = (Resolve-Path -LiteralPath $iconPath).Path
    }
    return $recipe
}

function Expand-PsadtRecipeTokens {
    param(
        [AllowEmptyString()][string]$Text,
        [string]$Vendor,
        [string]$AppName,
        [string]$AppVersion,
        [string]$InstallerType,
        [string]$InstallerFile
    )

    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }

    $tokens = [ordered]@{
        "{{Vendor}}"         = $Vendor
        "{{AppName}}"        = $AppName
        "{{AppVersion}}"     = $AppVersion
        "{{InstallerType}}"  = $InstallerType
        "{{InstallerFile}}"  = $InstallerFile
        "{{FilesDirectory}}" = '$adtSession.DirFiles'
    }

    $expanded = $Text
    foreach ($token in $tokens.Keys) {
        if ($expanded.Contains($token) -and [string]::IsNullOrWhiteSpace([string]$tokens[$token])) {
            throw "Recipe token '$token' requires a value, but the package field is blank."
        }
        $expanded = $expanded.Replace($token, [string]$tokens[$token])
    }

    $unresolved = [regex]::Matches($expanded, '\{\{[^{}]+\}\}') |
        ForEach-Object { $_.Value } |
        Select-Object -Unique
    if ($unresolved) {
        throw "Unknown recipe token(s): $($unresolved -join ', ')"
    }

    # Remove only surrounding line breaks. A general Trim() also removes the
    # first line's intentional indentation while leaving later lines untouched.
    return $expanded.Trim([char[]]@("`r", "`n"))
}

function Resolve-PsadtRecipeSections {
    param(
        [System.Collections.IDictionary]$Recipe,
        [string]$Vendor,
        [string]$AppName,
        [string]$AppVersion,
        [string]$InstallerType,
        [string]$InstallerFile
    )

    $resolved = @{}
    foreach ($sectionName in $RecipeSectionNames) {
        $raw = if ($Recipe.Sections.Contains($sectionName)) {
            [string]$Recipe.Sections[$sectionName]
        } else {
            ""
        }

        $expanded = Expand-PsadtRecipeTokens `
            -Text $raw `
            -Vendor $Vendor `
            -AppName $AppName `
            -AppVersion $AppVersion `
            -InstallerType $InstallerType `
            -InstallerFile $InstallerFile

        if (-not [string]::IsNullOrWhiteSpace($expanded)) {
            $parseTokens = $null
            $parseErrors = $null
            [void][System.Management.Automation.Language.Parser]::ParseInput(
                $expanded,
                [ref]$parseTokens,
                [ref]$parseErrors
            )
            if ($parseErrors.Count -gt 0) {
                $messages = $parseErrors | ForEach-Object { $_.Message } | Select-Object -Unique
                throw "Recipe section '$sectionName' has invalid PowerShell: $($messages -join '; ')"
            }
        }

        $resolved[$sectionName] = $expanded
    }

    return $resolved
}

function Format-PsadtRecipeBlock {
    param(
        [AllowEmptyString()][string]$Text,
        [string]$SectionName,
        [string]$RecipeName
    )

    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }

    return @"

    # Recipe: $RecipeName [$SectionName]
$($Text.Trim([char[]]@("`r", "`n")))
    # End Recipe: $RecipeName [$SectionName]
"@
}

function Save-PsadtRecipeSnapshot {
    param(
        [System.Collections.IDictionary]$Recipe,
        [Parameter(Mandatory = $true)][string]$PackageFolder
    )

    if (-not $Recipe -or -not $Recipe["_SourcePath"]) { return $null }

    $snapshotFolder = Join-Path $PackageFolder "SupportFiles\RecipeSnapshots"
    if (-not (Test-Path -LiteralPath $snapshotFolder)) {
        New-Item -ItemType Directory -Path $snapshotFolder -Force | Out-Null
    }

    $safeName = ([string]$Recipe.Name -replace '[^\w\-.]', '_').Trim([char]'_')
    if (-not $safeName) { $safeName = "PSADT-Recipe" }
    $recipeVersion = if ($Recipe.RecipeVersion) {
        "-" + ([string]$Recipe.RecipeVersion -replace '[^\w\-.]', '_')
    } else {
        ""
    }
    $snapshotPath = Join-Path $snapshotFolder "$safeName$recipeVersion.psd1"
    Copy-Item -LiteralPath $Recipe["_SourcePath"] -Destination $snapshotPath -Force
    return $snapshotPath
}

# =========================
# PSADT SCRIPT GENERATOR (4.x)
# =========================
function Generate-DeployScript {
    param(
        [string]$Vendor,
        [string]$AppName,
        [string]$AppVersion,
        [string[]]$ProcessesToKill,
        [string]$ShortcutName,
        [string]$Author,
        [string]$OutputPath,
        [string]$InstallerType = "",
        [string]$InstallerFile = "",
        [System.Collections.IDictionary]$Recipe
    )
    
    # Build process list string for the template
    $procListStr = ($ProcessesToKill | ForEach-Object { "`"$_`"" }) -join ", "
    
    # Build process kill block
    $processKillBlock = @"
    # Kill Processes
    `$processes = @($procListStr)
    forEach (`$process in `$processes) {
        if (Get-Process -Name `$process -ErrorAction SilentlyContinue) {
            Write-ADTLogEntry -Message "Stopping the process '`$process'..." -Source `$adtSession.InstallPhase
            Stop-Process -Name `$process -Force -ErrorAction SilentlyContinue
        } else {
            Write-ADTLogEntry -Message "'`$process' is not running." -Source `$adtSession.InstallPhase
        }
    }
"@
    
    # Build shortcut removal block
    $shortcutBlock = ""
    if ($ShortcutName) {
        $shortcutBlock = @"

    # Remove desktop shortcut
    `$shortcutPath = "`$env:PUBLIC\Desktop\$ShortcutName.lnk"
    if (Test-Path `$shortcutPath) {
        Write-ADTLogEntry -Message "Removing desktop shortcut: `$shortcutPath" -Source `$adtSession.InstallPhase
        Remove-Item `$shortcutPath -Force -ErrorAction SilentlyContinue
    }
"@
    }
    
    # Build install/uninstall command blocks based on installer type
    $installBlock = "    ## <Perform Installation tasks here>"
    $uninstallBlock = "    ## <Perform Uninstallation tasks here>"
    
    if ($InstallerType -eq "MSI" -and $InstallerFile) {
        $installBlock = @"
    # Install MSI
    `$msiPath = Join-Path `$adtSession.DirFiles '$InstallerFile'
    Write-ADTLogEntry -Message "Installing `$msiPath" -Source `$adtSession.InstallPhase
    Start-ADTMsiProcess -Action Install -FilePath `$msiPath -ArgumentList 'ALLUSERS=1 REBOOT=ReallySuppress /QN'
"@
        $uninstallBlock = @"
    # Uninstall MSI
    `$msiPath = Join-Path `$adtSession.DirFiles '$InstallerFile'
    Write-ADTLogEntry -Message "Uninstalling `$msiPath" -Source `$adtSession.InstallPhase
    Start-ADTMsiProcess -Action Uninstall -FilePath `$msiPath -ArgumentList 'REBOOT=ReallySuppress /QN'
"@
    } elseif ($InstallerType -eq "EXE" -and $InstallerFile) {
        $installBlock = @"
    # Install EXE
    `$exePath = Join-Path `$adtSession.DirFiles '$InstallerFile'
    Write-ADTLogEntry -Message "Installing `$exePath" -Source `$adtSession.InstallPhase
    Start-ADTProcess -FilePath `$exePath -ArgumentList '/S /v/qn' -WaitForMsiExec
"@
        $uninstallBlock = @"
    # Uninstall EXE - UPDATE THE UNINSTALL STRING FROM REGISTRY
    ## Find uninstall string: Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall | Get-ItemProperty | Where DisplayName -match '$AppName'
    Write-ADTLogEntry -Message "Uninstalling $Vendor $AppName" -Source `$adtSession.InstallPhase
    ## Start-ADTProcess -FilePath 'UNINSTALL_PATH_HERE' -ArgumentList '/S /v/qn' -WaitForMsiExec
"@
    }

    $recipeName = ""
    $recipeVersion = ""
    $recipeProvenance = ""
    $preInstallRecipeBlock = ""
    $postInstallRecipeBlock = ""
    $preUninstallRecipeBlock = ""
    $postUninstallRecipeBlock = ""

    if ($Recipe) {
        $recipeName = [string]$Recipe.Name
        $recipeVersion = [string]$Recipe.RecipeVersion
        $recipeSections = Resolve-PsadtRecipeSections `
            -Recipe $Recipe `
            -Vendor $Vendor `
            -AppName $AppName `
            -AppVersion $AppVersion `
            -InstallerType $InstallerType `
            -InstallerFile $InstallerFile

        $preInstallRecipeBlock = Format-PsadtRecipeBlock `
            -Text $recipeSections.PreInstall -SectionName "Pre-Install" -RecipeName $recipeName
        $postInstallRecipeBlock = Format-PsadtRecipeBlock `
            -Text $recipeSections.PostInstall -SectionName "Post-Install" -RecipeName $recipeName
        $preUninstallRecipeBlock = Format-PsadtRecipeBlock `
            -Text $recipeSections.PreUninstall -SectionName "Pre-Uninstall" -RecipeName $recipeName
        $postUninstallRecipeBlock = Format-PsadtRecipeBlock `
            -Text $recipeSections.PostUninstall -SectionName "Post-Uninstall" -RecipeName $recipeName

        # Recipe Install/Uninstall blocks replace the automatic command so the
        # application is never installed or uninstalled twice.
        if (-not [string]::IsNullOrWhiteSpace($recipeSections.Install)) {
            $installBlock = Format-PsadtRecipeBlock `
                -Text $recipeSections.Install `
                -SectionName "Install (replaces automatic command)" `
                -RecipeName $recipeName
        }
        if (-not [string]::IsNullOrWhiteSpace($recipeSections.Uninstall)) {
            $uninstallBlock = Format-PsadtRecipeBlock `
                -Text $recipeSections.Uninstall `
                -SectionName "Uninstall (replaces automatic command)" `
                -RecipeName $recipeName
        }

        $recipeProvenance = @"
	- PSADT recipe: $recipeName$(if ($recipeVersion) { " v$recipeVersion" })
	- Recipe source: $($Recipe["_SourcePath"])
	- Recipe SHA-256: $($Recipe["_Sha256"])
"@
    }
    
    $scriptDate = Get-Date -Format 'yyyy-MM-dd'
    
    $script = @"
<#
.SYNOPSIS
	PSAppDeployToolkit - $Vendor $AppName $AppVersion
.DESCRIPTION
	- The script either performs an "Install", "Uninstall", or "Repair" deployment type.
	- The install deployment type is broken down into 3 main sections/phases: Pre-Install, Install, and Post-Install.
	The script imports the PSAppDeployToolkit module which contains the logic and functions required to install or uninstall an application.
$recipeProvenance
.PARAMETER DeploymentType
	The type of deployment to perform.
.PARAMETER DeployMode
	Specifies whether the installation should be run in Interactive, Silent, NonInteractive, or Auto mode.
.PARAMETER SuppressRebootPassThru
	Suppresses the 3010 return code (requires restart) from being passed back to the parent process.
.PARAMETER TerminalServerMode
	Changes to "user install mode" and back to "user execute mode" for Remote Desktop Session Hosts/Citrix servers.
.PARAMETER DisableLogging
	Disables logging to file for the script.
.LINK
	https://psappdeploytoolkit.com
#>

[CmdletBinding()]
param
(
	[Parameter(Mandatory = `$false)]
	[ValidateSet('Install', 'Uninstall', 'Repair')]
	[System.String]`$DeploymentType,

	[Parameter(Mandatory = `$false)]
	[ValidateSet('Auto', 'Interactive', 'NonInteractive', 'Silent')]
	[System.String]`$DeployMode,

	[Parameter(Mandatory = `$false)]
	[System.Management.Automation.SwitchParameter]`$SuppressRebootPassThru,

	[Parameter(Mandatory = `$false)]
	[System.Management.Automation.SwitchParameter]`$TerminalServerMode,

	[Parameter(Mandatory = `$false)]
	[System.Management.Automation.SwitchParameter]`$DisableLogging
)


##================================================
## MARK: Variables
##================================================
`$adtSession = @{
	# App variables.
	AppVendor = '$Vendor'
	AppName = '$AppName'
	AppVersion = '$AppVersion'
	AppArch = ''
	AppLang = 'EN'
	AppRevision = '01'
	AppSuccessExitCodes = @(0)
	AppRebootExitCodes = @(1641, 3010)
	AppProcessesToClose = @()
	AppScriptVersion = '1.0.0'
	AppScriptDate = '$scriptDate'
	AppScriptAuthor = '$Author'
	RequireAdmin = `$true

	# Install Titles (Only set here to override defaults set by the toolkit).
	InstallName = ''
	InstallTitle = ''

	# Script variables.
	DeployAppScriptFriendlyName = `$MyInvocation.MyCommand.Name
	DeployAppScriptParameters = `$PSBoundParameters
	DeployAppScriptVersion = '4.1.8'
}


function Install-ADTDeployment
{
	[CmdletBinding()]
	param
	(
	)

	##================================================
	## MARK: Pre-Install
	##================================================
	`$adtSession.InstallPhase = "Pre-`$(`$adtSession.DeploymentType)"

	## Show Progress Message (with the default message).
	Show-ADTInstallationProgress

$processKillBlock
$preInstallRecipeBlock


	##================================================
	## MARK: Install
	##================================================
	`$adtSession.InstallPhase = `$adtSession.DeploymentType

$installBlock


	##================================================
	## MARK: Post-Install
	##================================================
	`$adtSession.InstallPhase = "Post-`$(`$adtSession.DeploymentType)"
$shortcutBlock
$postInstallRecipeBlock
}


function Uninstall-ADTDeployment
{
	[CmdletBinding()]
	param
	(
	)

	##================================================
	## MARK: Pre-Uninstall
	##================================================
	`$adtSession.InstallPhase = "Pre-`$(`$adtSession.DeploymentType)"

	## Show Progress Message (with the default message).
	Show-ADTInstallationProgress

$processKillBlock
$preUninstallRecipeBlock


	##================================================
	## MARK: Uninstall
	##================================================
	`$adtSession.InstallPhase = `$adtSession.DeploymentType

$uninstallBlock


	##================================================
	## MARK: Post-Uninstallation
	##================================================
	`$adtSession.InstallPhase = "Post-`$(`$adtSession.DeploymentType)"

	## <Perform Post-Uninstallation tasks here>
$postUninstallRecipeBlock
}


function Repair-ADTDeployment
{
	[CmdletBinding()]
	param
	(
	)

	##================================================
	## MARK: Pre-Repair
	##================================================
	`$adtSession.InstallPhase = "Pre-`$(`$adtSession.DeploymentType)"

	if (`$adtSession.AppProcessesToClose.Count -gt 0)
	{
		Show-ADTInstallationWelcome -CloseProcesses `$adtSession.AppProcessesToClose -CloseProcessesCountdown 60
	}

	Show-ADTInstallationProgress

	##================================================
	## MARK: Repair
	##================================================
	`$adtSession.InstallPhase = `$adtSession.DeploymentType

	if (`$adtSession.UseDefaultMsi)
	{
		`$ExecuteDefaultMSISplat = @{ Action = `$adtSession.DeploymentType; FilePath = `$adtSession.DefaultMsiFile }
		if (`$adtSession.DefaultMstFile)
		{
			`$ExecuteDefaultMSISplat.Add('Transforms', `$adtSession.DefaultMstFile)
		}
		Start-ADTMsiProcess @ExecuteDefaultMSISplat
	}

	##================================================
	## MARK: Post-Repair
	##================================================
	`$adtSession.InstallPhase = "Post-`$(`$adtSession.DeploymentType)"
}


##================================================
## MARK: Initialization
##================================================
`$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
`$ProgressPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
Set-StrictMode -Version 1

try
{
	if (Test-Path -LiteralPath "`$PSScriptRoot\PSAppDeployToolkit\PSAppDeployToolkit.psd1" -PathType Leaf)
	{
		Get-ChildItem -LiteralPath "`$PSScriptRoot\PSAppDeployToolkit" -Recurse -File | Unblock-File -ErrorAction Ignore
		Import-Module -FullyQualifiedName @{ ModuleName = "`$PSScriptRoot\PSAppDeployToolkit\PSAppDeployToolkit.psd1"; Guid = '8c3c366b-8606-4576-9f2d-4051144f7ca2'; ModuleVersion = '4.1.8' } -Force
	}
	else
	{
		Import-Module -FullyQualifiedName @{ ModuleName = 'PSAppDeployToolkit'; Guid = '8c3c366b-8606-4576-9f2d-4051144f7ca2'; ModuleVersion = '4.1.8' } -Force
	}

	`$iadtParams = Get-ADTBoundParametersAndDefaultValues -Invocation `$MyInvocation
	`$adtSession = Remove-ADTHashtableNullOrEmptyValues -Hashtable `$adtSession
	`$adtSession = Open-ADTSession @adtSession @iadtParams -PassThru
}
catch
{
	`$Host.UI.WriteErrorLine((Out-String -InputObject `$_ -Width ([System.Int32]::MaxValue)))
	exit 60008
}


##================================================
## MARK: Invocation
##================================================
try
{
	Get-ChildItem -LiteralPath `$PSScriptRoot -Directory | & {
		process
		{
			if (`$_.Name -match 'PSAppDeployToolkit\..+`$')
			{
				Get-ChildItem -LiteralPath `$_.FullName -Recurse -File | Unblock-File -ErrorAction Ignore
				Import-Module -Name `$_.FullName -Force
			}
		}
	}

	& "`$(`$adtSession.DeploymentType)-ADTDeployment"
	Close-ADTSession
}
catch
{
	`$mainErrorMessage = "An unhandled error within [`$(`$MyInvocation.MyCommand.Name)] has occurred.``n`$(Resolve-ADTErrorRecord -ErrorRecord `$_)"
	Write-ADTLogEntry -Message `$mainErrorMessage -Severity 3
	Close-ADTSession -ExitCode 60001
}
"@
    
    [System.IO.File]::WriteAllText($OutputPath, $script, [System.Text.Encoding]::UTF8)
    return $OutputPath
}

# =========================
# FIXLET XML BUILDER
# =========================
function Build-FixletXml {
    param(
        [string]$Title,
        [string]$Description,
        [string]$Relevance,
        [string]$ActionScript,
        [string]$Category,
        [string]$Source = "Internal",
        [string]$SourceSeverity = "",
        [string]$IconBase64DataUri = ""
    )
    
    $titleEsc = SafeEscape $Title
    $catEsc = SafeEscape $Category
    $sevEsc = SafeEscape $SourceSeverity
    $relCdata = $Relevance -replace ']]>', ']]]]><![CDATA[>'
    $actCdata = $ActionScript -replace ']]>', ']]]]><![CDATA[>'
    
    # Build icon element if provided
    $iconElement = ""
    if ($IconBase64DataUri) {
        $iconJson = "{`"icon`":`"$IconBase64DataUri`"}"
        $iconElement = "`n    <MIMEField>`n      <Name>action-ui-metadata</Name>`n      <Value>$iconJson</Value>`n    </MIMEField>"
    }
    
@"
<?xml version="1.0" encoding="UTF-8"?>
<BES xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="BES.xsd">
  <Fixlet>
    <Title>$titleEsc</Title>
    <Description>$(SafeEscape $Description)</Description>
    <Relevance><![CDATA[$relCdata]]></Relevance>
    <Category>$catEsc</Category>
    <Source>$(SafeEscape $Source)</Source>
    <SourceSeverity>$sevEsc</SourceSeverity>$iconElement
    <DefaultAction ID="Action1">
      <ActionScript MIMEType="application/x-Fixlet-Windows-Shell"><![CDATA[$actCdata]]></ActionScript>
    </DefaultAction>
  </Fixlet>
</BES>
"@
}

# =========================
# OFFER XML BUILDER
# =========================
function Build-OfferXml {
    param(
        [string]$DisplayName,
        [string]$SiteName,
        [string]$FixletId,
        [string]$FixletActionName,
        [string]$GroupRelevance,
        [string]$Kind,
        [string]$Phase,
        [string]$OfferDescription,
        [string]$IconBase64DataUri = ""
    )
    
    $siteEsc = SafeEscape $SiteName
    $fixletIdEsc = SafeEscape $FixletId
    $actionNameEsc = SafeEscape $FixletActionName
    $groupSafe = if ($GroupRelevance) { $GroupRelevance -replace ']]>', ']]]]><![CDATA[>' } else { "" }
    
    switch -Regex ($Kind) {
        '^(?i)install$' { $ing='Installing'; $cat='Install' }
        '^(?i)remove$'  { $ing='Removing';  $cat='Remove'  }
        default          { $ing='Updating';  $cat='Update'  }
    }
    
    $runningMsg = SafeEscape ("{0} {1}. Please wait..." -f $ing, $DisplayName)
    $offerTitle = SafeEscape ("{0}: {1} Win: {2} Offer" -f $cat, $DisplayName, $Phase)
    $descFallback = "This offer will $($cat.ToLower()) $DisplayName."
    $descRaw = if ([string]::IsNullOrWhiteSpace($OfferDescription)) { $descFallback } else { $OfferDescription }
    $descHtml = [System.Web.HttpUtility]::HtmlEncode($descRaw) -replace "`r?`n","<br/>"
    $offerCat = SafeEscape $cat
    $offerIconElement = ""
    if ($IconBase64DataUri) {
        $offerIconElement = ""
    }
    
@"
<?xml version="1.0" encoding="UTF-8"?>
<BES xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="BES.xsd">
  <SourcedFixletAction>
    <SourceFixlet>
      <Sitename>$siteEsc</Sitename>
      <FixletID>$fixletIdEsc</FixletID>
      <Action>$actionNameEsc</Action>
    </SourceFixlet>
    <Target>
      <CustomRelevance><![CDATA[$groupSafe]]></CustomRelevance>
    </Target>
    <Settings>
      <ActionUITitle>$(if ($cat -in 'Update','Remove') { SafeEscape ("{0}: {1}" -f $cat, $DisplayName) } else { SafeEscape $DisplayName })</ActionUITitle>
      <PreActionShowUI>$(To-XmlBool $OfferDefaults.PreActionShowUI)</PreActionShowUI>
      <HasRunningMessage>true</HasRunningMessage>
      <RunningMessage><Text>$runningMsg</Text></RunningMessage>
      <HasTimeRange>false</HasTimeRange>
      <HasStartTime>true</HasStartTime>
      <StartDateTimeLocalOffset>$($OfferDefaults.StartOffsetISO)</StartDateTimeLocalOffset>
      <HasEndTime>true</HasEndTime>
      <EndDateTimeLocalOffset>$($OfferDefaults.EndOffsetISO)</EndDateTimeLocalOffset>
      <UseUTCTime>false</UseUTCTime>
      <Reapply>$(To-XmlBool $OfferDefaults.Reapply)</Reapply>
      <HasReapplyLimit>false</HasReapplyLimit>
      <HasReapplyInterval>false</HasReapplyInterval>
      <HasRetry>true</HasRetry>
      <RetryCount>$($OfferDefaults.RetryCount)</RetryCount>
      <RetryWait Behavior="WaitForInterval">$($OfferDefaults.RetryWaitISO)</RetryWait>
      <HasTemporalDistribution>false</HasTemporalDistribution>
      <ContinueOnErrors>$(To-XmlBool $OfferDefaults.ContinueOnErr)</ContinueOnErrors>
      <PostActionBehavior Behavior="Nothing"></PostActionBehavior>
      <IsOffer>true</IsOffer>
      <OfferCategory>$offerCat</OfferCategory>
      <OfferDescriptionHTML><![CDATA[$descHtml]]></OfferDescriptionHTML>
    </Settings>
    <Title>$offerTitle</Title>$offerIconElement
  </SourcedFixletAction>
</BES>
"@
}

# =========================
# GUI
# =========================
$form = New-Object System.Windows.Forms.Form
$form.Text = "BigFix Unified Packager v$ToolVer"
$form.Size = New-Object System.Drawing.Size(940, 850)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#2D2D30")
$form.ForeColor = [System.Drawing.Color]::White
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.AutoScroll = $true
$form.MinimumSize = New-Object System.Drawing.Size(920, 600)

function Style-TextBox($tb) {
    $tb.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#3C3C3C")
    $tb.ForeColor = [System.Drawing.Color]::White
    $tb.BorderStyle = "FixedSingle"
}

function Style-Button($btn, [string]$Color = "#0078D7") {
    $btn.FlatStyle = "Flat"
    $btn.BackColor = [System.Drawing.ColorTranslator]::FromHtml($Color)
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.FlatAppearance.BorderSize = 0
    $btn.FlatAppearance.MouseOverBackColor = [System.Drawing.ColorTranslator]::FromHtml("#3399FF")
}

function New-StyledTextBox($x, $y, $width=580, $height=20, $multiline=$false) {
    $box = New-Object System.Windows.Forms.TextBox
    $box.Location = New-Object System.Drawing.Point($x, $y)
    $box.Size = New-Object System.Drawing.Size($width, $height)
    $box.Multiline = $multiline
    if ($multiline) { $box.ScrollBars = "Vertical" }
    Style-TextBox $box
    return $box
}

function Add-Label($text, $x, $y, [switch]$Bold) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $text
    $lbl.Location = New-Object System.Drawing.Point($x, $y)
    $lbl.AutoSize = $true
    if ($Bold) { $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold) }
    $form.Controls.Add($lbl)
    return $lbl
}

$y = 10

# --- Section: PSADT Package ---
Add-Label "PSADT Folder:" 10 $y | Out-Null
$tbPsadtPath = New-StyledTextBox 130 $y 430
$tbPsadtPath.Text = ""
$form.Controls.Add($tbPsadtPath)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse..."
$btnBrowse.Size = New-Object System.Drawing.Size(80, 25)
$btnBrowse.Location = New-Object System.Drawing.Point(570, $y)
Style-Button $btnBrowse
$form.Controls.Add($btnBrowse)

$y += 30
Add-Label "PSADT Recipe:" 10 $y | Out-Null
$cbRecipe = New-Object System.Windows.Forms.ComboBox
$cbRecipe.Location = New-Object System.Drawing.Point(130, $y)
$cbRecipe.Size = New-Object System.Drawing.Size(430, 25)
$cbRecipe.DropDownStyle = "DropDownList"
$form.Controls.Add($cbRecipe)

$btnPreviewRecipe = New-Object System.Windows.Forms.Button
$btnPreviewRecipe.Text = "Preview"
$btnPreviewRecipe.Size = New-Object System.Drawing.Size(80, 25)
$btnPreviewRecipe.Location = New-Object System.Drawing.Point(570, $y)
$btnPreviewRecipe.Enabled = $false
Style-Button $btnPreviewRecipe "#555555"
$form.Controls.Add($btnPreviewRecipe)

$y += 28
$lblRecipeStatus = Add-Label "Basic package - no recipe selected" 130 $y
$lblRecipeStatus.ForeColor = [System.Drawing.Color]::Silver

$y += 28
Add-Label "Recipe Library:" 10 $y | Out-Null
$tbRecipePath = New-StyledTextBox 130 $y 520
$tbRecipePath.Text = $RecipeSharePath
$form.Controls.Add($tbRecipePath)

$btnBrowseRecipes = New-Object System.Windows.Forms.Button
$btnBrowseRecipes.Text = "Browse..."
$btnBrowseRecipes.Size = New-Object System.Drawing.Size(80, 25)
$btnBrowseRecipes.Location = New-Object System.Drawing.Point(660, $y)
Style-Button $btnBrowseRecipes
$form.Controls.Add($btnBrowseRecipes)

$btnRefreshRecipes = New-Object System.Windows.Forms.Button
$btnRefreshRecipes.Text = "Refresh"
$btnRefreshRecipes.Size = New-Object System.Drawing.Size(90, 25)
$btnRefreshRecipes.Location = New-Object System.Drawing.Point(750, $y)
Style-Button $btnRefreshRecipes "#555555"
$form.Controls.Add($btnRefreshRecipes)

$y += 45
Add-Label "- Package Info -" 10 $y -Bold | Out-Null
$y += 25

Add-Label "Vendor:" 10 $y | Out-Null
$tbVendor = New-StyledTextBox 180 $y 300
$form.Controls.Add($tbVendor)

$y += 30
Add-Label "Application Name:" 10 $y | Out-Null
$tbAppName = New-StyledTextBox 180 $y 300
$form.Controls.Add($tbAppName)

$y += 30
Add-Label "Package Version:" 10 $y | Out-Null
$tbPkgVersion = New-StyledTextBox 180 $y 300
$form.Controls.Add($tbPkgVersion)

$y += 30
Add-Label "Author:" 10 $y | Out-Null
$tbAuthor = New-StyledTextBox 180 $y 300
$tbAuthor.Text = $env:USERNAME
$form.Controls.Add($tbAuthor)

$y += 30
Add-Label "Packaging KB Link:" 10 $y | Out-Null
$tbKbLink = New-StyledTextBox 180 $y 300
$tbKbLink.ForeColor = [System.Drawing.Color]::Gray
$tbKbLink.Text = "(optional) ServiceNow KB article URL"
$tbKbLink.Add_GotFocus({
    if ($tbKbLink.Text -eq "(optional) ServiceNow KB article URL") {
        $tbKbLink.Text = ""
        $tbKbLink.ForeColor = [System.Drawing.SystemColors]::WindowText
    }
})
$tbKbLink.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($tbKbLink.Text)) {
        $tbKbLink.Text = "(optional) ServiceNow KB article URL"
        $tbKbLink.ForeColor = [System.Drawing.Color]::Gray
    }
})
$form.Controls.Add($tbKbLink)

# --- Section: Icon ---
$y += 40
Add-Label "- Fixlet Icon (for Self Service) -" 10 $y -Bold | Out-Null
$y += 25

Add-Label "Icon File:" 10 $y | Out-Null
$cbIcon = New-Object System.Windows.Forms.ComboBox
$cbIcon.Location = New-Object System.Drawing.Point(180, $y)
$cbIcon.Width = 350
$cbIcon.DropDownStyle = "DropDownList"
$form.Controls.Add($cbIcon)

$btnBrowseIcon = New-Object System.Windows.Forms.Button
$btnBrowseIcon.Text = "Browse..."
$btnBrowseIcon.Size = New-Object System.Drawing.Size(80, 25)
$btnBrowseIcon.Location = New-Object System.Drawing.Point(540, $y)
Style-Button $btnBrowseIcon
$form.Controls.Add($btnBrowseIcon)

$y += 30
$picIcon = New-Object System.Windows.Forms.PictureBox
$picIcon.Location = New-Object System.Drawing.Point(180, $y)
$picIcon.Size = New-Object System.Drawing.Size(64, 64)
$picIcon.SizeMode = "Zoom"
$picIcon.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#3C3C3C")
$form.Controls.Add($picIcon)

$lblIconStatus = Add-Label "(no icon selected)" 260 ($y + 20)

# Track selected icon path
$script:SelectedIconPath = $null
$script:SelectedIconFromRecipe = $false

function Set-SelectedFixletIcon {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$Source = "manual selection",
        [switch]$FromRecipe
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Icon file not found: $Path"
    }

    if ($picIcon.Image) {
        $picIcon.Image.Dispose()
        $picIcon.Image = $null
    }

    $script:SelectedIconPath = (Resolve-Path -LiteralPath $Path).Path
    $script:SelectedIconFromRecipe = $FromRecipe.IsPresent
    $cbIcon.Tag = $null
    $cbIcon.Items.Clear()
    [void]$cbIcon.Items.Add([System.IO.Path]::GetFileName($script:SelectedIconPath))
    $cbIcon.SelectedIndex = 0

    $preview = Get-IconPreview -IconPath $script:SelectedIconPath
    if ($preview) { $picIcon.Image = $preview }
    $lblIconStatus.Text = "[OK] Icon selected from $Source"
    $lblIconStatus.ForeColor = [System.Drawing.Color]::LightGreen
    LogLine "Icon selected from $Source`: $($script:SelectedIconPath)"
}

function Clear-RecipeFixletIcon {
    if (-not $script:SelectedIconFromRecipe) { return }
    if ($picIcon.Image) {
        $picIcon.Image.Dispose()
        $picIcon.Image = $null
    }
    $script:SelectedIconPath = $null
    $script:SelectedIconFromRecipe = $false
    $cbIcon.Tag = $null
    $cbIcon.Items.Clear()
    $lblIconStatus.Text = "(no icon selected)"
    $lblIconStatus.ForeColor = [System.Drawing.Color]::Silver
}

# Icon browse button
$btnBrowseIcon.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Image Files (*.png;*.jpg;*.jpeg;*.ico)|*.png;*.jpg;*.jpeg;*.ico|All Files (*.*)|*.*"
    $dlg.Title = "Select Icon File"
    if ($dlg.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        Set-SelectedFixletIcon -Path $dlg.FileName
    }
})

# Icon combo selection changed
$cbIcon.Add_SelectedIndexChanged({
    if ($cbIcon.SelectedItem -and $cbIcon.Tag) {
        $iconFiles = $cbIcon.Tag
        $selected = $iconFiles | Where-Object { $_.Name -eq $cbIcon.SelectedItem } | Select-Object -First 1
        if ($selected) {
            $script:SelectedIconPath = $selected.FullName
            $preview = Get-IconPreview -IconPath $selected.FullName
            if ($preview) { $picIcon.Image = $preview }
            $lblIconStatus.Text = "[OK] Icon selected"
            $lblIconStatus.ForeColor = [System.Drawing.Color]::LightGreen
        }
    }
})

$y += 70

# --- Section: PSADT Script Generation ---
$y += 40
Add-Label "- PSADT Script Generation -" 10 $y -Bold | Out-Null
$y += 25

Add-Label "Processes to Kill:" 10 $y | Out-Null
$tbProcesses = New-StyledTextBox 180 $y 500
$tbProcesses.Text = "(optional, comma-separated for multiple)"
$form.Controls.Add($tbProcesses)

$y += 30
Add-Label "Desktop Shortcut Name:" 10 $y | Out-Null
$tbShortcut = New-StyledTextBox 180 $y 400
$form.Controls.Add($tbShortcut)

$y += 28
$btnGenScript = New-Object System.Windows.Forms.Button
$btnGenScript.Text = "Generate PSADT Script + Open in ISE"
$btnGenScript.Size = New-Object System.Drawing.Size(300, 32)
$btnGenScript.Location = New-Object System.Drawing.Point(10, $y)
Style-Button $btnGenScript "#6A0DAD"
$form.Controls.Add($btnGenScript)

$btnSignScript = New-Object System.Windows.Forms.Button
$btnSignScript.Text = "Sign Script"
$btnSignScript.Size = New-Object System.Drawing.Size(120, 32)
$btnSignScript.Location = New-Object System.Drawing.Point(320, $y)
Style-Button $btnSignScript "#AA5500"
$form.Controls.Add($btnSignScript)

$lblSignStatus = Add-Label "" 450 ($y + 6)

# --- Section: Prefetch ---
$y += 45
Add-Label "- Prefetch / Extract (paste both lines from Software Upload Wizard) -" 10 $y -Bold | Out-Null
$y += 25

Add-Label "Prefetch + Extract:" 10 $y | Out-Null
$tbPrefetch = New-StyledTextBox 180 $y 700 80 $true
$form.Controls.Add($tbPrefetch)

# --- Section: Relevance ---
$y += 55
Add-Label "- Relevance -" 10 $y -Bold | Out-Null
$y += 25

Add-Label "Install:" 10 $y | Out-Null
$tbInstallRel = New-StyledTextBox 80 $y 800 30 $true
$form.Controls.Add($tbInstallRel)

$y += 38
Add-Label "Update:" 10 $y | Out-Null
$tbUpdateRel = New-StyledTextBox 80 $y 800 30 $true
$form.Controls.Add($tbUpdateRel)

$y += 38
Add-Label "Remove:" 10 $y | Out-Null
$tbRemoveRel = New-StyledTextBox 80 $y 800 30 $true
$form.Controls.Add($tbRemoveRel)

# --- Section: BigFix Connection ---
$y += 45
Add-Label "- BigFix Connection -" 10 $y -Bold | Out-Null
$y += 25

Add-Label "Server:" 10 $y | Out-Null
$cbServer = New-Object System.Windows.Forms.ComboBox
$cbServer.Location = New-Object System.Drawing.Point(180, $y)
$cbServer.Width = 350
$cbServer.DropDownStyle = "DropDownList"
foreach ($s in $ServerList) { $cbServer.Items.Add($s) | Out-Null }
$cbServer.SelectedIndex = 0
$form.Controls.Add($cbServer)

$y += 30
Add-Label "Site:" 10 $y | Out-Null
$cbSite = New-Object System.Windows.Forms.ComboBox
$cbSite.Location = New-Object System.Drawing.Point(180, $y)
$cbSite.Width = 400
$cbSite.DropDownStyle = "DropDownList"
foreach ($s in $SiteList) { $cbSite.Items.Add($s) | Out-Null }
$cbSite.SelectedIndex = 0
$form.Controls.Add($cbSite)

# --- Main Action Button ---
$y += 40
$btnPostAll = New-Object System.Windows.Forms.Button
$btnPostAll.Text = "POST Fixlets + Create QA Offers"
$btnPostAll.Size = New-Object System.Drawing.Size(320, 40)
$btnPostAll.Location = New-Object System.Drawing.Point(10, $y)
$btnPostAll.FlatStyle = "Flat"
$btnPostAll.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#107C10")
$btnPostAll.ForeColor = [System.Drawing.Color]::White
$btnPostAll.FlatAppearance.BorderSize = 0
$btnPostAll.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnPostAll)

# --- Log Box ---
$y += 55
$script:LogBox = New-Object System.Windows.Forms.TextBox
$script:LogBox.Multiline = $true
$script:LogBox.ScrollBars = "Vertical"
$script:LogBox.ReadOnly = $true
$script:LogBox.WordWrap = $false
$script:LogBox.Location = New-Object System.Drawing.Point(10, $y)
$script:LogBox.Size = New-Object System.Drawing.Size(880, 130)
$script:LogBox.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#1E1E1E")
$script:LogBox.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#CCCCCC")
$script:LogBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($script:LogBox)

# Track generated script path
$script:GeneratedScriptPath = $null
$script:SelectedRecipe = $null
$script:RecipePathByDisplay = @{}

function Refresh-RecipeList {
    param([string]$PreferredPath = "")

    $previousPath = $PreferredPath
    if (-not $previousPath -and $script:SelectedRecipe) {
        $previousPath = [string]$script:SelectedRecipe["_SourcePath"]
    }

    $cbRecipe.BeginUpdate()
    try {
        $cbRecipe.Items.Clear()
        $script:RecipePathByDisplay = @{}
        [void]$cbRecipe.Items.Add("None - Basic Package")

        $libraryPath = $tbRecipePath.Text.Trim()
        if (-not $libraryPath) {
            $cbRecipe.SelectedIndex = 0
            $lblRecipeStatus.Text = "Recipe library unavailable - basic package mode"
            $lblRecipeStatus.ForeColor = [System.Drawing.Color]::Orange
            LogLine "Recipe library not found: $libraryPath"
            return
        }

        $selectedDisplay = $null
        try {
            # Keep startup fast on network shares: one directory listing only.
            # Recipe contents are parsed, validated, and hashed lazily when the
            # user selects a recipe and again immediately before generation.
            $recipeFiles = @(
                Get-ChildItem -LiteralPath $libraryPath -Filter "*.psd1" -File -ErrorAction Stop |
                    Sort-Object Name
            )
        } catch {
            $cbRecipe.SelectedIndex = 0
            $lblRecipeStatus.Text = "Recipe library unavailable - basic package mode"
            $lblRecipeStatus.ForeColor = [System.Drawing.Color]::Orange
            LogLine ("Recipe library could not be listed: {0}" -f $_.Exception.Message)
            return
        }

        foreach ($file in $recipeFiles) {
            $display = $file.BaseName
            $script:RecipePathByDisplay[$display] = $file.FullName
            [void]$cbRecipe.Items.Add($display)
            if ($previousPath -and $file.FullName -eq $previousPath) {
                $selectedDisplay = $display
            }
        }

        if ($selectedDisplay) {
            $cbRecipe.SelectedItem = $selectedDisplay
        } else {
            $cbRecipe.SelectedIndex = 0
            $lblRecipeStatus.Text = if ($recipeFiles.Count) {
                "Basic package - $($recipeFiles.Count) recipe(s) available"
            } else {
                "Basic package - no recipes found"
            }
            $lblRecipeStatus.ForeColor = [System.Drawing.Color]::Silver
        }
        LogLine "Discovered $($recipeFiles.Count) PSADT recipe(s) from $libraryPath"
    } finally {
        $cbRecipe.EndUpdate()
    }
}

function Get-SelectedPsadtRecipe {
    $selectedDisplay = [string]$cbRecipe.SelectedItem
    if (-not $selectedDisplay -or -not $script:RecipePathByDisplay.ContainsKey($selectedDisplay)) {
        return $null
    }
    return Import-PsadtRecipe -Path $script:RecipePathByDisplay[$selectedDisplay]
}

function Show-PsadtRecipePreview {
    param([System.Collections.IDictionary]$Recipe)

    if (-not $Recipe) { return }

    $version = $tbPkgVersion.Text.Trim()
    $sections = Resolve-PsadtRecipeSections `
        -Recipe $Recipe `
        -Vendor $tbVendor.Text.Trim() `
        -AppName $tbAppName.Text.Trim() `
        -AppVersion $version `
        -InstallerType $script:InstallerType `
        -InstallerFile $script:InstallerFile

    $previewParts = @(
        "Recipe: $($Recipe.Name)$(if ($Recipe.RecipeVersion) { " v$($Recipe.RecipeVersion)" })",
        "Source: $($Recipe["_SourcePath"])",
        "SHA-256: $($Recipe["_Sha256"])",
        "Fixlet icon: $(if ($Recipe["_ResolvedFixletIconPath"]) { $Recipe["_ResolvedFixletIconPath"] } else { "(not specified)" })",
        "Install/Uninstall: nonblank recipe sections replace automatic commands",
        ""
    )
    foreach ($sectionName in $RecipeSectionNames) {
        $content = if ([string]::IsNullOrWhiteSpace($sections[$sectionName])) {
            "(blank - nothing added)"
        } else {
            $sections[$sectionName]
        }
        $previewParts += "===== $sectionName ====="
        $previewParts += $content
        $previewParts += ""
    }

    $previewForm = New-Object System.Windows.Forms.Form
    $previewForm.Text = "PSADT Recipe Preview - $($Recipe.Name)"
    $previewForm.Size = New-Object System.Drawing.Size(900, 700)
    $previewForm.StartPosition = "CenterParent"
    $previewForm.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#2D2D30")
    $previewForm.ForeColor = [System.Drawing.Color]::White
    $previewForm.Font = New-Object System.Drawing.Font("Consolas", 10)

    $previewBox = New-Object System.Windows.Forms.TextBox
    $previewBox.Multiline = $true
    $previewBox.ReadOnly = $true
    $previewBox.ScrollBars = "Both"
    $previewBox.WordWrap = $false
    $previewBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $previewBox.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#1E1E1E")
    $previewBox.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#D4D4D4")
    $previewBox.Text = $previewParts -join "`r`n"
    $previewForm.Controls.Add($previewBox)

    [void]$previewForm.ShowDialog($form)
}

# =========================
# EVENT HANDLERS
# =========================

# Helper function to load PSADT folder and populate fields
function Load-PsadtFolder {
    param([string]$FolderPath)
    $FolderPath = $FolderPath.Trim().TrimEnd('\')
    if (-not $FolderPath) { return }
    if (-not (Test-Path $FolderPath)) {
        LogLine "Path not found: $FolderPath"
        [System.Windows.Forms.MessageBox]::Show("Folder not found:`n$FolderPath", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    $tbPsadtPath.Text = $FolderPath
    LogLine "Selected folder: $FolderPath"
    
    try {
        $info = Parse-PsadtFolder -FolderPath $FolderPath
        if ($info.AppName) { $tbAppName.Text = $info.AppName }
        if ($info.Version) { $tbPkgVersion.Text = $info.Version }
        
        $script:InstallerType = $info.InstallerType
        $script:InstallerFile = $info.InstallerFile
        LogLine ("Parsed: App={0}, Ver={1}, PSADT Exe={2}, Installer={3} ({4})" -f $info.AppName, $info.Version, $info.PsadtExeName, $info.InstallerType, $info.InstallerFile)
        
        # Icon is not picked from the PSADT folder. Keep a manual or
        # recipe-provided self-service icon if one is already selected.
        if (-not $script:SelectedIconPath) {
            $lblIconStatus.Text = "[!] No icon selected - use Browse or choose a recipe with FixletIconPath"
            $lblIconStatus.ForeColor = [System.Drawing.Color]::Orange
            LogLine "No Fixlet icon selected (choose one manually or through a recipe)"
        }
    } catch {
        LogLine ("Could not auto-parse: {0}" -f $_.Exception.Message)
    }
}

# Browse PSADT folder (uses OpenFileDialog trick for modern explorer with address bar)
$btnBrowse.Add_Click({
    # Use OpenFileDialog with a trick to select folders — shows modern explorer with address bar
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Select PSADT Package Folder (select any file inside, or type path)"
    $dlg.Filter = "Folder Selection|*.---"
    $dlg.CheckFileExists = $false
    $dlg.CheckPathExists = $true
    $dlg.FileName = "Select Folder"
    $dlg.ValidateNames = $false
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $path = [System.IO.Path]::GetDirectoryName($dlg.FileName)
        if ($path) {
            Load-PsadtFolder -FolderPath $path
        }
    }
})

# Browse for a recipe library. UNC paths can also be pasted directly into the
# Recipe Library field and loaded with Refresh.
$btnBrowseRecipes.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Select the trusted PSADT recipe library"
    $dlg.ShowNewFolderButton = $false
    if ((Test-Path -LiteralPath $tbRecipePath.Text -PathType Container)) {
        $dlg.SelectedPath = $tbRecipePath.Text
    }
    if ($dlg.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $tbRecipePath.Text = $dlg.SelectedPath
        Refresh-RecipeList
    }
})

$btnRefreshRecipes.Add_Click({
    Refresh-RecipeList
})

$tbRecipePath.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        Refresh-RecipeList
        $_.SuppressKeyPress = $true
    }
})

$cbRecipe.Add_SelectedIndexChanged({
    try {
        $script:SelectedRecipe = Get-SelectedPsadtRecipe
        if ($script:SelectedRecipe) {
            $description = [string]$script:SelectedRecipe.Description
            $lblRecipeStatus.Text = if ($description) {
                $description
            } else {
                "Selected: $($script:SelectedRecipe.Name)"
            }
            $lblRecipeStatus.ForeColor = [System.Drawing.Color]::LightGreen
            $btnPreviewRecipe.Enabled = $true
            $btnPreviewRecipe.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#0078D7")
            LogLine "Selected PSADT recipe: $($script:SelectedRecipe.Name)"
            if ($script:SelectedRecipe["_ResolvedFixletIconPath"]) {
                Set-SelectedFixletIcon `
                    -Path $script:SelectedRecipe["_ResolvedFixletIconPath"] `
                    -Source "recipe '$($script:SelectedRecipe.Name)'" `
                    -FromRecipe
            } else {
                Clear-RecipeFixletIcon
            }
        } else {
            Clear-RecipeFixletIcon
            $lblRecipeStatus.Text = "Basic package - no recipe selected"
            $lblRecipeStatus.ForeColor = [System.Drawing.Color]::Silver
            $btnPreviewRecipe.Enabled = $false
            $btnPreviewRecipe.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#555555")
        }
    } catch {
        Clear-RecipeFixletIcon
        $script:SelectedRecipe = $null
        $lblRecipeStatus.Text = "Recipe error: $($_.Exception.Message)"
        $lblRecipeStatus.ForeColor = [System.Drawing.Color]::Red
        $btnPreviewRecipe.Enabled = $false
        LogLine ("Recipe selection failed: {0}" -f $_.Exception.Message)
    }
})

$btnPreviewRecipe.Add_Click({
    try {
        # Selection already loaded and validated this recipe. Reuse it for the
        # preview instead of reading and hashing the network file a second time.
        $recipe = $script:SelectedRecipe
        if (-not $recipe) { $recipe = Get-SelectedPsadtRecipe }
        if ($recipe) { Show-PsadtRecipePreview -Recipe $recipe }
    } catch {
        LogLine ("Recipe preview failed: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show(
            $form,
            $_.Exception.Message,
            "Recipe Preview Error",
            "OK",
            "Error"
        ) | Out-Null
    }
})

# Allow pressing Enter in the path textbox to load
$tbPsadtPath.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        Load-PsadtFolder -FolderPath $tbPsadtPath.Text
        $_.SuppressKeyPress = $true
    }
})

# Generate PSADT Script + Open ISE
$btnGenScript.Add_Click({
    $vendor  = $tbVendor.Text.Trim()
    $appName = $tbAppName.Text.Trim()
    $version = $tbPkgVersion.Text.Trim()
    $author  = $tbAuthor.Text.Trim()
    
    if (-not $appName) {
        [System.Windows.Forms.MessageBox]::Show("Enter Application Name first.", "Missing Info", "OK", "Warning") | Out-Null
        return
    }
    
    # Parse processes (comma-separated)
    $procText = $tbProcesses.Text.Trim()
    if ($procText -match '^\((auto|optional)') { $procText = "" }
    $procs = if ($procText) { 
        ($procText -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ } 
    } else { @() }
    
    $shortcut = $tbShortcut.Text.Trim()
    
    # Determine output path
    $psadtFolder = $tbPsadtPath.Text
    if ([string]::IsNullOrWhiteSpace($psadtFolder) -or $psadtFolder -eq "(no folder selected)") {
        $psadtFolder = $env:TEMP
    }
    
    # Output as Invoke-AppDeployToolkit.ps1 in the PSADT folder
    $outputPath = Join-Path $psadtFolder "Invoke-AppDeployToolkit.ps1"
    
    try {
        # Reload the selected recipe at generation time so changes made on the
        # server share after the dropdown was refreshed are validated and used.
        $selectedRecipe = Get-SelectedPsadtRecipe

        Generate-DeployScript `
            -Vendor $vendor `
            -AppName $appName `
            -AppVersion $version `
            -ProcessesToKill $procs `
            -ShortcutName $shortcut `
            -Author $author `
            -OutputPath $outputPath `
            -InstallerType $script:InstallerType `
            -InstallerFile $script:InstallerFile `
            -Recipe $selectedRecipe
        
        $script:GeneratedScriptPath = $outputPath
        LogLine "Generated: $outputPath"

        if ($selectedRecipe) {
            $snapshotPath = Save-PsadtRecipeSnapshot `
                -Recipe $selectedRecipe `
                -PackageFolder $psadtFolder
            LogLine "Recipe snapshot saved: $snapshotPath"
        }
        
        # Rename .ps1 and .exe to VendorAppName-Version convention
        $baseName = ("{0}{1}-{2}" -f $vendor, $appName, $version) -replace '\s+', ''
        $psadtExeName = "$baseName.exe"
        if ($baseName) {
            
            # Rename the generated .ps1
            $newPs1Path = Join-Path $psadtFolder "$baseName.ps1"
            if ($outputPath -ne $newPs1Path) {
                if (Test-Path $newPs1Path) { Remove-Item $newPs1Path -Force }
                Rename-Item -Path $outputPath -NewName "$baseName.ps1" -Force
                $script:GeneratedScriptPath = $newPs1Path
                LogLine "Renamed script: $baseName.ps1"
            }
            
            # Rename the .exe (Invoke-AppDeployToolkit.exe -> VendorApp-Version.exe)
            $oldExePath = Join-Path $psadtFolder "Invoke-AppDeployToolkit.exe"
            $newExePath = Join-Path $psadtFolder $psadtExeName
            if ((Test-Path $oldExePath) -and $oldExePath -ne $newExePath) {
                if (Test-Path $newExePath) { Remove-Item $newExePath -Force }
                Rename-Item -Path $oldExePath -NewName $psadtExeName -Force
                LogLine "Renamed exe: $psadtExeName"
            }
        }
        
        LogLine "Opening in PowerShell ISE for review - close ISE when done to continue..."
        
        # Open in ISE and wait (use GeneratedScriptPath which reflects rename)
        $editPath = $script:GeneratedScriptPath
        $form.Enabled = $false
        try {
            Start-Process "powershell_ise.exe" -ArgumentList "`"$editPath`"" -Wait
        } catch {
            LogLine ("Could not open ISE: {0}. Trying VS Code..." -f $_.Exception.Message)
            try { Start-Process "code" -ArgumentList "`"$editPath`"" -Wait } catch {
                LogLine "Could not open editor. Please review the file manually: $editPath"
            }
        }
        $form.Enabled = $true
        
        LogLine "ISE closed. Script ready for signing."
        $lblSignStatus.Text = "Ready to sign"
        $lblSignStatus.ForeColor = [System.Drawing.Color]::Yellow
        
    } catch {
        LogLine ("Failed to generate script: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show(
            $form,
            $_.Exception.Message,
            "PSADT Script Generation Failed",
            "OK",
            "Error"
        ) | Out-Null
    }
})

# Sign Script
$btnSignScript.Add_Click({
    if (-not $script:GeneratedScriptPath -or -not (Test-Path $script:GeneratedScriptPath)) {
        [System.Windows.Forms.MessageBox]::Show("Generate the PSADT script first.", "No Script", "OK", "Warning") | Out-Null
        return
    }
    
    try {
        if ($SigningCertThumbprint -eq "YOUR-CERT-THUMBPRINT-HERE") {
            LogLine "[!] Signing cert thumbprint not configured - update `$SigningCertThumbprint in config section"
            $lblSignStatus.Text = "[!] No cert configured"
            $lblSignStatus.ForeColor = [System.Drawing.Color]::Orange
            return
        }
        
        $cert = Get-ChildItem -Path "Cert:\CurrentUser\My\$SigningCertThumbprint" -ErrorAction SilentlyContinue
        if (-not $cert) {
            $cert = Get-ChildItem -Path "Cert:\LocalMachine\My\$SigningCertThumbprint" -ErrorAction SilentlyContinue
        }
        
        if (-not $cert) {
            LogLine "[X] Certificate not found with thumbprint: $SigningCertThumbprint"
            $lblSignStatus.Text = "[X] Cert not found"
            $lblSignStatus.ForeColor = [System.Drawing.Color]::Red
            return
        }
        
        $result = Set-AuthenticodeSignature -FilePath $script:GeneratedScriptPath -Certificate $cert -TimestampServer "http://timestamp.digicert.com"
        
        if ($result.Status -eq "Valid") {
            LogLine "[OK] Script signed successfully"
            $lblSignStatus.Text = "[OK] Signed"
            $lblSignStatus.ForeColor = [System.Drawing.Color]::LightGreen
        } else {
            LogLine ("[!] Signing result: {0}" -f $result.StatusMessage)
            $lblSignStatus.Text = "[!] $($result.Status)"
            $lblSignStatus.ForeColor = [System.Drawing.Color]::Orange
        }
    } catch {
        LogLine ("[X] Signing failed: {0}" -f $_.Exception.Message)
        $lblSignStatus.Text = "[X] Failed"
        $lblSignStatus.ForeColor = [System.Drawing.Color]::Red
    }
})

# POST Fixlets + Create QA Offers
$btnPostAll.Add_Click({
    $script:LogBox.Clear()
    LogLine "== Starting Unified Package Pipeline =="
    
    # Validate
    $vendor    = $tbVendor.Text.Trim()
    $appName   = $tbAppName.Text.Trim()
    $version   = $tbPkgVersion.Text.Trim()
    $psadtExe  = ("{0}{1}-{2}.exe" -f $vendor, $appName, $version) -replace '\s+', ''
    $prefetch  = $tbPrefetch.Text.Trim()
    $extract   = ""  # combined into prefetch field
    $server    = $cbServer.SelectedItem
    $site      = $cbSite.SelectedItem
    
    if (-not ($appName -and $version -and $prefetch -and $server -and $site)) {
        LogLine "[X] Fill in all required fields (App Name, Version, Prefetch, Server, Site)"
        return
    }
    
    if (-not ($tbInstallRel.Text.Trim() -and $tbUpdateRel.Text.Trim() -and $tbRemoveRel.Text.Trim())) {
        LogLine "[X] Enter Install, Update, and Remove relevance before posting"
        return
    }
    
    $displayName = if ($vendor) { "$vendor $appName" } else { $appName }
    
    # Confirm
    $msg = "Ready to create 3 fixlets + 3 QA offers for:`n`n$displayName v$version`nSite: $site`nServer: $server`n`nYou will be prompted for credentials twice:`n1. Fixlet creation credentials`n2. Offer/Action creation credentials`n`nProceed?"
    $dlg = [System.Windows.Forms.MessageBox]::Show($form, $msg, "Confirm", "YesNo", "Question", "Button2")
    if ($dlg -ne [System.Windows.Forms.DialogResult]::Yes) {
        LogLine "Cancelled."
        return
    }
    
    # --- Credential Prompt #1: Fixlet Creation ---
    LogLine "Requesting fixlet creation credentials..."
    $fixletCreds = Show-CredentialDialog -Title "Fixlet Creation Credentials" -Message "Enter credentials for fixlet creation:" -Owner $form
    if (-not $fixletCreds) {
        LogLine "Cancelled - no fixlet credentials provided."
        return
    }
    LogLine ("Fixlet creds: user={0}" -f $fixletCreds.User)
    
    try {
        $base = Get-BaseUrl $server
        $encodedSite = Encode-SiteName $site
        $fixletPostUrl = Join-ApiUrl -BaseUrl $base -RelativePath "/api/fixlets/custom/$encodedSite"
        $actionPostUrl = Join-ApiUrl -BaseUrl $base -RelativePath "/api/actions"
        
        LogLine "Fixlet POST URL: $fixletPostUrl"
        
        # Build action scripts (prefetch content pasted by user, no extra "prefetch" keyword)
        $installAS = "$prefetch`r`naction uses wow64 redirection false`r`nwait __Download\$psadtExe -DeploymentType Install -DeployMode Silent"
        $updateAS  = "$prefetch`r`naction uses wow64 redirection false`r`nwait __Download\$psadtExe -DeploymentType Install -DeployMode Silent"
        $removeAS  = "$prefetch`r`naction uses wow64 redirection false`r`nwait __Download\$psadtExe -DeploymentType Uninstall -DeployMode Silent"
        
        # Get icon base64 if available
        $iconDataUri = ""
        LogLine ("DEBUG Icon path: '{0}'" -f $script:SelectedIconPath)
        if ($script:SelectedIconPath -and (Test-Path $script:SelectedIconPath)) {
            $iconInfo = Get-IconBase64 -IconPath $script:SelectedIconPath
            if ($iconInfo) {
                $iconDataUri = $iconInfo.DataUri
                LogLine ("Icon for fixlet: {0} ({1}, {2} chars)" -f $script:SelectedIconPath, $iconInfo.MimeType, $iconDataUri.Length)
            } else {
                LogLine "[!] Icon file found but failed to encode to base64"
            }
        } else {
            LogLine "[!] No icon selected - fixlets will have no self-service icon"
        }
        
        # Build and post fixlets
        $fixlets = @(
            @{ Kind="Install"; Title="Install: $displayName $version Win"; Description="This fixlet will install $displayName $version."; Category="Install"; Source="Install"; Relevance=$tbInstallRel.Text; ActionScript=$installAS },
            @{ Kind="Update";  Title="Update: $displayName $version Win";  Description="This fixlet will update $displayName to version $version."; Category="Pending"; Source="Pending"; Relevance=$tbUpdateRel.Text;  ActionScript=$updateAS  },
            @{ Kind="Remove";  Title="Remove: $displayName $version Win";  Description="This fixlet will remove $displayName $version."; Category="Remove"; Source="Remove"; Relevance=$tbRemoveRel.Text;  ActionScript=$removeAS  }
        )
        
        $fixletIds = @()
        
        foreach ($fx in $fixlets) {
            LogLine ("Creating fixlet: {0}" -f $fx.Title)
            $xml = Build-FixletXml -Title $fx.Title -Description $fx.Description -Relevance $fx.Relevance -ActionScript $fx.ActionScript -Category $fx.Category -Source $fx.Source -IconBase64DataUri $iconDataUri
            
            $posted = $false
            $retries = 0
            while (-not $posted -and $retries -lt 3) {
                try {
                    $resp = Post-Xml -Url $fixletPostUrl -User $fixletCreds.User -Pass $fixletCreds.Pass -XmlBody $xml
                    $posted = $true
                } catch {
                    $errMsg = $_.Exception.Message
                    $statusMatch = [regex]::Match($errMsg, '\((\d{3})\)')
                    $httpCode = if ($statusMatch.Success) { $statusMatch.Groups[1].Value } else { "unknown" }
                    LogLine ("[ERROR] POST failed: HTTP $httpCode - $errMsg")
                    
                    if ($httpCode -eq "401" -or $errMsg -match "401|Unauthorized") {
                        LogLine "[!] Authentication failed - please re-enter credentials"
                        $fixletCreds = Show-CredentialDialog -Title "Fixlet Credentials (Retry)" -Message "Authentication failed (401). Please re-enter your credentials:" -Owner $form
                        if (-not $fixletCreds) {
                            LogLine "[ABORT] No credentials provided. Pipeline stopped."
                            return
                        }
                        $retries++
                    } elseif ($httpCode -eq "400") {
                        LogLine "[ERROR] Bad Request (400) - possible malformed XML. Check log for details."
                        LogLine ("XML length: {0} chars" -f $xml.Length)
                        return
                    } else {
                        LogLine "[ERROR] Unexpected error - pipeline stopped."
                        return
                    }
                }
            }
            if (-not $posted) {
                LogLine "[ABORT] Failed after 3 credential attempts. Pipeline stopped."
                return
            }
            
            # Parse fixlet ID from response
            $fixletId = $null
            if ($resp.Content -match 'ID>(\d+)<') { $fixletId = $Matches[1] }
            elseif ($resp.Content -match '"id"\s*:\s*(\d+)') { $fixletId = $Matches[1] }
            elseif ($resp.Headers -and $resp.Headers["Location"] -and $resp.Headers["Location"] -match '/(\d+)$') { $fixletId = $Matches[1] }
            
            if ($fixletId) {
                $fixletIds += $fixletId
                LogLine ("[OK] {0} fixlet created: ID {1}" -f $fx.Kind, $fixletId)
            } else {
                LogLine ("[!] {0} fixlet posted but couldn't parse ID" -f $fx.Kind)
                $fixletIds += "UNKNOWN"
            }
        }
        
        LogLine ("Fixlet IDs: Install={0}, Update={1}, Remove={2}" -f $fixletIds[0], $fixletIds[1], $fixletIds[2])
        
        # --- Credential Prompt #2: Offer/Action Creation ---
        LogLine "Requesting offer/action creation credentials..."
        $offerCreds = Show-CredentialDialog -Title "Offer/Action Creation Credentials" -Message "Enter credentials for offer/action creation:" -Owner $form -StayOnTop
        if (-not $offerCreds) {
            LogLine "Cancelled - no offer credentials. Fixlets were created successfully."
            LogLine ("Fixlet IDs for manual offer creation: {0}, {1}, {2}" -f $fixletIds[0], $fixletIds[1], $fixletIds[2])
            return
        }
        LogLine ("Offer creds: user={0}" -f $offerCreds.User)
        
        # --- Create QA Offers ---
        LogLine "== Creating QA Offers =="
        
        $groupNum = Get-NumericGroupId $QA_GroupIdWithPrefix
        $groupRel = Build-GroupMembershipRelevance -SiteName $site -GroupIdNumeric $groupNum
        LogLine ("QA group relevance: {0}" -f $groupRel)
        
        $kinds = @("Install","Update","Remove")
        for ($i = 0; $i -lt 3; $i++) {
            if ($fixletIds[$i] -eq "UNKNOWN") {
                LogLine ("[!] Skipping {0} offer - no fixlet ID" -f $kinds[$i])
                continue
            }
            
            $offerXml = Build-OfferXml `
                -DisplayName "$displayName $version" `
                -SiteName $site `
                -FixletId $fixletIds[$i] `
                -FixletActionName $FixletActionName_Default `
                -GroupRelevance $groupRel `
                -Kind $kinds[$i] `
                -Phase "QA" `
                -OfferDescription "" `
                -IconBase64DataUri $iconDataUri
            
            $offerPosted = $false
            $offerRetries = 0
            while (-not $offerPosted -and $offerRetries -lt 3) {
                try {
                    Post-Xml -Url $actionPostUrl -User $offerCreds.User -Pass $offerCreds.Pass -XmlBody $offerXml | Out-Null
                    LogLine ("[OK] QA Offer created: {0}" -f $kinds[$i])
                    $offerPosted = $true
                } catch {
                    $errMsg = $_.Exception.Message
                    LogLine ("[ERROR] QA Offer failed for {0}: {1}" -f $kinds[$i], $errMsg)
                    if ($errMsg -match "401|Unauthorized") {
                        LogLine "[!] Authentication failed for offers - please re-enter credentials"
                        $offerCreds = Show-CredentialDialog -Title "Offer Credentials (Retry)" -Message "Authentication failed (401). Please re-enter your offer credentials:" -Owner $form -StayOnTop
                        if (-not $offerCreds) {
                            LogLine "[ABORT] No credentials provided. Remaining offers skipped."
                            break
                        }
                        $offerRetries++
                    } else {
                        LogLine "[ERROR] Non-auth error - skipping this offer."
                        break
                    }
                }
            }
        }
        
        LogLine "========================================="
        LogLine "[OK] Pipeline Complete!"
        LogLine ("  App: {0} v{1}" -f $displayName, $version)
        LogLine ("  Fixlets: {0}, {1}, {2}" -f $fixletIds[0], $fixletIds[1], $fixletIds[2])
        LogLine "  QA Offers: Created"
        LogLine "========================================="
        LogLine "Log saved to: $LogFile"
        LogLine "Click 'Create Deployment Doc' to generate PDF/Word documentation."
        
        # Enable deployment doc button
        $script:PipelineFixletIds = $fixletIds
        $script:PipelineComplete = $true
        
    } catch {
        LogLine ("[X] Fatal error: {0}" -f ($_.Exception.GetBaseException().Message))
    }
})

# =========================
# DEPLOYMENT DOC GENERATOR
# =========================
function Generate-DeploymentDocHtml {
    param(
        [string]$Vendor,
        [string]$AppName,
        [string]$Version,
        [string]$Author,
        [string]$PsadtFolder,
        [string]$Server,
        [string]$Site,
        [string]$IconDataUri,
        [hashtable[]]$Fixlets,
        [string[]]$FixletIds,
        [string]$InstallRelevance,
        [string]$UpdateRelevance,
        [string]$RemoveRelevance,
        [string]$InstallActionScript,
        [string]$UpdateActionScript,
        [string]$RemoveActionScript,
        [string]$KbLink
    )
    
    $displayName = if ($Vendor) { "$Vendor $AppName" } else { $AppName }
    $dateGenerated = Get-Date -Format "MMMM dd, yyyy 'at' hh:mm tt"
    $iconImg = if ($IconDataUri) { "<img src='$IconDataUri' style='width:48px;height:48px;vertical-align:middle;margin-right:12px;'/>" } else { "" }
    
    $installId = if ($FixletIds.Count -ge 1) { $FixletIds[0] } else { "N/A" }
    $updateId  = if ($FixletIds.Count -ge 2) { $FixletIds[1] } else { "N/A" }
    $removeId  = if ($FixletIds.Count -ge 3) { $FixletIds[2] } else { "N/A" }
    
    $escInstallRel = [System.Web.HttpUtility]::HtmlEncode($InstallRelevance)
    $escUpdateRel  = [System.Web.HttpUtility]::HtmlEncode($UpdateRelevance)
    $escRemoveRel  = [System.Web.HttpUtility]::HtmlEncode($RemoveRelevance)
    $escInstallAS  = [System.Web.HttpUtility]::HtmlEncode($InstallActionScript)
    $escUpdateAS   = [System.Web.HttpUtility]::HtmlEncode($UpdateActionScript)
    $escRemoveAS   = [System.Web.HttpUtility]::HtmlEncode($RemoveActionScript)
    
    return @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<title>Deployment Document - $([System.Web.HttpUtility]::HtmlEncode($displayName)) v$([System.Web.HttpUtility]::HtmlEncode($Version))</title>
<style>
  @page { margin: 0.75in; }
  * { box-sizing: border-box; }
  body { font-family: 'Segoe UI', Calibri, Arial, sans-serif; color: #1a1a1a; margin: 0; padding: 40px; background: #fff; line-height: 1.5; }
  .header { background: linear-gradient(135deg, #0078D4 0%, #004E8C 100%); color: #fff; padding: 30px 35px; border-radius: 8px; margin-bottom: 30px; display: flex; align-items: center; }
  .header-icon { margin-right: 20px; }
  .header-icon img { width: 64px; height: 64px; border-radius: 8px; background: rgba(255,255,255,0.15); padding: 4px; }
  .header-text h1 { margin: 0 0 4px 0; font-size: 24px; font-weight: 600; }
  .header-text .subtitle { opacity: 0.85; font-size: 13px; }
  .badge { display: inline-block; background: rgba(255,255,255,0.2); padding: 3px 10px; border-radius: 12px; font-size: 11px; margin-top: 6px; }
  .toc { background: #f8f9fa; border: 1px solid #e0e0e0; border-radius: 6px; padding: 18px 24px; margin-bottom: 28px; }
  .toc h3 { margin: 0 0 10px 0; color: #0078D4; font-size: 14px; text-transform: uppercase; letter-spacing: 0.5px; }
  .toc ul { list-style: none; padding: 0; margin: 0; }
  .toc li { padding: 4px 0; }
  .toc a { color: #0078D4; text-decoration: none; font-size: 13px; }
  .toc a:hover { text-decoration: underline; }
  .section { margin-bottom: 28px; page-break-inside: avoid; }
  .section h2 { font-size: 16px; color: #0078D4; border-bottom: 2px solid #0078D4; padding-bottom: 6px; margin-bottom: 14px; }
  table { width: 100%; border-collapse: collapse; margin-bottom: 10px; font-size: 13px; }
  th { background: #0078D4; color: #fff; text-align: left; padding: 9px 12px; font-weight: 600; }
  td { padding: 8px 12px; border-bottom: 1px solid #e8e8e8; }
  tr:nth-child(even) td { background: #f8f9fa; }
  .code-block { background: #1e1e1e; color: #d4d4d4; padding: 14px 16px; border-radius: 5px; font-family: 'Cascadia Code', 'Consolas', monospace; font-size: 12px; white-space: pre-wrap; word-break: break-all; overflow-x: auto; margin: 8px 0; }
  .fixlet-card { border: 1px solid #e0e0e0; border-radius: 6px; padding: 16px 20px; margin-bottom: 16px; background: #fafbfc; }
  .fixlet-card h3 { margin: 0 0 10px 0; font-size: 14px; }
  .fixlet-card .fixlet-id { display: inline-block; background: #0078D4; color: #fff; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; margin-left: 8px; }
  .install-card h3 { color: #107C10; }
  .update-card h3 { color: #CA5010; }
  .remove-card h3 { color: #D13438; }
  .label { font-weight: 600; color: #555; font-size: 12px; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 4px; }
  .footer { margin-top: 40px; padding-top: 16px; border-top: 1px solid #e0e0e0; font-size: 11px; color: #888; text-align: center; }
  .summary-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 10px; }
  .summary-item { background: #f0f6ff; border-radius: 5px; padding: 10px 14px; }
  .summary-item .val { font-size: 15px; font-weight: 600; color: #1a1a1a; }
  .summary-item .lbl { font-size: 11px; color: #666; text-transform: uppercase; }
  @media print {
    body { padding: 0; }
    .header { border-radius: 0; }
    .no-print { display: none !important; }
  }
</style>
</head>
<body>

<div class="header">
  $(if ($IconDataUri) { "<div class='header-icon'><img src='$IconDataUri'/></div>" })
  <div class="header-text">
    <h1>$([System.Web.HttpUtility]::HtmlEncode($displayName))</h1>
    <div class="subtitle">Application Deployment Document &mdash; Version $([System.Web.HttpUtility]::HtmlEncode($Version))</div>
    <div class="badge">Generated $dateGenerated</div>
  </div>
</div>

<div class="toc">
  <h3>Contents</h3>
  <ul>
    <li><a href="#overview">1. Package Overview</a></li>
    <li><a href="#fixlets">2. Fixlet Details</a></li>
    <li><a href="#offers">3. QA Offers</a></li>
    <li><a href="#deployment">4. Deployment Notes</a></li>
  </ul>
</div>

<div class="section" id="overview">
  <h2>1. Package Overview</h2>
  <div class="summary-grid">
    <div class="summary-item"><div class="lbl">Vendor</div><div class="val">$([System.Web.HttpUtility]::HtmlEncode($Vendor))</div></div>
    <div class="summary-item"><div class="lbl">Application</div><div class="val">$([System.Web.HttpUtility]::HtmlEncode($AppName))</div></div>
    <div class="summary-item"><div class="lbl">Version</div><div class="val">$([System.Web.HttpUtility]::HtmlEncode($Version))</div></div>
    <div class="summary-item"><div class="lbl">Author</div><div class="val">$([System.Web.HttpUtility]::HtmlEncode($Author))</div></div>
  </div>
  <table>
    <tr><th style="width:200px;">Property</th><th>Value</th></tr>
    <tr><td>Package Location</td><td><code>$([System.Web.HttpUtility]::HtmlEncode($PsadtFolder))</code></td></tr>
    <tr><td>BigFix Server</td><td>$([System.Web.HttpUtility]::HtmlEncode($Server))</td></tr>
    <tr><td>BigFix Site</td><td>$([System.Web.HttpUtility]::HtmlEncode($Site))</td></tr>
    <tr><td>Date Created</td><td>$dateGenerated</td></tr>
    $(if ($KbLink) { "<tr><td>Packaging KB</td><td><a href='$([System.Web.HttpUtility]::HtmlEncode($KbLink))' style='color:#0078D4;'>$([System.Web.HttpUtility]::HtmlEncode($KbLink))</a></td></tr>" })
  </table>
</div>

<div class="section" id="fixlets">
  <h2>2. Fixlet Details</h2>
  
  <div class="fixlet-card install-card">
    <h3>Install Fixlet <span class="fixlet-id">ID: $installId</span></h3>
    <div class="label">Relevance</div>
    <div class="code-block">$escInstallRel</div>
    <div class="label">Action Script</div>
    <div class="code-block">$escInstallAS</div>
  </div>
  
  <div class="fixlet-card update-card">
    <h3>Update Fixlet <span class="fixlet-id">ID: $updateId</span></h3>
    <div class="label">Relevance</div>
    <div class="code-block">$escUpdateRel</div>
    <div class="label">Action Script</div>
    <div class="code-block">$escUpdateAS</div>
  </div>
  
  <div class="fixlet-card remove-card">
    <h3>Remove Fixlet <span class="fixlet-id">ID: $removeId</span></h3>
    <div class="label">Relevance</div>
    <div class="code-block">$escRemoveRel</div>
    <div class="label">Action Script</div>
    <div class="code-block">$escRemoveAS</div>
  </div>
</div>

<div class="section" id="offers">
  <h2>3. QA Offers</h2>
  <table>
    <tr><th>Offer Type</th><th>Source Fixlet ID</th><th>Target Group</th><th>Phase</th><th>Status</th></tr>
    <tr><td>Install</td><td>$installId</td><td>QA Group</td><td>QA</td><td style="color:#107C10;font-weight:600;">Created</td></tr>
    <tr><td>Update</td><td>$updateId</td><td>QA Group</td><td>QA</td><td style="color:#107C10;font-weight:600;">Created</td></tr>
    <tr><td>Remove</td><td>$removeId</td><td>QA Group</td><td>QA</td><td style="color:#107C10;font-weight:600;">Created</td></tr>
  </table>
</div>

<div class="section" id="deployment">
  <h2>4. Deployment Notes</h2>
  <table>
    <tr><th>Step</th><th>Description</th></tr>
    <tr><td>1</td><td>QA offers are live and targeted to the QA group. Verify install/update/remove on test machines.</td></tr>
    <tr><td>2</td><td>After QA approval, create Production offers targeting the appropriate production groups.</td></tr>
    <tr><td>3</td><td>Monitor deployment progress via BigFix Console &gt; Actions tab.</td></tr>
    <tr><td>4</td><td>Confirm successful deployment and close change ticket.</td></tr>
  </table>
</div>

<div class="footer">
  BigFix Unified Packager v$ToolVer &mdash; Auto-generated deployment document &mdash; $dateGenerated
</div>

</body>
</html>
"@
}

# --- Create Deployment Doc Button ---
$btnCreateDoc = New-Object System.Windows.Forms.Button
$btnCreateDoc.Text = "Create Deployment Doc"
$btnCreateDoc.Size = New-Object System.Drawing.Size(220, 40)
$btnCreateDoc.Location = New-Object System.Drawing.Point(340, $btnPostAll.Location.Y)
$btnCreateDoc.FlatStyle = "Flat"
$btnCreateDoc.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#0078D7")
$btnCreateDoc.ForeColor = [System.Drawing.Color]::White
$btnCreateDoc.FlatAppearance.BorderSize = 0
$btnCreateDoc.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$btnCreateDoc.Enabled = $false
$form.Controls.Add($btnCreateDoc)

# Track pipeline results for doc generation
$script:PipelineFixletIds = @()
$script:PipelineComplete = $false

$btnCreateDoc.Add_Click({
    $vendor   = $tbVendor.Text.Trim()
    $appName  = $tbAppName.Text.Trim()
    $version  = $tbPkgVersion.Text.Trim()
    $author   = $tbAuthor.Text.Trim()
    $psadtFolder = $tbPsadtPath.Text
    $server   = $cbServer.SelectedItem
    $site     = $cbSite.SelectedItem
    $displayName = if ($vendor) { "$vendor $appName" } else { $appName }
    $safeFileName = ($displayName -replace '[^\w\-\.]','_') + "_v" + ($version -replace '[^\w\.\-]','_')
    
    $psadtExe = ("{0}{1}-{2}.exe" -f $vendor, $appName, $version) -replace '\s+', ''
    $installAS = "$($tbPrefetch.Text.Trim())`r`naction uses wow64 redirection false`r`nwait __Download\$psadtExe -DeploymentType Install -DeployMode Silent"
    $updateAS  = "$($tbPrefetch.Text.Trim())`r`naction uses wow64 redirection false`r`nwait __Download\$psadtExe -DeploymentType Install -DeployMode Silent"
    $removeAS  = "$($tbPrefetch.Text.Trim())`r`naction uses wow64 redirection false`r`nwait __Download\$psadtExe -DeploymentType Uninstall -DeployMode Silent"
    
    $iconUri = ""
    if ($script:SelectedIconPath -and (Test-Path $script:SelectedIconPath)) {
        $iconInfo = Get-IconBase64 -IconPath $script:SelectedIconPath
        if ($iconInfo) { $iconUri = $iconInfo.DataUri }
    }
    
    $html = Generate-DeploymentDocHtml `
        -Vendor $vendor -AppName $appName -Version $version -Author $author `
        -PsadtFolder $psadtFolder -Server $server -Site $site `
        -IconDataUri $iconUri -FixletIds $script:PipelineFixletIds `
        -InstallRelevance $tbInstallRel.Text -UpdateRelevance $tbUpdateRel.Text -RemoveRelevance $tbRemoveRel.Text `
        -InstallActionScript $installAS -UpdateActionScript $updateAS -RemoveActionScript $removeAS `
        -KbLink $(if ($tbKbLink.Text -ne "(optional) ServiceNow KB article URL" -and -not [string]::IsNullOrWhiteSpace($tbKbLink.Text)) { $tbKbLink.Text } else { "" })
    
    # Save HTML directly — no format picker
    $saveDir = $env:TEMP
    if ($psadtFolder -ne "(no folder selected)" -and (Test-Path $psadtFolder)) {
        $parentFolder = Split-Path $psadtFolder -Parent
        if ($parentFolder -and (Test-Path $parentFolder)) {
            $saveDir = $parentFolder
        } else {
            $saveDir = $psadtFolder
        }
    }
    
    # If pipeline hasn't run, use placeholder IDs
    if (-not $script:PipelineFixletIds -or $script:PipelineFixletIds.Count -eq 0) {
        $script:PipelineFixletIds = @("(pending)", "(pending)", "(pending)")
    }
    
    $htmlPath = Join-Path $saveDir "$safeFileName`_DeploymentDoc.html"
    [System.IO.File]::WriteAllText($htmlPath, $html, [System.Text.Encoding]::UTF8)
    
    [System.Windows.Forms.MessageBox]::Show(
        "Deployment document saved:`n`n$([System.IO.Path]::GetFileName($htmlPath))`n`nLocation: $saveDir",
        "Document Created",
        "OK", "Information"
    ) | Out-Null
    
    # Open the folder
    Start-Process "explorer.exe" -ArgumentList "/select,`"$htmlPath`""
    
    LogLine "Deployment document generation complete."
})

# Enable the doc button after pipeline completes - hook into the end of btnPostAll
$btnPostAll.Tag = $btnPostAll.Add_Click  # original handler already registered above

# We need to enable btnCreateDoc after pipeline success.
# Patch: add a timer that checks for pipeline completion
$script:DocButtonTimer = New-Object System.Windows.Forms.Timer
$script:DocButtonTimer.Interval = 500
$script:DocButtonTimer.Add_Tick({
    if ($script:PipelineComplete) {
        $btnCreateDoc.Enabled = $true
        $btnCreateDoc.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#0078D7")
        $script:DocButtonTimer.Stop()
    }
})
$script:DocButtonTimer.Start()

# Load the configured local or server recipe library after all controls and
# event handlers are ready.
Refresh-RecipeList

$form.TopMost = $false
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
