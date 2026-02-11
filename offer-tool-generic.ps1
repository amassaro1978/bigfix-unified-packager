# =========================================================
# BigFix Offer Action Generator (Install / Update / Remove)
# Baseline: Offers-2025-09-24-v17-DirectGroupSitesPlural
# - Default targeting: (member of group <id> of sites)
# - Optional: site-specific clause or fetched group relevance
# - Includes: LLM-normalized descriptions, schema-safe XML,
#             detailed error logging, updated running messages,
#             confirmation dialog
# =========================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

# =========================
# CONFIG (hardcoded)
# =========================
$LogFile = "C:\temp\BigFixOfferGenerator.log"

# Fixlets' site (where the 3 fixlets live)
$CustomSiteName = "Test Group Managed (Workstations)"

# ---- GROUP TARGETING MODE ----
# Recommended: use direct membership relevance (no XML fetch)
$UseDirectGroupMembershipRelevance = $true

# If you turn OFF the direct mode (set to $false), the tool will fetch the group's relevance XML.
# When groups are in a DIFFERENT site than the fixlets, hardcode that site here:
$GroupSiteNameOverride = "Production Device Groups"  # used only if $UseDirectGroupMembershipRelevance -eq $false

# Site-agnostic vs site-specific membership clause
# true  -> (member of group <id> of sites)
# false -> (member of group <id> of site whose(name of it = "<site>"))
$UseSitesPlural = $true

# Fixlet Action name inside each Fixlet
$FixletActionName_Default = "Action1"

# Hardcoded group IDs for the two rounds
$QA_GroupIdWithPrefix     = "00-12345"
$Deploy_GroupIdWithPrefix = "00-67890"

# Offer settings
$OfferDefaults = @{
    PreActionShowUI = $false
    RetryCount      = 3
    RetryWaitISO    = "PT1H"
    StartOffsetISO  = "PT0S"        # starts now
    EndOffsetISO    = "P365DT0S"    # ends in 1 year
    Reapply         = $true
    ContinueOnErr   = $true
}

# ===== LLM CONFIG (env var) =====
$LLMConfig = @{
    EnableAuto = $true
    ApiUrl     = "https://redacted/v1/chat/completions"  # your LiteLLM endpoint
    ApiKeyEnv  = "LITELLM_KEY"
    Model      = "gpt-40"
}

# Behavior toggles
$IgnoreCertErrors           = $true
$DumpFetchedXmlToTemp       = $true
$AggressiveRegexFallback    = $true
$SaveActionXmlToTemp        = $true
$PostUsingInvokeWebRequest  = $true

# =========================
# UTIL / LOGGING
# =========================
function Encode-SiteName([string]$Name) {
    $enc = [System.Web.HttpUtility]::UrlEncode($Name, [System.Text.Encoding]::UTF8)
    $enc = $enc -replace '\+','%20' -replace '\(','%28' -replace '\)','%29'
    return $enc
}
function Get-BaseUrl([string]$ServerInput) {
    if (-not $ServerInput) { throw "Server is empty." }
    $s = $ServerInput.Trim()
    if ($s -notmatch '^(?i)https?://') {
        if ($s -match ':\d+$') { $s = "https://$s" } else { $s = "https://$s:52311" }
    }
    return $s.TrimEnd('/')
}
function Join-ApiUrl([string]$BaseUrl,[string]$RelativePath) {
    $rp = if ($RelativePath.StartsWith("/")) { $RelativePath } else { "/$RelativePath" }
    $BaseUrl.TrimEnd('/') + $rp
}
function Get-AuthHeader([string]$User,[string]$Pass) {
    $pair  = "$User`:$Pass"
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    "Basic " + [Convert]::ToBase64String($bytes)
}
function LogLine($txt) {
    try {
        $line = "{0}  {1}" -f (Get-Date -Format 'u'), $txt
        if ($LogBox) { $LogBox.AppendText($line + "`r`n"); $LogBox.SelectionStart = $LogBox.Text.Length; $LogBox.ScrollToCaret() }
        Add-Content -Path $LogFile -Value $line
    } catch {}
}
function Get-NumericGroupId([string]$GroupIdWithPrefix) {
    if ($GroupIdWithPrefix -match '^\d{2}-(\d+)$') { return $Matches[1] }
    return ($GroupIdWithPrefix -replace '[^\d]','')
}
function SafeEscape([string]$s) {
    if ($null -eq $s) { return "" }
    [System.Security.SecurityElement]::Escape($s)
}
function To-XmlBool([bool]$b) { if ($b) { 'true' } else { 'false' } }

# Clean product name from Fixlet title
function Parse-FixletTitleToProduct([string]$Title) {
    $t = [string]$Title
    $t = $t -replace '^\s*(?i)(Install|Update|Remove)\s*[:\-]\s*',''
    $t = $t -replace '^\s*(?i)Update:\s*',''
    $t = $t -replace '\s+Win$',''
    $t.Trim()
}

function Truncate([string]$s,[int]$max=600) {
    if ([string]::IsNullOrWhiteSpace($s)) { return "" }
    if ($s.Length -le $max) { return $s.Trim() }
    return ($s.Substring(0,$max)).Trim() + "…"
}

function Build-LLMPrompt([string]$displayName) {
@"
You are generating a short end-user description for a BigFix software offer.

Product: $displayName

Requirements:
- Output must BEGIN with: $displayName
- Then continue with 1–3 concise sentences describing what the user gets (features/benefits/notes).
- Do NOT say "This offer will", "This update will", "Offer:", or similar boilerplate.
- Do NOT lead with verbs like "Install/Update/Remove".
- Plain English, no HTML/Markdown, under ~90 words.
- Output ONLY the text. No headings or labels.
"@
}

function Normalize-OfferDescription {
    param([string]$text, [string]$displayName)

    if ([string]::IsNullOrWhiteSpace($text)) { return $null }

    $t = $text.Trim()

    # Remove common boilerplate openings like "This offer will install/update..."
    $t = $t -replace '^(?is)\s*(this\s+(offer|update)\s+will|it\s+will|the\s+offer\s+will)\b.*?(install|update|remove)\b[^.:!?]*[:\-]?\s*', ''

    # Remove leading "Install/Update/Remove:" labels if present
    $t = $t -replace '^(?is)\s*(install|update|remove)\s*[:\-]\s*', ''

    # If it still doesn't start with the product name, prefix it
    $dnRe = [Regex]::Escape($displayName)
    if ($t -notmatch ('^(?i)' + $dnRe)) {
        if ($t -match '^[\w]') { $t = "$displayName — $t" } else { $t = "$displayName $t" }
    }

    # Normalize whitespace
    $t = ($t -replace '\s+',' ').Trim()
    return $t
}

function Get-LLMApiKey([string]$EnvVarName) {
    $k = [Environment]::GetEnvironmentVariable($EnvVarName, "Process")
    if (-not $k) { $k = [Environment]::GetEnvironmentVariable($EnvVarName, "User") }
    if (-not $k) { $k = [Environment]::GetEnvironmentVariable($EnvVarName, "Machine") }
    return $k
}

function Invoke-LLMDescription {
    param(
        [Parameter(Mandatory=$true)][string]$ApiUrl,
        [Parameter(Mandatory=$false)][string]$ApiKey,
        [Parameter(Mandatory=$false)][string]$Model = "gpt-40",
        [Parameter(Mandatory=$true)][string]$DisplayName,
        [int]$TimeoutMs = 20000
    )
    try {
        $prompt = Build-LLMPrompt -displayName $DisplayName
        $headers = @{}
        if ($ApiKey) { $headers["Authorization"] = "Bearer $ApiKey" }
        $headers["Content-Type"] = "application/json"

        $isOpenAI = $ApiUrl -match '/v1/(chat/)?completions'
        if ($isOpenAI) {
            $body = @{
                model = $Model
                messages = @(
                    @{ role="system"; content="You are a concise technical writer for end-user software offers." },
                    @{ role="user";   content=$prompt }
                )
                temperature = 0.3
                max_tokens  = 200
            }
        } else {
            $body = @{
                model  = $Model
                prompt = $prompt
                temperature = 0.3
                max_tokens  = 200
            }
        }

        $json = $body | ConvertTo-Json -Depth 6
        $resp = Invoke-RestMethod -Method Post -Uri $ApiUrl -Headers $headers -Body $json -TimeoutSec ([math]::Ceiling($TimeoutMs/1000))

        $candidates = @()
        if ($resp -and $resp.choices -and $resp.choices.Count -gt 0) {
            if ($resp.choices[0].message.content) { $candidates += [string]$resp.choices[0].message.content }
            elseif ($resp.choices[0].text)        { $candidates += [string]$resp.choices[0].text }
        }
        if ($resp.description) { $candidates += [string]$resp.description }
        if ($resp.output)      { $candidates += [string]$resp.output }
        if ($resp.result)      { $candidates += [string]$resp.result }
        if ($resp.message)     { $candidates += [string]$resp.message }

        $txt = $null
        foreach ($c in $candidates) { if (-not [string]::IsNullOrWhiteSpace($c)) { $txt = $c; break } }

        if (-not $txt) { return $null }
        $txt = $txt -replace '\s+',' ' | ForEach-Object { $_.Trim() }
        $txt = Normalize-OfferDescription -text $txt -displayName $DisplayName
        return (Truncate $txt 600)
    }
    catch {
        LogLine ("❌ LLM call failed: {0}" -f ($_.Exception.GetBaseException().Message))
        return $null
    }
}

# =========================
# HTTP
# =========================
if ($IgnoreCertErrors) { try { [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } } catch { } }
[System.Net.ServicePointManager]::Expect100Continue = $false

function HttpGetXml {
    param([string]$Url,[string]$AuthHeader)
    $req = [System.Net.HttpWebRequest]::Create($Url)
    $req.Method = "GET"
    $req.Accept = "application/xml"
    $req.Headers["Accept-Encoding"] = "gzip, deflate"
    $req.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate
    if ($AuthHeader) { $req.Headers["Authorization"] = $AuthHeader }
    $req.ProtocolVersion = [Version]"1.1"
    $req.PreAuthenticate = $true
    $req.AllowAutoRedirect = $false
    $req.Timeout = 45000
    try {
        $resp = $req.GetResponse()
        try {
            $sr = New-Object IO.StreamReader($resp.GetResponseStream(), [Text.Encoding]::UTF8)
            $content = $sr.ReadToEnd(); $sr.Close()
        } finally { $resp.Close() }
        return $content
    } catch {
        throw ($_.Exception.GetBaseException().Message)
    }
}

# Enhanced: detailed error capture (PS 5.x safe)
function Post-XmlFile-InFile {
    param([string]$Url,[string]$User,[string]$Pass,[string]$XmlFilePath)

    try {
        $pair  = "$User`:$Pass"
        $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
        $basic = "Basic " + [Convert]::ToBase64String($bytes)

        $resp = Invoke-WebRequest -Method Post -Uri $Url `
            -Headers @{ "Authorization" = $basic } `
            -ContentType "application/xml" `
            -InFile $XmlFilePath `
            -UseBasicParsing `
            -ErrorAction Stop

        if ($resp.StatusCode) { LogLine ("POST HTTP {0} {1}" -f [int]$resp.StatusCode, $resp.StatusDescription) }
        if ($resp.Content)    { LogLine ("POST response body (first 500 chars): {0}" -f ($resp.Content.Substring(0,[Math]::Min(500,$resp.Content.Length)))) }
        return
    }
    catch {
        $ex = $_.Exception
        $status     = $null
        $statusDesc = $null
        $errBody    = $null
        $headersStr = $null

        $respErr = $ex.Response
        if ($respErr) {
            try {
                if ($respErr.StatusCode)        { $status     = [int]$respErr.StatusCode }
                if ($respErr.StatusDescription) { $statusDesc = $respErr.StatusDescription }
            } catch {}

            try {
                $hdrs = @()
                foreach ($k in $respErr.Headers.Keys) {
                    $v = $respErr.Headers[$k]
                    if ($k -and $v) { $hdrs += ("{0}: {1}" -f $k, $v) }
                    if ($hdrs.Count -ge 8) { break }
                }
                if ($hdrs.Count -gt 0) { $headersStr = [string]::Join("; ", $hdrs) }
            } catch {}
        }

        if ($null -eq $errBody -and $_.ErrorDetails -and $_.ErrorDetails.Message) { $errBody = [string]$_.ErrorDetails.Message }

        if ($null -eq $errBody -and $respErr -and $respErr.GetResponseStream) {
            try {
                $rs = $respErr.GetResponseStream()
                if ($rs) {
                    $sr = New-Object IO.StreamReader($rs, [Text.Encoding]::UTF8)
                    $errBody = $sr.ReadToEnd(); $sr.Close()
                }
            } catch {}
        }

        if ($null -eq $errBody -and $ex.InnerException) {
            $errBody = [string]$ex.InnerException.Message
        }

        $statusText = if ($null -ne $status) { [string]$status } else { "?" }
        $descText   = if ($null -ne $statusDesc) { [string]$statusDesc } else { "" }

        LogLine ("❌ POST failed :: HTTP {0} {1}" -f $statusText, $descText)
        if ($headersStr) { LogLine ("Response headers: {0}" -f $headersStr) }

        if ($errBody) {
            $trimmed = $errBody.Substring(0, [Math]::Min(2000, $errBody.Length))
            LogLine ("❌ Server error body (first 2000 chars): {0}" -f $trimmed)
            $errFile = Join-Path $env:TEMP ("BES_Post_Error_{0:yyyyMMdd_HHmmss}.txt" -f (Get-Date))
            try { [System.IO.File]::WriteAllText($errFile, $errBody, [Text.Encoding]::UTF8); LogLine ("Saved full error to: {0}" -f $errFile) } catch {}
        } else {
            LogLine ("❌ Server did not include an error body.")
        }

        $short = if ($errBody) { $errBody.Substring(0, [Math]::Min(300, $errBody.Length)) } else { "" }
        throw ("Invoke-WebRequest POST failed :: HTTP {0} {1} :: {2}" -f $statusText, $descText, $short)
    }
}

# =========================
# FIXLET & GROUP PARSING
# =========================
function Get-FixletContainer { param([xml]$Xml)
    if ($Xml.BES.Fixlet)   { return @{ Type="Fixlet";   Node=$Xml.BES.Fixlet } }
    if ($Xml.BES.Task)     { return @{ Type="Task";     Node=$Xml.BES.Task } }
    if ($Xml.BES.Baseline) { return @{ Type="Baseline"; Node=$Xml.BES.Baseline } }
    throw "Unknown BES content type (no <Fixlet>, <Task>, or <Baseline>)."
}

function Extract-AllRelevanceFromXmlString {
    param([string]$XmlString,[string]$Context = "Unknown")
    $all = @()
    try {
        $x = [xml]$XmlString
        $cgRels = $x.SelectNodes("//*[local-name()='ComputerGroup']//*[local-name()='Relevance']")
        if ($cgRels) { foreach ($n in $cgRels) { $t = ($n.InnerText).Trim(); if ($t) { $all += $t } } }
        if ($all.Count -gt 0) { return ,$all }
        $globalRels = $x.SelectNodes("//*[local-name()='Relevance']")
        if ($globalRels) { foreach ($n in $globalRels) { $t = ($n.InnerText).Trim(); if ($t) { $all += $t } } }
    } catch {}
    if ($AggressiveRegexFallback -and $all.Count -eq 0) {
        $regex = [regex]'(?is)<Relevance\b[^>]*>(.*?)</Relevance>'
        foreach ($mm in $regex.Matches($XmlString)) { $t = ($mm.Groups[1].Value).Trim(); if ($t) { $all += $t } }
    }
    return ,$all
}

function Extract-SCRFragments {
    param([string]$XmlString,[string]$Context="Unknown")
    $frags = @()
    try {
        $x = [xml]$XmlString
        $scrNodes = $x.SelectNodes("//*[local-name()='SearchComponentRelevance']")
        if ($scrNodes) {
            foreach ($n in $scrNodes) {
                $innerR = $n.SelectNodes(".//*[local-name()='Relevance']")
                if ($innerR -and $innerR.Count -gt 0) {
                    foreach ($ir in $innerR) { $t = ($ir.InnerText).Trim(); if ($t) { $frags += $t } }
                } else {
                    $t = ($n.InnerText).Trim(); if ($t) { $frags += $t }
                }
            }
        }
    } catch {}
    return ,$frags
}

function Get-GroupClientRelevance {
    param([string]$BaseUrl,[string]$AuthHeader,[string]$SiteName,[string]$GroupIdNumeric)

    $encSite = Encode-SiteName $SiteName
    $candidates = @(
        "/api/computergroup/custom/$encSite/$GroupIdNumeric",
        "/api/computergroup/master/$GroupIdNumeric",
        "/api/computergroup/operator/$($env:USERNAME)/$GroupIdNumeric"
    )

    foreach ($relPath in $candidates) {
        $url = Join-ApiUrl -BaseUrl $BaseUrl -RelativePath $relPath
        try {
            $xmlStr = HttpGetXml -Url $url -AuthHeader $AuthHeader
            if ($DumpFetchedXmlToTemp) {
                $tmp = Join-Path $env:TEMP ("BES_ComputerGroup_{0}.xml" -f $GroupIdNumeric)
                [System.IO.File]::WriteAllText($tmp, $xmlStr)
                LogLine "Saved fetched group XML to: $tmp"
            }
            $rels = Extract-AllRelevanceFromXmlString -XmlString $xmlStr -Context "Group:$GroupIdNumeric"
            if ($rels.Count -gt 0) {
                $joined = ($rels | ForEach-Object { "($_)" }) -join " AND "
                return $joined
            }
            $frags = Extract-SCRFragments -XmlString $xmlStr -Context "Group:$GroupIdNumeric"
            if ($frags.Count -gt 0) {
                $joined = ($frags | ForEach-Object { "($_)" }) -join " AND "
                return $joined
            }
            LogLine "No usable relevance at ${url}"
        } catch {
            LogLine ("❌ Group relevance fetch failed ({0}): {1}" -f $GroupIdNumeric, $_.Exception.Message)
        }
    }
    throw "No relevance found or derivable for group ${GroupIdNumeric}."
}

# NEW: Direct membership clause (toggleable form)
function Build-GroupMembershipRelevance(
    [string]$SiteName,
    [string]$GroupIdNumeric,
    [bool]$UseSitesPluralLocal = $UseSitesPlural
) {
    if ($UseSitesPluralLocal) {
        # Site-agnostic: matches the group id in any site
        return "(member of group $GroupIdNumeric of sites)"
    } else {
        # Site-specific: match only the named site
        $siteEsc = $SiteName.Replace('"','\"')
        return "(member of group $GroupIdNumeric of site whose (name of it = `"$siteEsc`"))"
    }
}

# =========================
# OFFER XML BUILDER
# =========================
function Build-OfferXml {
    param(
        [string]$UiBaseTitle,     # fixlet title (for console ActionUITitle context)
        [string]$DisplayName,     # derived from title for user-facing text
        [string]$SiteName,
        [string]$FixletId,
        [string]$FixletActionName,
        [string]$GroupRelevance,
        [string]$Kind,            # "Install" | "Update" | "Remove"
        [string]$Phase,           # "QA" | "Deploy"
        [string]$OfferDescription # from UI (Application Description)
    )

    $uiTitleMessage  = SafeEscape($DisplayName)
    $siteEsc         = SafeEscape($SiteName)
    $fixletIdEsc     = SafeEscape($FixletId)
    $actionNameEsc   = SafeEscape($FixletActionName)
    $groupSafe       = if ([string]::IsNullOrWhiteSpace($GroupRelevance)) { "" } else { $GroupRelevance }
    $groupSafe       = $groupSafe -replace ']]>', ']]]]><![CDATA[>'

    switch -Regex ($Kind) {
        '^(?i)install$' { $ing='Installing'; $cat='Install'; $verb='install' }
        '^(?i)remove$'  { $ing='Removing' ; $cat='Remove' ; $verb='remove'  }
        default         { $ing='Updating' ; $cat='Update' ; $verb='update'  }
    }

    # Messages (no parentheses; only Update uses "to")
    if ($cat -ieq 'Update') {
        $runningMsgText = ("{0} to {1}. Please wait..." -f $ing, $DisplayName)
    } else {
        $runningMsgText = ("{0} {1}. Please wait..." -f $ing, $DisplayName)
    }
    $runningMsg = SafeEscape($runningMsgText)

    $offerTitle = SafeEscape(("{0}: {1} Win: {2} Offer" -f $cat, $DisplayName, $Phase))

    $descFallback = "This offer will $verb $DisplayName."
    $descRaw = if ([string]::IsNullOrWhiteSpace($OfferDescription)) { $descFallback } else { $OfferDescription }
    $descHtml = [System.Web.HttpUtility]::HtmlEncode($descRaw) -replace "`r?`n","<br/>"
    $offerCat = SafeEscape($cat)

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
      <!-- Messages -->
      <ActionUITitle>$uiTitleMessage</ActionUITitle>
      <PreActionShowUI>$(To-XmlBool $OfferDefaults.PreActionShowUI)</PreActionShowUI>
      <HasRunningMessage>true</HasRunningMessage>
      <RunningMessage><Text>$runningMsg</Text></RunningMessage>

      <!-- Execution -->
      <HasTimeRange>false</HasTimeRange>
      <HasStartTime>true</HasStartTime>
      <StartDateTimeLocalOffset>$($OfferDefaults.StartOffsetISO)</StartDateTimeLocalOffset>
      <HasEndTime>true</HasEndTime>
      <EndDateTimeLocalOffset>$($OfferDefaults.EndOffsetISO)</EndDateTimeLocalOffset>
      <UseUTCTime>false</UseUTCTime>

      <!-- Reapply & Retry -->
      <Reapply>$(To-XmlBool $OfferDefaults.Reapply)</Reapply>
      <HasReapplyLimit>false</HasReapplyLimit>
      <HasReapplyInterval>false</HasReapplyInterval>
      <HasRetry>true</HasRetry>
      <RetryCount>$($OfferDefaults.RetryCount)</RetryCount>
      <RetryWait Behavior="WaitForInterval">$($OfferDefaults.RetryWaitISO)</RetryWait>

      <!-- Other defaults (schema-safe across versions) -->
      <HasTemporalDistribution>false</HasTemporalDistribution>
      <ContinueOnErrors>$(To-XmlBool $OfferDefaults.ContinueOnErr)</ContinueOnErrors>
      <PostActionBehavior Behavior="Nothing"></PostActionBehavior>

      <!-- Offer tab -->
      <IsOffer>true</IsOffer>
      <OfferCategory>$offerCat</OfferCategory>
      <OfferDescriptionHTML><![CDATA[$descHtml]]></OfferDescriptionHTML>
    </Settings>

    <!-- Console action title -->
    <Title>$offerTitle</Title>
  </SourcedFixletAction>
</BES>
"@
}

# =========================
# GUI (inputs → buttons → log)
# =========================
$form = New-Object System.Windows.Forms.Form
$form.Text = "BigFix Offer Action Generator"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(900, 740)
$form.MinimumSize = New-Object System.Drawing.Size(900, 620)

$root = New-Object System.Windows.Forms.TableLayoutPanel
$root.Dock = 'Fill'
$root.AutoSize = $false
$root.RowCount = 3
$root.ColumnCount = 1
$root.Padding = '10,10,10,10'
$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,100)))
$root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))
$form.Controls.Add($root)

$grid = New-Object System.Windows.Forms.TableLayoutPanel
$grid.Dock = 'Top'
$grid.AutoSize = $true
$grid.AutoSizeMode = 'GrowAndShrink'
$grid.ColumnCount = 2
$grid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 240)))
$grid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

function Add-Row([string]$labelText, [System.Windows.Forms.Control]$ctrl) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $labelText
    $lbl.AutoSize = $true
    $lbl.Margin = '0,6,12,6'
    $ctrl.Margin = '0,2,0,6'
    $ctrl.Anchor = 'Left,Right'
    $grid.RowCount += 1
    $grid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $grid.Controls.Add($lbl, 0, $grid.RowCount - 1)
    $grid.Controls.Add($ctrl, 1, $grid.RowCount - 1)
}

$tbServer = New-Object System.Windows.Forms.TextBox
$tbServer.Text = "https://test.server:52311"
$tbServer.ReadOnly = $true
$tbServer.BackColor = [System.Drawing.SystemColors]::Window
Add-Row "BigFix Server:" $tbServer

$tbUser = New-Object System.Windows.Forms.TextBox
Add-Row "Username:" $tbUser

$tbPass = New-Object System.Windows.Forms.MaskedTextBox
$tbPass.PasswordChar = '*'
Add-Row "Password:" $tbPass

$tbFixletInstall = New-Object System.Windows.Forms.TextBox
Add-Row "Fixlet ID (Install):" $tbFixletInstall

$tbFixletUpdate = New-Object System.Windows.Forms.TextBox
Add-Row "Fixlet ID (Update):" $tbFixletUpdate

$tbFixletRemove = New-Object System.Windows.Forms.TextBox
Add-Row "Fixlet ID (Remove):" $tbFixletRemove

$tbAppDesc = New-Object System.Windows.Forms.TextBox
$tbAppDesc.Multiline = $true
$tbAppDesc.ScrollBars = 'Vertical'
$tbAppDesc.Height = 120
$tbAppDesc.Anchor = 'Left,Right'
$tbAppDesc.Margin = '0,2,0,6'
$lblDesc = New-Object System.Windows.Forms.Label
$lblDesc.Text = "Application Description (Offer tab):"
$lblDesc.AutoSize = $true
$lblDesc.Margin = '0,6,12,6'
$grid.RowCount += 1
$grid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$grid.Controls.Add($lblDesc, 0, $grid.RowCount - 1)
$grid.Controls.Add($tbAppDesc, 1, $grid.RowCount - 1)

$btnLLM = New-Object System.Windows.Forms.Button
$btnLLM.Text = "Draft Description with LLM"
$btnLLM.Height = 28
$btnLLM.Width  = 240
$btnLLM.Anchor = 'Left'
Add-Row "" $btnLLM

$root.Controls.Add($grid, 0, 0)

$btnPanel = New-Object System.Windows.Forms.TableLayoutPanel
$btnPanel.Dock = 'Top'
$btnPanel.AutoSize = $true
$btnPanel.AutoSizeMode = 'GrowAndShrink'
$btnPanel.ColumnCount = 1
$btnPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))

$btnQA = New-Object System.Windows.Forms.Button
$btnQA.Text = "Create QA Offers (Install / Update / Remove)"
$btnQA.Height = 36
$btnQA.Dock = 'Top'
$btnQA.Margin = '0,4,0,6'
$btnPanel.Controls.Add($btnQA)

$btnDeploy = New-Object System.Windows.Forms.Button
$btnDeploy.Text = "Create Deploy Offers (Install / Update / Remove)"
$btnDeploy.Height = 36
$btnDeploy.Dock = 'Top'
$btnDeploy.Margin = '0,0,0,6'
$btnPanel.Controls.Add($btnDeploy)

$root.Controls.Add($btnPanel, 0, 1)

$LogBox = New-Object System.Windows.Forms.TextBox
$LogBox.Multiline = $true
$LogBox.ScrollBars = 'Vertical'
$LogBox.ReadOnly = $false
$LogBox.WordWrap = $false
$LogBox.Dock = 'Fill'
$LogBox.Margin = '0,6,0,0'
$LogBox.ContextMenu = New-Object System.Windows.Forms.ContextMenu
$LogBox.ContextMenu.MenuItems.AddRange(@(
    (New-Object System.Windows.Forms.MenuItem "Copy",       { $LogBox.Copy() }),
    (New-Object System.Windows.Forms.MenuItem "Select All", { $LogBox.SelectAll() })
))
$root.Controls.Add($LogBox, 0, 2)

# =========================
# CORE: confirm + post
# =========================
function Build-OfferXml-Wrapper {
    param(
        [string]$Kind,
        [string]$Phase,
        [string]$CustomSiteName,
        [string]$FixletId,
        [string]$FixletActionName,
        [string]$GroupRelevance,
        [string]$TitleFromServer,
        [string]$OfferDescriptionFromUi
    )
    $dispName = Parse-FixletTitleToProduct -Title $TitleFromServer
    return (Build-OfferXml `
        -UiBaseTitle          $TitleFromServer `
        -DisplayName          $dispName `
        -SiteName             $CustomSiteName `
        -FixletId             $FixletId `
        -FixletActionName     $FixletActionName `
        -GroupRelevance       $GroupRelevance `
        -Kind                 $Kind `
        -Phase                $Phase `
        -OfferDescription     $OfferDescriptionFromUi)
}

function Confirm-And-Post-Offers {
    param(
        [string]$Phase,                 # "QA" or "Deploy"
        [string]$Server,
        [string]$User,
        [string]$Pass,
        [string]$FixletInstall,
        [string]$FixletUpdate,
        [string]$FixletRemove,
        [string]$OfferDescriptionFromUi
    )

    $LogBox.Clear()
    LogLine "== Begin $Phase Offers =="

    if (-not ($Server -and $User -and $Pass -and $FixletInstall -and $FixletUpdate -and $FixletRemove)) {
        LogLine "❌ Fill Server, Username, Password, and all 3 Fixlet IDs."
        return
    }

    try {
        $base = Get-BaseUrl $Server
        $encodedSite = Encode-SiteName $CustomSiteName
        $auth = Get-AuthHeader -User $User -Pass $Pass
        $postUrl = Join-ApiUrl -BaseUrl $base -RelativePath "/api/actions"
        LogLine "POST URL: ${postUrl}"

        # Pull titles
        $ids = @($FixletInstall,$FixletUpdate,$FixletRemove)
        $titles = @()
        foreach ($fx in $ids) {
            $fixUrl = Join-ApiUrl -BaseUrl $base -RelativePath "/api/fixlet/custom/$encodedSite/$fx"
            $xmlStr = HttpGetXml -Url $fixUrl -AuthHeader $auth
            if ($DumpFetchedXmlToTemp) {
                $tmpFix = Join-Path $env:TEMP ("BES_Fixlet_{0}.xml" -f $fx)
                [System.IO.File]::WriteAllText($tmpFix, $xmlStr)
                LogLine "Saved fetched fixlet XML to: $tmpFix"
            }
            $x = [xml]$xmlStr
            $node = (Get-FixletContainer -Xml $x).Node
            $title = [string]$node.Title
            if (-not $title) { $title = "(Unknown Title: $fx)" }
            $titles += $title
        }

        # Resolve group relevance based on mode
        $groupRaw = if ($Phase -ieq "QA") { $QA_GroupIdWithPrefix } else { $Deploy_GroupIdWithPrefix }
        $groupNum = Get-NumericGroupId $groupRaw
        if (-not $groupNum) { throw "Could not parse numeric ID from '$groupRaw'." }

        $groupRel = $null

        if ($UseDirectGroupMembershipRelevance) {
            # Use direct membership relevance (recommended) — site agnostic by default
            $groupRel = Build-GroupMembershipRelevance -SiteName $GroupSiteNameOverride -GroupIdNumeric $groupNum
            LogLine ("Using direct membership relevance: {0}" -f $groupRel)
        } else {
            # Fetch the group's actual relevance (supports cross-site via override)
            $siteForGroup = if ([string]::IsNullOrWhiteSpace($GroupSiteNameOverride)) { $CustomSiteName } else { $GroupSiteNameOverride }
            LogLine ("Fetching group relevance from site: {0}" -f $siteForGroup)
            $groupRel = Get-GroupClientRelevance -BaseUrl $base -AuthHeader $auth -SiteName $siteForGroup -GroupIdNumeric $groupNum
        }

        # Confirm (no inline if in array literal)
        $lines = @()
        $lines += "Round: $Phase"
        $lines += "Target Group: $groupRaw"
        if ($UseDirectGroupMembershipRelevance) {
            if ($UseSitesPlural) {
                $lines += "Target Mode: Direct membership clause (of sites)"
            } else {
                $lines += "Target Mode: Direct membership clause (specific site)"
            }
        } else {
            $lines += "Target Mode: Fetched group relevance"
        }
        $lines += ""
        $lines += "Fixlets:"
        $lines += " - Install: $($titles[0])"
        $lines += " - Update : $($titles[1])"
        $lines += " - Remove : $($titles[2])"
        $lines += ""
        $lines += "Create OFFER actions for these 3 fixlets?"

        $msg = [string]::Join("`r`n", $lines)
        $dlg = [System.Windows.Forms.MessageBox]::Show(
            $form, $msg, "Confirm: Create $Phase Offers",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button2
        )
        if ($dlg -ne [System.Windows.Forms.DialogResult]::Yes) {
            LogLine "🚫 User canceled."
            return
        }

        # Post each
        $triples = @(
            @{ Name="Install"; FixId=$FixletInstall; Title=$titles[0] },
            @{ Name="Update";  FixId=$FixletUpdate;  Title=$titles[1] },
            @{ Name="Remove";  FixId=$FixletRemove;  Title=$titles[2] }
        )

        foreach ($t in $triples) {
            $xmlBody = Build-OfferXml-Wrapper `
                -Kind $t.Name `
                -Phase $Phase `
                -CustomSiteName $CustomSiteName `
                -FixletId $t.FixId `
                -FixletActionName $FixletActionName_Default `
                -GroupRelevance $groupRel `
                -TitleFromServer $t.Title `
                -OfferDescriptionFromUi $tbAppDesc.Text

            $safeTitle = ($t.Name -replace '[^\w\-. ]','_') -replace '\s+','_'
            $tmpAction = Join-Path $env:TEMP ("BES_Offer_{0}_{1:yyyyMMdd_HHmmss}.xml" -f $safeTitle,(Get-Date))
            if ($SaveActionXmlToTemp) {
                [System.IO.File]::WriteAllText($tmpAction, $xmlBody)
                LogLine "Saved OFFER action XML for $($t.Name) to: $tmpAction"
                LogLine ("curl -k -u USER:PASS -H `"Content-Type: application/xml`" -d @`"$tmpAction`" {0}" -f $postUrl)
            }

            try {
                if ($PostUsingInvokeWebRequest -and (Test-Path $tmpAction)) {
                    Post-XmlFile-InFile -Url $postUrl -User $User -Pass $Pass -XmlFilePath $tmpAction
                } else {
                    LogLine "⚠️ Direct POST path disabled; enable if needed."
                }
                LogLine ("✅ OFFER posted: {0}" -f $t.Name)
            } catch {
                LogLine ("❌ OFFER POST failed for {0}: {1}" -f $t.Name, $_.Exception.Message)
            }
        }

        LogLine "All $Phase offers attempted. Log file: $LogFile"
    }
    catch {
        LogLine ("❌ Fatal error ($Phase): {0}" -f ($_.Exception.GetBaseException().Message))
    }
}

# Wire up buttons
$btnQA.Add_Click({
    Confirm-And-Post-Offers `
        -Phase "QA" `
        -Server $tbServer.Text `
        -User $tbUser.Text `
        -Pass $tbPass.Text `
        -FixletInstall $tbFixletInstall.Text `
        -FixletUpdate  $tbFixletUpdate.Text `
        -FixletRemove  $tbFixletRemove.Text `
        -OfferDescriptionFromUi $tbAppDesc.Text
})
$btnDeploy.Add_Click({
    Confirm-And-Post-Offers `
        -Phase "Deploy" `
        -Server $tbServer.Text `
        -User $tbUser.Text `
        -Pass $tbPass.Text `
        -FixletInstall $tbFixletInstall.Text `
        -FixletUpdate  $tbFixletUpdate.Text `
        -FixletRemove  $tbFixletRemove.Text `
        -OfferDescriptionFromUi $tbAppDesc.Text
})

# LLM draft button (uses env var key)
$btnLLM.Add_Click({
    try {
        if ([string]::IsNullOrWhiteSpace($tbServer.Text) -or
            [string]::IsNullOrWhiteSpace($tbUser.Text)   -or
            [string]::IsNullOrWhiteSpace($tbPass.Text)) {
            [System.Windows.Forms.MessageBox]::Show($form, "Enter Server, Username and Password first.", "Missing Credentials", 'OK', 'Warning') | Out-Null
            return
        }
        $fx = $tbFixletInstall.Text
        if ([string]::IsNullOrWhiteSpace($fx)) {
            LogLine "❌ Enter 'Fixlet ID (Install)' first so I can infer the product name."
            [System.Windows.Forms.MessageBox]::Show($form, "Enter the Install Fixlet ID first so I can derive the product name from its title.", "Need Install Fixlet ID", "OK", "Information") | Out-Null
            return
        }

        $apiKey = Get-LLMApiKey -EnvVarName $LLMConfig.ApiKeyEnv
        if ([string]::IsNullOrWhiteSpace($apiKey)) {
            LogLine "❌ LLM API key not found in env var '$($LLMConfig.ApiKeyEnv)'."
            [System.Windows.Forms.MessageBox]::Show($form, "LLM API key not found in environment variable '$($LLMConfig.ApiKeyEnv)'. Set it and try again.", "Missing LLM Key", "OK", "Warning") | Out-Null
            return
        }

        $base = Get-BaseUrl $tbServer.Text
        $encodedSite = Encode-SiteName $CustomSiteName
        $auth = Get-AuthHeader -User $tbUser.Text -Pass $tbPass.Text

        $fixUrl = Join-ApiUrl -BaseUrl $base -RelativePath "/api/fixlet/custom/$encodedSite/$fx"
        $xmlStr = HttpGetXml -Url $fixUrl -AuthHeader $auth
        $x      = [xml]$xmlStr
        $node   = (Get-FixletContainer -Xml $x).Node
        $title  = [string]$node.Title
        if (-not $title) { $title = "(Unknown Title: $fx)" }
        $displayName = Parse-FixletTitleToProduct -Title $title

        LogLine ("LLM drafting for product: {0}" -f $displayName)

        $desc = Invoke-LLMDescription `
            -ApiUrl $LLMConfig.ApiUrl `
            -ApiKey $apiKey `
            -Model  $LLMConfig.Model `
            -DisplayName $displayName

        if ($desc) {
            $tbAppDesc.Text = $desc
            LogLine "✅ LLM description generated and inserted."
        } else {
            LogLine "⚠️ LLM returned no description; leaving the box unchanged."
            [System.Windows.Forms.MessageBox]::Show($form, "The LLM didn’t return a description. You can type one manually.", "No Description", 'OK', 'Warning') | Out-Null
        }
    } catch {
        LogLine ("❌ LLM draft failed: {0}" -f ($_.Exception.GetBaseException().Message))
        [System.Windows.Forms.MessageBox]::Show($form, "LLM draft failed. Check the log for details.", "Error", 'OK', 'Error') | Out-Null
    }
})

$form.Topmost = $false
[void]$form.ShowDialog()
