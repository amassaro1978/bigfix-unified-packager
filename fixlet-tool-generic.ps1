<#
.SYNOPSIS
 A PowerShell based tool to present a front GUI to allow automatic creation of packaging fixlets.

.DESCRIPTION
 Click the "Select Deployment Specs Document" button to select your deployment specs document and it will auto-populate all of the fields. 
 Enter your credentials at the bottom of the tool and click the "Generate and POST Fixlets" button to create and post the 3 packaging fixlets.

.AUTHOR
 Anthony Massaro
#>

# Add some needed Assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Specify the current version of the tool
$ToolVer = "0.1.2"

# Function to retreive the information from the deployment specs document
function Get-WordXML {
 param($docxpath)

 $tempFolder = Join-Path $env:TEMP ("docx_" + [System.IO.Path]::GetFileNameWithoutExtension($docxPath))

 if (Test-path $tempFolder) { Remove-Item -Path $tempFolder -Recurse -Force }
 [System.IO.Compression.ZipFile]::ExtractToDirectory($docxPath, $tempFolder)

 $xmlPath = Join-Path $tempFolder "word/document.xml"
 $xmlDoc = New-Object System.Xml.XmlDocument
 $xmlDoc.Load($xmlPath)
 return $xmlDoc
}

# Function to create and customize the text boxes for information entry
function New-TextBox {
 param ($x, $y, $width = 580, $height = 20, $multiline = $false)
 $box = New-Object System.Windows.Forms.TextBox
 $box.Location = New-Object System.Drawing.Point($x, $y)
 $box.Size = New-Object System.Drawing.Size($width, $height)
 $box.Multiline = $multiline
 return $box
}

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "BigFix Fixlet Generator and Importer v $toolVer"
$form.Size = New-Object System.Drawing.Size(750, 950)
$form.StartPosition = "CenterScreen"

# Add Select Deployment Doc Button
$btnLoadDocx = New-Object System.Windows.Forms.Button
$btnLoadDocx.text = "Select Deployment Specs Document"
$btnLoadDocx.Width = 200
$btnLoadDocx.Height = 30
$btnLoadDocx.Location = New-Object System.Drawing.Point(10,10) # Top of the form

# Text boxes for form entry
$labelVendor = New-Object System.Windows.Forms.Label -Property @{Text="Vendor:"; Location=New-Object System.Drawing.Point(10,60)}
$textVendor = New-TextBox 120 60

$labelApp = New-Object System.Windows.Forms.Label -Property @{Text="Application Name:"; Location=New-Object System.Drawing.Point(10,90)}
$textApp = New-TextBox 120 90

$labelVer = New-Object System.Windows.Forms.Label -Property @{Text="Version:"; Location=New-Object System.Drawing.Point(10,120)}
$textVer = New-TextBox 120 120

$labelSite = New-Object System.Windows.Forms.Label -Property @{Text="Site:"; Location=New-Object System.Drawing.Point(10,150)}
$textSite = New-Object System.Windows.Forms.ComboBox
$textSite.Location = New-Object System.Drawing.Point(120,150)
$textSite.Width = 400
$textSite.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$textSite.Items.Add("test site 1")
$textSite.Items.Add("test 2 (Workstations)")
$textSite.SelectedIndex = 0
$Form.Controls.Add($textsite)

# Icon file picker
$labelIcon = New-Object System.Windows.Forms.Label -Property @{Text="Icon File:"; Location=New-Object System.Drawing.Point(10,180)}
$textIcon = New-TextBox 120 180 400
$btnBrowseIcon = New-Object System.Windows.Forms.Button -Property @{
 Text="Browse..."
 Location=New-Object System.Drawing.Point(530,180)
 Size=New-Object System.Drawing.Size(80,22)
}
$btnBrowseIcon.Add_Click({
 $ofd = New-Object System.Windows.Forms.OpenFileDialog
 $ofd.Filter = "Image Files (*.ico;*.png;*.gif)|*.ico;*.png;*.gif|All files (*.*)|*.*"
 if ($ofd.ShowDialog() -eq "OK") {
 $textIcon.Text = $ofd.FileName
 }
})
$form.Controls.AddRange(@($labelVendor, $textVendor, $labelapp, $textApp, $labelver, $textVer, $labelSite, $textSite, $labelIcon,$textIcon,$btnBrowseIcon))

# Function to create the fixlet blocks
function Add-FixletBlock {
 param($form, $labelText, $yStart, [ref]$relRef, [ref]$actRef)$labelRel = New-Object System.Windows.Forms.Label -Property @{Text="$labelText Relevance:"; Location=New-Object System.Drawing.Point(10, $ystart); Size=New-Object System.Drawing.Size(140,20)}
 $boxRel = New-TextBox 150 $yStart 580 60 $true

 $labelAct = New-Object System.Windows.Forms.Label -Property @{Text="$labelText Action Script:"; Location=New-Object System.Drawing.Point(10, ($ystart + 70)); Size=New-Object System.Drawing.Size(140,20)}
 $boxAct = New-TextBox 150 ($yStart +70) 580 60 $true

 $form.Controls.AddRange(@($labelRel, $boxRel, $labelAct, $boxAct))
 $relRef.Value = $boxRel
 $actRef.Value = $boxAct
}

$InstallRel = $null; $InstallAct = $null
$UpdateRel = $null; $UpdateAct = $null
$RemoveRel = $null; $RemoveAct = $null

Add-FixletBlock -form $form -labeltext "Install" -yStart 220 -relRef ([ref]$InstallRel) -actRef ([ref]$InstallAct)
Add-FixletBlock -form $form -labeltext "Update" -yStart 380 -relRef ([ref]$UpdateRel) -actRef ([ref]$UpdateAct)
Add-FixletBlock -form $form -labeltext "Remove" -yStart 540 -relRef ([ref]$RemoveRel) -actRef ([ref]$RemoveAct)

# Create the drop down menu for DEV and PROD Bigfix servers
$labelServer = New-Object System.Windows.Forms.Label -Property @{Text="BigFix Server:"; Location=New-Object System.Drawing.Point(10,700); Size=New-Object System.Drawing.Size(140,20)}

$textServer = New-Object System.Windows.Forms.ComboBox
$textServer.Location = New-Object System.Drawing.Point(150,700)
$textServer.Width = 400
$textServer.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$textServer.Items.Add("https://dev.server:52311")
$textServer.Items.Add("https://prod.server:52311")
$textServer.SelectedIndex = 0
$Form.Controls.Add($textserver)

#Create the username and password fields
$labelUser = New-Object System.Windows.Forms.Label -Property @{Text="Username:"; Location=New-Object System.Drawing.Point(10,730); Size=New-Object System.Drawing.Size(140,20)}
$textUser = New-TextBox 150 730 200

$labelPass = New-Object System.Windows.Forms.Label -Property @{Text="Password:"; Location=New-Object System.Drawing.Point(10,760); Size=New-Object System.Drawing.Size(140,20)}
$textPass = New-TextBox 150 760 200
$TextPass.PasswordChar = '*'

$form.Controls.AddRange(@($labelServer, $textServer, $labelUser, $textUser, $labelPass, $textPass))

# Function to form the XML body and to post the fixlet
function Post-Fixlet {
 param($type, $vendor, $app, $ver, $rel, $act, $encodedsite, $server, $cred, $logpath, $timestamp)

 $title = "${Type}: $vendor $app $ver Win"
 $releaseDate = (Get-Date).ToString("yyyy-MM-dd")
 $description = "This fixlet will $($type.ToLower()) $vendor $app $ver"
 $map = @{ 'Install' = 'Install'; 'Update' = 'Pending'; 'Remove' = 'Remove' }
 $category = $map[$type]
 $source = $map[$type]

 # Encode relevance
 $relevanceXML = ($rel -split "`r?`n" | Where-Object { $_.Trim() -ne "" } | ForEach-Object { "<Relevance><![CDATA[$_]]></Relevance>" }) -join "`n"

 # Encode icon
 $iconBlock = ""
 if ($textIcon.Text -and (Test-Path $textIcon.Text)) {
 try {
 $bytes = [System.IO.File]::ReadAllBytes($textIcon.Text)
 $b64 = [Convert]::ToBase64String($bytes)

 switch ([System.IO.Path]::GetExtension($textIcon.Text).ToLower()) {
 ".png" { $mime = "image/png" }
 ".ico" { $mime = "image/x-icon" }
 ".gif" { $mime = "image/gif" }
 default { $mime = "application/octet-stream" }
 }

 $dataUri = "data:$mime;base64,$b64"
 $iconBlock = @"
 <MIMEField>
 <Name>action-ui-metadata</Name>
 <Value>{"icon":"$dataUri"}</Value>
 </MIMEField>
"@
 } catch {
 Add-Content -Path $logPath -Value "$timestamp - Failed to encode icon: $($_.Exception.Message)"
 }
 }

 $xml =@"
<?xml version="1.0" encoding="UTF-8"?>
<BES xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="BES.xsd">
 <Fixlet>
 <Title>$title</Title><Description>$description</Description>
 $relevancexml
 <Category>$Category</Category>
 <Source>$Source</Source>
 <SourceReleaseDate>$releaseDate</SourceReleaseDate>
 $iconBlock
 <DefaultAction ID="Action1">
 <ActionScript MIMEType="application/x-Fixlet-Windows-Shell"><![CDATA[$act]]></ActionScript>
 </DefaultAction>
 </Fixlet>
</BES>
"@

 $uri = "$server/api/fixlets/custom/$encodedsite"
 [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

 Add-Content -Path $logPath -Value "$timestamp - Encoded Site: $encodedSite"
 Add-Content -Path $logPath -Value "$timestamp - POST to URL: $uri"
 Add-Content -Path $logPath -Value "$timestamp - XML Body: `n$xml`n---"

 try {
 [xml]$parsed = $xml
 Add-Content -Path $logPath -Value "$Timestamp - XML Parsed Successfully"
 } catch { 
 Add-Content -Path $logPath -Value "$Timestamp - XML Parsing Failed: $($_.Exception.Message)"
 }

 return Invoke-RestMethod -Uri $uri -Method Post -Credential $cred -ContentType 'application/xml' -Body $xml -ErrorAction Stop
}

# Event handler for select deployment specs document button
$btnLoadDocx.Add_Click({
 $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $fileDialog.Filter = "Word Documents (*docx)|*.docx"
 $fileDialog.Title = "Select Word Document"

 if ($fileDialog.ShowDialog() -eq "OK") {
 $docPath = $fileDialog.FileName

 try {
 $xmlDoc = Get-WordXML -docxpath $docPath
 $namespaceManager = New-Object System.Xml.XmlNamespaceManager($xmlDoc.NameTable)
 $namespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
 $rows = $xmlDoc.SelectNodes("//w:tbl//w:tr", $namespaceManager)

 $fixletData = @{}
 foreach ($row in $rows) {
 $cells = $row.SelectNodes("w:tc", $namespaceManager)

 if($cells.Count -ge 2) {
 $fieldName = ""
 $fieldValue = ""

 $fieldNameParagraphs = $cells[0].SelectNodes(".//w:t", $namespaceManager)
 foreach ($p in $fieldNameParagraphs) {
 $fieldname += $p.InnerText
 }
 $fieldName = $fieldName.Trim()

 $fieldValueParagraphs = $cells[1].SelectNodes(".//w:p", $namespaceManager)
 foreach ($para in $fieldValueParagraphs) {
 $texts = $para.SelectNodes(".//w:t", $namespaceManager)
 foreach ($text in $texts) {
 $fieldValue += $text.InnerText
 }
 $fieldvalue += "`r`n"
 }

 $fixletData[$fieldName] = $fieldValue.TrimEnd()
 }
 }

 #Populate GUI fields
 $textVendor.Text = $fixletData["Vendor"]
 $textApp.Text = $fixletData["App Name"]
 $textVer.text = $fixletData["Version"]

 $InstallRel.text = $fixletData["Install Relevance"]
 $InstallAct.text = $fixletData["Install Action Script"]

 $UpdateRel.text = $fixletData["Update Relevance"]
 $UpdateAct.text = $fixletData["Update Action Script"]

 $RemoveRel.text = $fixletData["Remove Relevance"]
 $RemoveAct.text = $fixletData["Remove Action Script"]

 [System.Windows.Forms.Messagebox]::Show("Deployment Document Successfully parsed and fields have been updated.", "Success")
 } catch {
 [System.Windows.Forms.MessageBox]::Show("Failed to read or parse the Word Document: $($_.Exception.Message)", "Error")
 }
 }
})

# Add the "Select Deployment Specs Document" button
$form.Controls.Add($btnLoadDocx)

# Add the "Generate and POST Fixlets" button
$btnsend = New-Object System.Windows.Forms.Button
$btnsend.Text = "Generate and POST Fixlets"
$btnsend.Size = '250,30'
$btnsend.Location = New-Object System.Drawing.Point(240, 810)

# Event handler for the "Generate and POST Fixlets" button
$btnsend.Add_Click({
 $vendor = $textVendor.Text.Trim()
 $app = $textApp.Text.Trim()
 $ver = $textVer.Text.Trim()
 $site = $textSite.Text.Trim()
 $encodedSite = $site.Trim() -replace ' ', '%20' -replace '\(', '%28' -replace '\)', '%29'
 $server = $textServer.Text.Trim().TrimEnd('/')
 $user = $textUser.Text.Trim()
 $pass = $textpass.Text

 if (-not ($vendor -and $app -and $ver -and $server -and $user -and $pass)) {
 [System.Windows.Forms.MessageBox]::Show("Please complete all fields")
 return
 }

 $cred = New-Object System.Management.Automation.PSCredential($user, ($pass | ConvertTo-SecureString -AsPlainText -Force))
 $status = ""
 $logPath = Join-Path $PSScriptRoot "FixletPOSTviaAPI.log"
 $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
 $installTitle = "Install: $vendor $app $ver Win"
 $removeTitle = "Remove: $vendor $app $ver Win"
 $updateTitle = "Update: $vendor $app $ver Win"

 function LogResult($name, $result) {
 $entry = "$timestamp - ${name}: $result"
 $status += "${name}: $result`n"
 Add-Content -Path $logpath -Value $entry
 }

 try {
 Post-Fixlet "Install" $vendor $app $ver $installrel.Text $installact.Text $encodedSite $server $cred $logpath $timestamp| Out-Null
 LogResult "$installTitle" "Created"
 } catch {
 LogResult "$installTitle" "Failed - $($_.Exception.Message)"
 }

 try {
 Post-Fixlet "Update" $vendor $app $ver $updaterel.Text $updateact.Text $EncodedSite $server $cred $logpath $timestamp | Out-Null
 LogResult "$updateTitle" "Created"
 } catch {
 LogResult "$updateTitle" "Failed - $($_.Exception.Message)"
 }

 try {
 Post-Fixlet "Remove" $vendor $app $ver $removerel.Text $removeact.Text $EncodedSite $server $cred $logpath $timestamp | Out-Null
 LogResult "$removeTitle" "Created"
 } catch {
 LogResult "$removeTitle" "Failed - $($_.Exception.Message)"
 }

 $lastLine = Get-Content -Path $LogPath | Select-Object -Last 1
 $cleanline = $lastLine -replace '^.*Remove:\s*', ''
 [System.Windows.Forms.MessageBox]::Show("Result: `n$cleanLine")
})

$form.Controls.Add($btnSend)

$form.TopMost = $true
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
