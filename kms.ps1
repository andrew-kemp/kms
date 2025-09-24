# KMS Provisioning Script (PAC & Approvers logic removed)
# This script provisions a SharePoint KMS site and document library.
# Prerequisites: PowerShell 7+, PnP.PowerShell

function Write-Section { param($msg) Write-Host "`n$('='*50)`n$msg`n$('='*50)`n" }

# --- 1. PowerShell 7+ check ---
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Section "ERROR: PowerShell 7 or later is required!"
    Write-Host "You are running PowerShell $($PSVersionTable.PSVersion)."
    Write-Host "Install PowerShell 7 from https://aka.ms/powershell"
    exit 1
}

# --- 2. PnP.PowerShell module check ---
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Section "PnP.PowerShell module not found! Installing..."
    try {
        Install-Module -Name PnP.PowerShell -Force -Scope CurrentUser
        Write-Host "PnP.PowerShell module installed. Please relaunch PowerShell and re-run this script."
        exit 0
    } catch {
        Write-Host "Failed to install PnP.PowerShell. Please install manually:"
        Write-Host "Install-Module -Name PnP.PowerShell -Force -Scope CurrentUser"
        exit 1
    }
}

# --- 3. GUI for KMS Provisioning ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "KMS Provisioner"
$form.Size = "900,830"
$form.StartPosition = "CenterScreen"

$sectionFont = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$labelFont = New-Object System.Drawing.Font("Segoe UI", 10)
$inputFont = New-Object System.Drawing.Font("Segoe UI", 10)
$radioFont = New-Object System.Drawing.Font("Segoe UI", 10)

# --- Certificate Group ---
$grpCert = New-Object System.Windows.Forms.GroupBox
$grpCert.Text = "Certificate Options"
$grpCert.Location = "10, 20"
$grpCert.Size = "850, 130"
$form.Controls.Add($grpCert)

$rbExistingCert = New-Object System.Windows.Forms.RadioButton
$rbExistingCert.Text = "Use Existing Certificate"
$rbExistingCert.Location = "15,30"
$rbExistingCert.Size = "200,25"
$rbExistingCert.Checked = $true
$rbExistingCert.Font = $radioFont

$rbNewCert = New-Object System.Windows.Forms.RadioButton
$rbNewCert.Text = "Create New Certificate"
$rbNewCert.Location = "240,30"
$rbNewCert.Size = "200,25"
$rbNewCert.Font = $radioFont

$lblCertPath = New-Object System.Windows.Forms.Label
$lblCertPath.Text = "Certificate Path (.pfx):"
$lblCertPath.Location = "15,60"
$lblCertPath.Size = "160,22"
$lblCertPath.Font = $labelFont

$txtCertPath = New-Object System.Windows.Forms.TextBox
$txtCertPath.Location = "180,60"
$txtCertPath.Size = "430,22"
$txtCertPath.Font = $inputFont

$btnBrowseCert = New-Object System.Windows.Forms.Button
$btnBrowseCert.Text = "Browse"
$btnBrowseCert.Location = "620,60"
$btnBrowseCert.Size = "80,22"
$btnBrowseCert.Font = $inputFont

$lblCertPass = New-Object System.Windows.Forms.Label
$lblCertPass.Text = "Certificate Password:"
$lblCertPass.Location = "15,90"
$lblCertPass.Size = "160,22"
$lblCertPass.Font = $labelFont

$txtCertPass = New-Object System.Windows.Forms.TextBox
$txtCertPass.Location = "180,90"
$txtCertPass.Size = "180,22"
$txtCertPass.Font = $inputFont
$txtCertPass.UseSystemPasswordChar = $true

$lblCertCN = New-Object System.Windows.Forms.Label
$lblCertCN.Text = "Certificate CN (Subject):"
$lblCertCN.Location = "380,90"
$lblCertCN.Size = "150,22"
$lblCertCN.Font = $labelFont

$txtCertCN = New-Object System.Windows.Forms.TextBox
$txtCertCN.Location = "530,90"
$txtCertCN.Size = "170,22"
$txtCertCN.Font = $inputFont

$btnCreateCert = New-Object System.Windows.Forms.Button
$btnCreateCert.Text = "Create Certificate"
$btnCreateCert.Location = "720,90"
$btnCreateCert.Size = "110,22"
$btnCreateCert.Font = $inputFont
$btnCreateCert.Enabled = $false

$grpCert.Controls.AddRange(@(
    $rbExistingCert, $rbNewCert, $lblCertPath, $txtCertPath, $btnBrowseCert, $lblCertPass, $txtCertPass, $lblCertCN, $txtCertCN, $btnCreateCert
))

# --- KMS Site/App Section ---
$grpKMS = New-Object System.Windows.Forms.GroupBox
$grpKMS.Text = "KMS Site & App Registration"
$grpKMS.Location = "10, 160"
$grpKMS.Size = "850, 330"
$form.Controls.Add($grpKMS)

$labelX = 15
$inputX = 200
$labelW = 180
$inputW = 620
$fy = 30
$gap = 32

# --- Placeholder values for settings ---
$fields = @(
    @{Label="Tenant ID";         Name="txtTenant";        Default="11016236-4dbc-43a6-8310-be803173fc43"},
    @{Label="Tenant Name";       Name="txtTenantName";    Default="kempy"},
    @{Label="App/Client ID";     Name="txtClientId";      Default="1525fc4b-5873-49e3-b0ca-aeb47fee4abd"},
    @{Label="Site Title";        Name="txtSiteTitle";     Default="Knowledge Management System"},
    @{Label="Site short name";   Name="txtSiteShort";     Default="kms"},
    @{Label="Site Owner Email";  Name="txtOwnerEmail";    Default="andrew@kemponline.co.uk"},
    @{Label="Approver Email(s) (comma separated):"; Name="txtApprovers"; Default="andrew@kemponline.co.uk"}
)
$controls = @{}
foreach ($f in $fields) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $f.Label
    $lbl.Location = "$labelX,$fy"
    $lbl.Size = "$labelW,22"
    $lbl.Font = $labelFont
    $grpKMS.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Name = $f.Name
    $txt.Location = "$inputX,$fy"
    $txt.Size = "$inputW,22"
    $txt.Font = $inputFont
    $txt.Text = $f.Default
    $controls[$f.Name] = $txt
    $grpKMS.Controls.Add($txt)
    $fy += $gap
}

$lblDept = New-Object System.Windows.Forms.Label
$lblDept.Text = "Departments:"
$lblDept.Location = "$labelX,$fy"
$lblDept.Size = "$labelW,22"
$lblDept.Font = $labelFont
$grpKMS.Controls.Add($lblDept)

$lstDept = New-Object System.Windows.Forms.ListBox
$lstDept.Location = "$inputX,$fy"
$lstDept.Size = "350,70"
$lstDept.Font = $inputFont
$lstDept.SelectionMode = "One"
@("IT","Finance","HR","Operations","Sales","Marketing","Legal") | ForEach-Object { $lstDept.Items.Add($_) }
$grpKMS.Controls.Add($lstDept)

$txtDeptAdd = New-Object System.Windows.Forms.TextBox
$txtDeptAdd.Location = "570,$fy"
$txtDeptAdd.Size = "120,22"
$txtDeptAdd.Font = $inputFont
$grpKMS.Controls.Add($txtDeptAdd)

$btnDeptAdd = New-Object System.Windows.Forms.Button
$btnDeptAdd.Text = "Add"
$btnDeptAdd.Location = "700,$fy"
$btnDeptAdd.Size = "60,22"
$btnDeptAdd.Font = $inputFont
$grpKMS.Controls.Add($btnDeptAdd)

$btnDeptRemove = New-Object System.Windows.Forms.Button
$btnDeptRemove.Text = "Remove"
$btnDeptRemove.Location = "770,$fy"
$btnDeptRemove.Size = "60,22"
$btnDeptRemove.Font = $inputFont
$grpKMS.Controls.Add($btnDeptRemove)

# --- Provision Button ---
$btnProvision = New-Object System.Windows.Forms.Button
$btnProvision.Text = "Provision Knowledge Management System"
$btnProvision.Font = $sectionFont
$btnProvision.Width = 420
$btnProvision.Height = 40
$btnProvision.Location = New-Object System.Drawing.Point([int](($form.ClientSize.Width - $btnProvision.Width) / 2), 500)
$btnProvision.Enabled = $true
$form.Controls.Add($btnProvision)

# --- Status Box ---
$txtStatus = New-Object System.Windows.Forms.TextBox
$txtStatus.Multiline = $true
$txtStatus.ScrollBars = "Vertical"
$txtStatus.Location = "10,560"
$txtStatus.Size = "850,170"
$txtStatus.Font = New-Object System.Drawing.Font("Consolas",9)
$txtStatus.ReadOnly = $true
$form.Controls.Add($txtStatus)

# --- LinkLabel for site URL ---
$linkSite = New-Object System.Windows.Forms.LinkLabel
$linkSite.Text = ""
$linkSite.Location = "10,735"
$linkSite.Size = "850,22"
$linkSite.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Underline)
$linkSite.Visible = $false
$form.Controls.Add($linkSite)

# --- GUI Logic ---
function SetCertMode {
    param($mode)
    if ($mode -eq "existing") {
        $btnCreateCert.Enabled = $false
        $lblCertCN.Enabled = $false
        $txtCertCN.Enabled = $false
        $txtCertPath.Enabled = $true
        $btnBrowseCert.Enabled = $true
        $txtCertPass.Enabled = $true
        $grpKMS.Enabled = $true
        $btnProvision.Enabled = $true
    } elseif ($mode -eq "new") {
        $btnCreateCert.Enabled = $true
        $lblCertCN.Enabled = $true
        $txtCertCN.Enabled = $true
        $txtCertPath.Enabled = $true
        $btnBrowseCert.Enabled = $true
        $txtCertPass.Enabled = $true
        $grpKMS.Enabled = $false
        $btnProvision.Enabled = $false
    }
}
$rbExistingCert.Add_Click({ SetCertMode "existing" })
$rbNewCert.Add_Click({ SetCertMode "new" })
SetCertMode "existing"

$btnBrowseCert.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "PFX files (*.pfx)|*.pfx|All files (*.*)|*.*"
    if ($dialog.ShowDialog() -eq "OK") { $txtCertPath.Text = $dialog.FileName }
})

$btnDeptAdd.Add_Click({
    $newDept = $txtDeptAdd.Text.Trim()
    if ($newDept -and -not $lstDept.Items.Contains($newDept)) {
        $lstDept.Items.Add($newDept)
        $txtDeptAdd.Text = ""
    }
})
$btnDeptRemove.Add_Click({
    $idx = $lstDept.SelectedIndex
    if ($idx -ge 0) { $lstDept.Items.RemoveAt($idx) }
})

$btnCreateCert.Add_Click({
    $txtStatus.Text = ""
    $certPath = $txtCertPath.Text
    $certPass = $txtCertPass.Text
    $certCN = $txtCertCN.Text
    try {
        $SecurePassword = ConvertTo-SecureString $certPass -AsPlainText -Force
        $Cert = New-SelfSignedCertificate -Type Custom -Subject $certCN -KeySpec Signature `
            -KeyExportPolicy Exportable -HashAlgorithm SHA256 -KeyLength 2048 -CertStoreLocation "Cert:\CurrentUser\My" -NotAfter (Get-Date).AddYears(2)
        Export-PfxCertificate -Cert $Cert -FilePath $certPath -Password $SecurePassword
        $cerPath = [System.IO.Path]::ChangeExtension($certPath,".cer")
        Export-Certificate -Cert $Cert -FilePath $cerPath
        $txtStatus.AppendText("Certificate created and exported to $certPath`r`n")
        $txtStatus.AppendText("Upload $cerPath to your Azure App Registration > Certificates & Secrets.`r`n")
        $txtStatus.AppendText("Once done, continue below.`r`n")
        $rbExistingCert.Checked = $true
        SetCertMode "existing"
        $rbNewCert.Enabled = $false
    } catch {
        $txtStatus.AppendText("Error creating certificate: $_`r`n")
    }
})

$btnProvision.Add_Click({
    $txtStatus.Text = ""
    $linkSite.Text = ""
    $linkSite.Visible = $false

    $Tenant = $controls["txtTenant"].Text
    $TenantName = $controls["txtTenantName"].Text
    $ClientId = $controls["txtClientId"].Text
    $CertificatePath = $txtCertPath.Text
    $CertPassword = $txtCertPass.Text
    $SiteTitle = $controls["txtSiteTitle"].Text
    $SiteShort = $controls["txtSiteShort"].Text
    $OwnerEmail = $controls["txtOwnerEmail"].Text
    $Departments = @()
    for ($i=0; $i -lt $lstDept.Items.Count; $i++) { $Departments += $lstDept.Items[$i] }
    $SiteUrl = "https://$TenantName.sharepoint.com/sites/$SiteShort"
    $AdminUrl = "https://$TenantName-admin.sharepoint.com"
    $txtStatus.AppendText("Connecting to SharePoint...`r`n")
    try {
        Import-Module PnP.PowerShell -Force
        Connect-PnPOnline -Url $AdminUrl -ClientId $ClientId -Tenant $Tenant -CertificatePath $CertificatePath -CertificatePassword (ConvertTo-SecureString $CertPassword -AsPlainText -Force)
        $txtStatus.AppendText("Connected to $AdminUrl`r`n")
        $site = Get-PnPTenantSite -Url $SiteUrl -ErrorAction SilentlyContinue
        if (-not $site) {
            $txtStatus.AppendText("Creating site $SiteUrl...`r`n")
            New-PnPTenantSite -Title $SiteTitle -Url $SiteUrl -Owner $OwnerEmail -TimeZone 2 -Template "STS#3"
            $maxAttempts = 9; $attempt = 1
            do {
                Start-Sleep -Seconds 10
                $site = Get-PnPTenantSite -Url $SiteUrl -ErrorAction SilentlyContinue
                $attempt++
            } while ((-not $site) -and ($attempt -le $maxAttempts))
            if (-not $site) {
                $txtStatus.AppendText("Site creation timed out.`r`n")
                Disconnect-PnPOnline
                return
            }
            $txtStatus.AppendText("Site created and available.`r`n")
        } else {
            $txtStatus.AppendText("Site already exists: $SiteUrl`r`n")
        }
        Disconnect-PnPOnline
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $Tenant -CertificatePath $CertificatePath -CertificatePassword (ConvertTo-SecureString $CertPassword -AsPlainText -Force)
        $LibraryTitle = $SiteTitle
        if (-not (Get-PnPList -Identity $LibraryTitle -ErrorAction SilentlyContinue)) {
            $txtStatus.AppendText("Creating document library: $LibraryTitle`r`n")
            New-PnPList -Title $LibraryTitle -Template DocumentLibrary
        } else {
            $txtStatus.AppendText("Library exists: $LibraryTitle`r`n")
        }
        Set-PnPList -Identity $LibraryTitle -EnableVersioning $true -MajorVersions 50

        # --- Add metadata columns (NO Approvers column) ---
        $defaultMetadata = @(
            @{DisplayName="Department"; InternalName="Department"; Type="Choice"; Choices=$Departments; Required=$true},
            @{DisplayName="Document Type"; InternalName="DocumentType"; Type="Choice"; Choices=@("Policy","Procedure","Training","Lessons Learned","Client"); Required=$true},
            @{DisplayName="Approval Status"; InternalName="ApprovalStatus"; Type="Choice"; Choices=@("Pending","Approved","Rejected"); DefaultValue="Pending"; Required=$true},
            @{DisplayName="Approval Comments"; InternalName="ApprovalComments"; Type="Note"; Required=$false},
            @{DisplayName="Owner"; InternalName="Owner"; Type="User"; Required=$true},
            @{DisplayName="Review Date"; InternalName="ReviewDate"; Type="DateTime"; Required=$false}
        )
        foreach ($field in $defaultMetadata) {
            $existingField = Get-PnPField -List $LibraryTitle -Identity $field.InternalName -ErrorAction SilentlyContinue
            if (-not $existingField) {
                $params = @{
                    List = $LibraryTitle
                    DisplayName = $field.DisplayName
                    InternalName = $field.InternalName
                    Type = $field.Type
                    Required = $field.Required
                }
                if ($field.Type -eq "Choice" -or $field.Type -eq "MultiChoice") {
                    $params.Choices = $field.Choices
                }
                Add-PnPField @params
                if ($field.ContainsKey("DefaultValue")) {
                    Set-PnPField -List $LibraryTitle -Identity $field.InternalName -Values @{DefaultValue = $field.DefaultValue}
                }
                $txtStatus.AppendText("Added column: $($field.DisplayName)`r`n")
            }
        }

        $builtinColumns = @("Title", "Created", "Modified", "Author", "Editor", "Version")
        $customColumns = $defaultMetadata | ForEach-Object { $_.InternalName }
        $fields = $builtinColumns + $customColumns
        try { Remove-PnPView -List $LibraryTitle -Identity "All Documents" -Force -ErrorAction SilentlyContinue } catch {}
        Add-PnPView -List $LibraryTitle -Title "All Columns" -Fields $fields -SetAsDefault

        Set-PnPField -List $LibraryTitle -Identity "ApprovalStatus" -Values @{ReadOnlyField = $true}
        Set-PnPHomePage -RootFolderRelativeUrl $LibraryTitle
        $txtStatus.AppendText("Knowledge Management System provisioned and home page set!`r`n")
        $txtStatus.AppendText("Open your KMS site here: $SiteUrl`r`n")
        $linkSite.Text = $SiteUrl
        $linkSite.Visible = $true
    } catch {
        $txtStatus.AppendText("Error: $_`r`n")
    }
})

$linkSite.Add_LinkClicked({
    if ($linkSite.Text) {
        Start-Process $linkSite.Text
    }
})

[void]$form.ShowDialog()
