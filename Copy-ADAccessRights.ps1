<#
.SYNOPSIS
    Copie les droits d'acces d'un utilisateur Active Directory a un autre

.DESCRIPTION
    Interface GUI pour scanner les dossiers partages et copier les permissions
    de l'utilisateur source vers l'utilisateur cible.
    - Auto-elevation en Administrateur si necessaire
    - Verification automatique des prerequis (AD, Windows Forms)

.VERSION
    1.2 - Auto-elevation + gestion robuste des erreurs
#>

# =====================================================
# AUTO-ELEVATION EN ADMINISTRATEUR
# =====================================================
$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal   = New-Object Security.Principal.WindowsPrincipal($currentUser)
$isAdmin     = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "Relancement en mode Administrateur..." -ForegroundColor Yellow
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`""
    Start-Process powershell.exe -Verb RunAs -ArgumentList $arguments
    exit
}

# =====================================================
# VERIFICATION DE LA POLITIQUE D'EXECUTION
# =====================================================
$execPolicy = Get-ExecutionPolicy -Scope CurrentUser
if ($execPolicy -eq 'Restricted' -or $execPolicy -eq 'AllSigned') {
    try {
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
        Write-Host "ExecutionPolicy mise a jour : RemoteSigned" -ForegroundColor Green
    }
    catch {
        Write-Host "Impossible de modifier ExecutionPolicy : $_" -ForegroundColor Red
    }
}

# =====================================================
# VERIFICATION ET IMPORT DU MODULE ACTIVEDIRECTORY
# =====================================================
$adAvailable = $false
if (Get-Module -ListAvailable -Name ActiveDirectory) {
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        $adAvailable = $true
        Write-Host "Module ActiveDirectory charge." -ForegroundColor Green
    }
    catch {
        Write-Host "Erreur chargement ActiveDirectory : $_" -ForegroundColor Red
    }
}
else {
    Write-Host "Module ActiveDirectory introuvable. Tentative installation RSAT..." -ForegroundColor Yellow
    try {
        # Windows 10/11
        Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 -ErrorAction Stop
        Import-Module ActiveDirectory -ErrorAction Stop
        $adAvailable = $true
        Write-Host "RSAT AD installe et charge." -ForegroundColor Green
    }
    catch {
        try {
            # Windows Server
            Install-WindowsFeature RSAT-AD-PowerShell -ErrorAction Stop
            Import-Module ActiveDirectory -ErrorAction Stop
            $adAvailable = $true
            Write-Host "RSAT AD (Server) installe et charge." -ForegroundColor Green
        }
        catch {
            Write-Host "Impossible d'installer ActiveDirectory. Mode limite." -ForegroundColor Red
        }
    }
}

# =====================================================
# CHARGEMENT DE WINDOWS FORMS
# =====================================================
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing       -ErrorAction Stop
    [System.Windows.Forms.Application]::EnableVisualStyles()
    Write-Host "Windows Forms charge." -ForegroundColor Green
}
catch {
    Write-Host "ERREUR CRITIQUE : Impossible de charger Windows Forms : $_" -ForegroundColor Red
    Write-Host "Appuyez sur une touche pour quitter..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# =====================================================
# VARIABLES GLOBALES
# =====================================================
$script:ScanResults = @()
$script:IsScanning  = $false

# =====================================================
# FONCTIONS UTILITAIRES
# =====================================================

function Get-AllADUsers {
    if (-not $adAvailable) { return @() }
    try {
        $users = Get-ADUser -Filter * -Properties SamAccountName, DisplayName -ErrorAction Stop |
            Select-Object @{Name="DisplayName";    Expression={ if ($_.DisplayName) { $_.DisplayName } else { $_.SamAccountName } }},
                          @{Name="SamAccountName"; Expression={ $_.SamAccountName }} |
            Sort-Object DisplayName
        return $users
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Erreur recuperation utilisateurs AD :`n$_",
            "Erreur AD",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return @()
    }
}

# =====================================================
# CONSTRUCTION DE L'INTERFACE GUI
# =====================================================

# Fenetre principale
$form               = New-Object System.Windows.Forms.Form
$form.Text          = "Copie des Droits d'Acces - Active Directory"
$form.Size          = New-Object System.Drawing.Size(920, 720)
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.BackColor     = [System.Drawing.Color]::White
$form.MinimumSize   = New-Object System.Drawing.Size(920, 720)
try { $form.Icon    = [System.Drawing.SystemIcons]::Shield } catch {}

# Panel principal
$mainPanel         = New-Object System.Windows.Forms.Panel
$mainPanel.Dock    = [System.Windows.Forms.DockStyle]::Fill
$mainPanel.Padding = New-Object System.Windows.Forms.Padding(20)
$form.Controls.Add($mainPanel)

# Titre
$titleLabel          = New-Object System.Windows.Forms.Label
$titleLabel.Text     = "Copie des Droits d'Acces - Active Directory"
$titleLabel.Font     = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$titleLabel.Size     = New-Object System.Drawing.Size(870, 40)
$titleLabel.Location = New-Object System.Drawing.Point(0, 0)
$mainPanel.Controls.Add($titleLabel)

# Bandeau avertissement si AD indisponible
if (-not $adAvailable) {
    $warnLabel           = New-Object System.Windows.Forms.Label
    $warnLabel.Text      = "  ATTENTION : Module ActiveDirectory non disponible - verifiez l'installation des outils RSAT."
    $warnLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $warnLabel.ForeColor = [System.Drawing.Color]::White
    $warnLabel.BackColor = [System.Drawing.Color]::FromArgb(200, 50, 50)
    $warnLabel.Size      = New-Object System.Drawing.Size(870, 26)
    $warnLabel.Location  = New-Object System.Drawing.Point(0, 44)
    $warnLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $mainPanel.Controls.Add($warnLabel)
}

# GroupBox Selection
$selectionGroupBox          = New-Object System.Windows.Forms.GroupBox
$selectionGroupBox.Text     = "Selection des Utilisateurs"
$selectionGroupBox.Font     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$selectionGroupBox.Size     = New-Object System.Drawing.Size(870, 215)
$selectionGroupBox.Location = New-Object System.Drawing.Point(0, 54)
$mainPanel.Controls.Add($selectionGroupBox)

# Label Source
$sourceLabel          = New-Object System.Windows.Forms.Label
$sourceLabel.Text     = "Utilisateur Source :"
$sourceLabel.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
$sourceLabel.AutoSize = $true
$sourceLabel.Location = New-Object System.Drawing.Point(20, 32)
$selectionGroupBox.Controls.Add($sourceLabel)

# ComboBox Source
$sourceComboBox                    = New-Object System.Windows.Forms.ComboBox
$sourceComboBox.Size               = New-Object System.Drawing.Size(370, 30)
$sourceComboBox.Location           = New-Object System.Drawing.Point(20, 55)
$sourceComboBox.DropDownStyle      = [System.Windows.Forms.ComboBoxStyle]::DropDown
$sourceComboBox.AutoCompleteMode   = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
$sourceComboBox.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
$sourceComboBox.Font               = New-Object System.Drawing.Font("Segoe UI", 9)
$selectionGroupBox.Controls.Add($sourceComboBox)

# Label Cible
$targetLabel          = New-Object System.Windows.Forms.Label
$targetLabel.Text     = "Utilisateur Cible :"
$targetLabel.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
$targetLabel.AutoSize = $true
$targetLabel.Location = New-Object System.Drawing.Point(490, 32)
$selectionGroupBox.Controls.Add($targetLabel)

# ComboBox Cible
$targetComboBox                    = New-Object System.Windows.Forms.ComboBox
$targetComboBox.Size               = New-Object System.Drawing.Size(370, 30)
$targetComboBox.Location           = New-Object System.Drawing.Point(490, 55)
$targetComboBox.DropDownStyle      = [System.Windows.Forms.ComboBoxStyle]::DropDown
$targetComboBox.AutoCompleteMode   = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
$targetComboBox.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
$targetComboBox.Font               = New-Object System.Drawing.Font("Segoe UI", 9)
$selectionGroupBox.Controls.Add($targetComboBox)

# Bouton Rechercher
$searchButton           = New-Object System.Windows.Forms.Button
$searchButton.Text      = "RECHERCHER LES DROITS"
$searchButton.Size      = New-Object System.Drawing.Size(210, 38)
$searchButton.Location  = New-Object System.Drawing.Point(20, 112)
$searchButton.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
$searchButton.ForeColor = [System.Drawing.Color]::White
$searchButton.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$searchButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$searchButton.Cursor    = [System.Windows.Forms.Cursors]::Hand
$selectionGroupBox.Controls.Add($searchButton)

# Bouton Details
$detailsButton           = New-Object System.Windows.Forms.Button
$detailsButton.Text      = "VERIFIER DETAILS"
$detailsButton.Size      = New-Object System.Drawing.Size(210, 38)
$detailsButton.Location  = New-Object System.Drawing.Point(245, 112)
$detailsButton.BackColor = [System.Drawing.Color]::FromArgb(100, 149, 237)
$detailsButton.ForeColor = [System.Drawing.Color]::White
$detailsButton.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$detailsButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$detailsButton.Cursor    = [System.Windows.Forms.Cursors]::Hand
$detailsButton.Enabled   = $false
$selectionGroupBox.Controls.Add($detailsButton)

# Bouton Copier
$copyButton           = New-Object System.Windows.Forms.Button
$copyButton.Text      = "COPIER LES DROITS"
$copyButton.Size      = New-Object System.Drawing.Size(210, 38)
$copyButton.Location  = New-Object System.Drawing.Point(470, 112)
$copyButton.BackColor = [System.Drawing.Color]::FromArgb(34, 139, 34)
$copyButton.ForeColor = [System.Drawing.Color]::White
$copyButton.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$copyButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$copyButton.Cursor    = [System.Windows.Forms.Cursors]::Hand
$copyButton.Enabled   = $false
$selectionGroupBox.Controls.Add($copyButton)

# Barre de progression
$progressBar          = New-Object System.Windows.Forms.ProgressBar
$progressBar.Size     = New-Object System.Drawing.Size(830, 22)
$progressBar.Location = New-Object System.Drawing.Point(20, 162)
$progressBar.Style    = [System.Windows.Forms.ProgressBarStyle]::Continuous
$selectionGroupBox.Controls.Add($progressBar)

# Label Statut
$statusLabel           = New-Object System.Windows.Forms.Label
$statusLabel.Text      = "Pret"
$statusLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$statusLabel.AutoSize  = $true
$statusLabel.Location  = New-Object System.Drawing.Point(20, 190)
$statusLabel.ForeColor = [System.Drawing.Color]::Green
$selectionGroupBox.Controls.Add($statusLabel)

# Panel Details
$detailsPanel             = New-Object System.Windows.Forms.Panel
$detailsPanel.Visible     = $false
$detailsPanel.Size        = New-Object System.Drawing.Size(870, 255)
$detailsPanel.Location    = New-Object System.Drawing.Point(0, 278)
$detailsPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$detailsPanel.BackColor   = [System.Drawing.Color]::FromArgb(240, 248, 255)
$mainPanel.Controls.Add($detailsPanel)

$detailsTitleLabel          = New-Object System.Windows.Forms.Label
$detailsTitleLabel.Text     = "Dossiers trouves :"
$detailsTitleLabel.Font     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$detailsTitleLabel.AutoSize = $true
$detailsTitleLabel.Location = New-Object System.Drawing.Point(10, 10)
$detailsPanel.Controls.Add($detailsTitleLabel)

$detailsGridView                       = New-Object System.Windows.Forms.DataGridView
$detailsGridView.Size                  = New-Object System.Drawing.Size(848, 215)
$detailsGridView.Location              = New-Object System.Drawing.Point(10, 32)
$detailsGridView.AutoGenerateColumns   = $false
$detailsGridView.AllowUserToAddRows    = $false
$detailsGridView.AllowUserToDeleteRows = $false
$detailsGridView.ReadOnly              = $true
$detailsGridView.RowHeadersVisible     = $false
$detailsGridView.BackgroundColor       = [System.Drawing.Color]::White
$detailsGridView.Font                  = New-Object System.Drawing.Font("Segoe UI", 9)

foreach ($col in @(
    @{Name="Dossier"; Width=140},
    @{Name="Chemin";  Width=380},
    @{Name="Droits";  Width=180},
    @{Name="Type";    Width=110}
)) {
    $c            = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $c.Name       = $col.Name
    $c.HeaderText = $col.Name
    $c.Width      = $col.Width
    $c.ReadOnly   = $true
    $detailsGridView.Columns.Add($c) | Out-Null
}
$detailsPanel.Controls.Add($detailsGridView)

# Panel Resultats
$resultsPanel             = New-Object System.Windows.Forms.Panel
$resultsPanel.Visible     = $false
$resultsPanel.Size        = New-Object System.Drawing.Size(870, 255)
$resultsPanel.Location    = New-Object System.Drawing.Point(0, 278)
$resultsPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$resultsPanel.BackColor   = [System.Drawing.Color]::FromArgb(245, 245, 245)
$mainPanel.Controls.Add($resultsPanel)

$resultsTitleLabel          = New-Object System.Windows.Forms.Label
$resultsTitleLabel.Text     = "Resultats de la copie :"
$resultsTitleLabel.Font     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$resultsTitleLabel.AutoSize = $true
$resultsTitleLabel.Location = New-Object System.Drawing.Point(10, 10)
$resultsPanel.Controls.Add($resultsTitleLabel)

$resultsTextBox           = New-Object System.Windows.Forms.RichTextBox
$resultsTextBox.Size      = New-Object System.Drawing.Size(848, 215)
$resultsTextBox.Location  = New-Object System.Drawing.Point(10, 32)
$resultsTextBox.ReadOnly  = $true
$resultsTextBox.Font      = New-Object System.Drawing.Font("Consolas", 9)
$resultsTextBox.BackColor = [System.Drawing.Color]::White
$resultsPanel.Controls.Add($resultsTextBox)

# =====================================================
# EVENEMENTS
# =====================================================

$searchButton.Add_Click({

    if ([string]::IsNullOrWhiteSpace($sourceComboBox.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Veuillez selectionner un utilisateur source.",
            "Avertissement",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # FIX Bug 2 : verification SelectedItem non null
    $selectedSource = $sourceComboBox.SelectedItem
    if ($null -eq $selectedSource) {
        [System.Windows.Forms.MessageBox]::Show(
            "Veuillez choisir l'utilisateur source dans la liste deroulante (ne pas saisir manuellement).",
            "Avertissement",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $script:IsScanning     = $true
    $searchButton.Enabled  = $false
    $detailsButton.Enabled = $false
    $copyButton.Enabled    = $false
    $progressBar.Value     = 0
    $resultsPanel.Visible  = $false
    $detailsPanel.Visible  = $false
    $statusLabel.ForeColor = [System.Drawing.Color]::Blue
    $statusLabel.Text      = "Scan en cours..."

    $sourceUserName = $selectedSource.SamAccountName

    # FIX Bug 1 : Import-Module AD dans le job
    # FIX Bug 3 : regex corrigee pour filtrer ADMIN$, IPC$
    # FIX Bug 4 : suppression du progressMonitor defectueux
    $job = Start-Job -ScriptBlock {
        param($UserName)
        Import-Module ActiveDirectory -ErrorAction SilentlyContinue

        try {
            $shares = Get-SmbShare -ErrorAction Stop |
                Where-Object { $_.Name -notmatch '^\$|ADMIN\$|IPC\$|print\$' }
        }
        catch { return @() }

        $results      = @()
        $totalShares  = @($shares).Count
        $currentIndex = 0

        foreach ($share in $shares) {
            $currentIndex++
            Write-Progress -Activity "Scan" -Status $share.Name `
                           -PercentComplete (($currentIndex / $totalShares) * 100)
            try {
                if (-not (Test-Path $share.Path)) { continue }
                $acl = Get-Acl -Path $share.Path -ErrorAction Continue
                foreach ($access in $acl.Access) {
                    if ($access.IdentityReference.Value -match $UserName) {
                        $results += [PSCustomObject]@{
                            Dossier = $share.Name
                            Chemin  = $share.Path
                            Droits  = $access.FileSystemRights.ToString()
                            Acces   = $access.AccessControlType.ToString()
                            Herite  = $access.IsInherited
                        }
                        break
                    }
                }
            }
            catch {}
        }
        return $results
    } -ArgumentList $sourceUserName

    $script:ScanResults = Receive-Job -Job $job -Wait
    Remove-Job -Job $job -Force -ErrorAction SilentlyContinue

    $progressBar.Value = 100

    if (@($script:ScanResults).Count -gt 0) {
        $statusLabel.ForeColor = [System.Drawing.Color]::Green
        $statusLabel.Text      = "Scan termine : $(@($script:ScanResults).Count) dossier(s) trouve(s)"

        $detailsGridView.Rows.Clear()
        foreach ($r in $script:ScanResults) {
            $detailsGridView.Rows.Add($r.Dossier, $r.Chemin, $r.Droits, $r.Acces) | Out-Null
        }
        $detailsButton.Enabled = $true
        $detailsPanel.Visible  = $true

        if (-not [string]::IsNullOrWhiteSpace($targetComboBox.Text)) {
            $copyButton.Enabled = $true
        }
    }
    else {
        $statusLabel.ForeColor = [System.Drawing.Color]::Orange
        $statusLabel.Text      = "Aucun dossier trouve pour cet utilisateur."
    }

    $searchButton.Enabled = $true
    $script:IsScanning    = $false
})

$detailsButton.Add_Click({
    if ($detailsPanel.Visible) {
        $detailsPanel.Visible = $false
        $detailsButton.Text   = "VERIFIER DETAILS"
    }
    else {
        $resultsPanel.Visible = $false
        $detailsPanel.Visible = $true
        $detailsButton.Text   = "MASQUER DETAILS"
    }
})

$copyButton.Add_Click({
    if (@($script:ScanResults).Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Aucun dossier a traiter. Lancez d'abord une recherche.",
            "Avertissement",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # FIX Bug 2 : verification SelectedItem cible non null
    $selectedTarget = $targetComboBox.SelectedItem
    if ($null -eq $selectedTarget) {
        [System.Windows.Forms.MessageBox]::Show(
            "Veuillez choisir l'utilisateur cible dans la liste deroulante (ne pas saisir manuellement).",
            "Avertissement",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Copier les droits de`n  $($sourceComboBox.Text)`nvers`n  $($targetComboBox.Text)`n`nSur $(@($script:ScanResults).Count) dossier(s) ?",
        "Confirmation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $copyButton.Enabled    = $false
    $searchButton.Enabled  = $false
    $detailsButton.Enabled = $false
    $progressBar.Value     = 0
    $statusLabel.ForeColor = [System.Drawing.Color]::Blue
    $statusLabel.Text      = "Copie des droits en cours..."

    $sourceUserName = $sourceComboBox.SelectedItem.SamAccountName
    $targetUserName = $selectedTarget.SamAccountName

    # FIX Bug 1 : Import-Module AD dans le job
    # FIX Bug 3 : regex corrigee
    # FIX Bug 4 : suppression progressMonitor defectueux
    $copyJob = Start-Job -ScriptBlock {
        param($SourceUser, $TargetUser, $Folders)
        Import-Module ActiveDirectory -ErrorAction SilentlyContinue

        $results      = @()
        $successCount = 0
        $errorCount   = 0
        $totalFolders = @($Folders).Count
        $currentIndex = 0

        foreach ($folder in $Folders) {
            $currentIndex++
            Write-Progress -Activity "Copie" -Status $folder.Dossier `
                           -PercentComplete (($currentIndex / $totalFolders) * 100)
            try {
                $aclPath = $folder.Chemin
                if (-not (Test-Path $aclPath)) {
                    $results += "ERREUR $aclPath - Chemin inaccessible"
                    $errorCount++
                    continue
                }

                $acl        = Get-Acl -Path $aclPath
                $sourceRule = $null
                foreach ($access in $acl.Access) {
                    if ($access.IdentityReference.Value -match $SourceUser) {
                        $sourceRule = $access
                        break
                    }
                }

                if ($sourceRule) {
                    try {
                        $targetAD      = Get-ADUser -Identity $TargetUser -ErrorAction Stop
                        $targetAccount = New-Object System.Security.Principal.SecurityIdentifier($targetAD.SID)

                        # Verifier si la regle existe deja
                        $ruleExists = $false
                        foreach ($access in $acl.Access) {
                            if ($access.IdentityReference.Value -like "*$TargetUser" -or
                                $access.IdentityReference.Value -eq $targetAccount.Value) {
                                $ruleExists = $true
                                break
                            }
                        }

                        if (-not $ruleExists) {
                            $newRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                                $targetAccount,
                                $sourceRule.FileSystemRights,
                                $sourceRule.InheritanceFlags,
                                $sourceRule.PropagationFlags,
                                $sourceRule.AccessControlType
                            )
                            $acl.AddAccessRule($newRule)
                        }

                        Set-Acl -Path $aclPath -AclObject $acl -ErrorAction Stop
                        $results += "OK $aclPath - Droits copies avec succes"
                        $successCount++
                    }
                    catch {
                        $results += "ERREUR $aclPath - $($_.Exception.Message)"
                        $errorCount++
                    }
                }
                else {
                    $results += "ERREUR $aclPath - Aucune regle source trouvee"
                    $errorCount++
                }
            }
            catch {
                $results += "ERREUR $($folder.Chemin) - $($_.Exception.Message)"
                $errorCount++
            }
        }

        return @{ Results = $results; Success = $successCount; Errors = $errorCount }
    } -ArgumentList $sourceUserName, $targetUserName, $script:ScanResults

    $copyResult = Receive-Job -Job $copyJob -Wait
    Remove-Job -Job $copyJob -Force -ErrorAction SilentlyContinue

    # Affichage des resultats
    $sep  = "=" * 65
    $sep2 = "-" * 65
    $resultsTextBox.Clear()
    $resultsTextBox.AppendText("$sep`n")
    $resultsTextBox.AppendText("  RESUME DE LA COPIE`n")
    $resultsTextBox.AppendText("$sep`n`n")
    $resultsTextBox.AppendText("  Source : $($sourceComboBox.Text)`n")
    $resultsTextBox.AppendText("  Cible  : $($targetComboBox.Text)`n")
    $resultsTextBox.AppendText("`n$sep2`n`n")
    foreach ($r in $copyResult.Results) {
        $resultsTextBox.AppendText("  $r`n")
    }
    $resultsTextBox.AppendText("`n$sep2`n")
    $resultsTextBox.AppendText("  Succes : $($copyResult.Success)   |   Erreurs : $($copyResult.Errors)`n")
    $resultsTextBox.AppendText("$sep`n")

    $progressBar.Value = 100
    if ($copyResult.Errors -eq 0) {
        $statusLabel.ForeColor = [System.Drawing.Color]::Green
        $statusLabel.Text      = "Copie reussie : $($copyResult.Success) dossier(s)"
    }
    else {
        $statusLabel.ForeColor = [System.Drawing.Color]::Orange
        $statusLabel.Text      = "Partiel : $($copyResult.Success) succes, $($copyResult.Errors) erreur(s)"
    }

    $detailsPanel.Visible  = $false
    $resultsPanel.Visible  = $true
    $copyButton.Enabled    = $true
    $searchButton.Enabled  = $true
    $detailsButton.Enabled = $true
})

$targetComboBox.Add_TextChanged({
    if (@($script:ScanResults).Count -gt 0 -and
        -not [string]::IsNullOrWhiteSpace($targetComboBox.Text)) {
        $copyButton.Enabled = $true
    }
    else {
        $copyButton.Enabled = $false
    }
})

# =====================================================
# INITIALISATION - CHARGEMENT DES UTILISATEURS AD
# =====================================================
$statusLabel.Text      = "Chargement des utilisateurs AD..."
$statusLabel.ForeColor = [System.Drawing.Color]::Blue
[System.Windows.Forms.Application]::DoEvents()

$adUsers = Get-AllADUsers

if ($adUsers.Count -eq 0) {
    $statusLabel.Text      = if ($adAvailable) { "Aucun utilisateur trouve dans l'AD." } else { "Module AD indisponible." }
    $statusLabel.ForeColor = [System.Drawing.Color]::Red
}
else {
    foreach ($user in $adUsers) {
        $sourceComboBox.Items.Add($user) | Out-Null
        $targetComboBox.Items.Add($user) | Out-Null
    }
    $sourceComboBox.DisplayMember = "DisplayName"
    $sourceComboBox.ValueMember   = "SamAccountName"
    $targetComboBox.DisplayMember = "DisplayName"
    $targetComboBox.ValueMember   = "SamAccountName"

    $statusLabel.Text      = "Pret - $($adUsers.Count) utilisateur(s) charge(s)"
    $statusLabel.ForeColor = [System.Drawing.Color]::Green
}

# =====================================================
# LANCEMENT DE LA FENETRE
# =====================================================
try {
    [System.Windows.Forms.Application]::Run($form)
}
catch {
    Write-Host "Erreur fatale affichage fenetre : $_" -ForegroundColor Red
    Write-Host "Appuyez sur une touche pour quitter..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
finally {
    try { $form.Dispose() } catch {}
}