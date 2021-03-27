<#
.SYNOPSIS
PowerShell Tool for renaming network computers with graphical user interface.

.DESCRIPTION
PowerShell Tool for renaming network computers with graphical user interface that can be used to rename single computer or to
perform bulk rename drawing data from a `.csv` file. The `.csv` file must have OldName and NewName columns for tool to work.

.NOTES
Version:        1.0
Author:         Zoran Jankov
#>

$SearchBase = "OU=Korisnici,OU=Centrala,DC=uni,DC=net"

Import-Module "$PSScriptRoot\Modules\Write-Log.psm1"
$CSVFile = $null

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
[System.Windows.Forms.Application]::EnableVisualStyles()

$MainForm                        = New-Object system.Windows.Forms.Form
$MainForm.ClientSize             = New-Object System.Drawing.Point(670,295)
$MainForm.Text                   = "Rename Computers Tool"
$MainForm.TopMost                = $true
$MainForm.FormBorderStyle        = 'Fixed3D'
$MainForm.MaximizeBox            = $false
$MainForm.ShowIcon               = $false

$ComputerNameLabel               = New-Object system.Windows.Forms.Label
$ComputerNameLabel.Text          = "Computer"
$ComputerNameLabel.AutoSize      = $true
$ComputerNameLabel.Width         = 25
$ComputerNameLabel.Height        = 10
$ComputerNameLabel.Location      = New-Object System.Drawing.Point(25,70)
$ComputerNameLabel.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$NewNameLabel                    = New-Object system.Windows.Forms.Label
$NewNameLabel.Text               = "New Name"
$NewNameLabel.AutoSize           = $true
$NewNameLabel.Width              = 25
$NewNameLabel.Height             = 10
$NewNameLabel.Location           = New-Object System.Drawing.Point(25,110)
$NewNameLabel.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$ComputerNameComboBox            = New-Object system.Windows.Forms.ComboBox
$ComputerNameComboBox.Width      = 380
$ComputerNameComboBox.Height     = 25
$ComputerNameComboBox.Location   = New-Object System.Drawing.Point(120,65)
$ComputerNameComboBox.Font       = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$ComputerNameComboBox.AutoCompleteMode = 'SuggestAppend'
$ComputerNameComboBox.AutoCompleteSource = 'ListItems'

$NewNameTextBox                  = New-Object system.Windows.Forms.TextBox
$NewNameTextBox.Multiline        = $false
$NewNameTextBox.Width            = 380
$NewNameTextBox.Height           = 25
$NewNameTextBox.Location         = New-Object System.Drawing.Point(120,105)
$NewNameTextBox.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$SingleRenameButton              = New-Object system.Windows.Forms.Button
$SingleRenameButton.Text         = "Rename"
$SingleRenameButton.Width        = 115
$SingleRenameButton.Height       = 30
$SingleRenameButton.Location     = New-Object System.Drawing.Point(525,80)
$SingleRenameButton.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$SingleRenameLabel               = New-Object system.Windows.Forms.Label
$SingleRenameLabel.Text          = "Single Rename"
$SingleRenameLabel.AutoSize      = $true
$SingleRenameLabel.Width         = 25
$SingleRenameLabel.Height        = 10
$SingleRenameLabel.Location      = New-Object System.Drawing.Point(280,25)
$SingleRenameLabel.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$BulkRenameLabel                 = New-Object system.Windows.Forms.Label
$BulkRenameLabel.Text            = "Bulk Rename"
$BulkRenameLabel.AutoSize        = $true
$BulkRenameLabel.Width           = 25
$BulkRenameLabel.Height          = 10
$BulkRenameLabel.Location        = New-Object System.Drawing.Point(280,145)
$BulkRenameLabel.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$CSVFileLabel                    = New-Object system.Windows.Forms.Label
$CSVFileLabel.Text               = "CSV File"
$CSVFileLabel.AutoSize           = $true
$CSVFileLabel.Width              = 25
$CSVFileLabel.Height             = 10
$CSVFileLabel.Location           = New-Object System.Drawing.Point(25,210)
$CSVFileLabel.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$FilePathTextBox                 = New-Object system.Windows.Forms.TextBox
$FilePathTextBox.Multiline       = $false
$FilePathTextBox.Width           = 377
$FilePathTextBox.Height          = 25
$FilePathTextBox.Enabled         = $false
$FilePathTextBox.Location        = New-Object System.Drawing.Point(115,205)
$FilePathTextBox.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$SelectFileButton1               = New-Object system.Windows.Forms.Button
$SelectFileButton1.Text          = "Select File"
$SelectFileButton1.Width         = 120
$SelectFileButton1.Height        = 30
$SelectFileButton1.Location      = New-Object System.Drawing.Point(525,200)
$SelectFileButton1.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$BulkRenameButton                = New-Object system.Windows.Forms.Button
$BulkRenameButton.Text           = "Bulk Rename"
$BulkRenameButton.Width          = 160
$BulkRenameButton.Height         = 30
$BulkRenameButton.Location       = New-Object System.Drawing.Point(250,250)
$BulkRenameButton.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$MainForm.controls.AddRange(@(
    $ComputerNameLabel,
    $NewNameLabel,
    $ComputerNameComboBox,
    $NewNameTextBox,
    $SingleRenameButton,
    $SingleRenameLabel,
    $BulkRenameLabel,
    $CSVFileLabel,
    $FilePathTextBox,
    $SelectFileButton1,
    $BulkRenameButton
))

$Computers = Get-ADComputer -Filter {Enabled -eq $true} -SearchBase $SearchBase | Sort-Object Name
foreach ($Computer in $Computers) {
    $ComputerNameComboBox.Items.Add($Computer.Name);
}

$SingleRenameButton.Add_Click({
    $OldName = $ComputerNameComboBox.text
    $NewName = $NewNameTextBox.text
    try {
        Rename-Computer -ComputerName $OldName -NewName $NewName
    }
    catch {
        $Message ="Faild to rename $OldName computer `n" + $_.Exception
        break
    }
    $Message = "Successfully renamed " + $OldName + " computer to " + $NewName
    Write-Log -Message $Message
    [System.Windows.MessageBox]::Show($Message)
})
$SelectFileButton1.Add_Click({
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    $OpenFileDialog.Filter = "CSV Files (*.csv)| *.csv*"
    $OpenFileDialog.ShowDialog() | Out-Null
    $CSVFile = Get-Content -Path $OpenFileDialog.Filename
    $FilePathTextBox.Text = $OpenFileDialog.Filename
})

$BulkRenameButton.Add_Click({
    foreach ($Computer in $CSVFile) {
        try {
            Rename-Computer -ComputerName $OldName -NewName $NewName
        }
        catch {
            $Message ="Faild to rename " + $Computer.OldName + " computer `n" + $_.Exception
            continue
        }
        $Message = "Successfully renamed " + $Computer.OldName + " computer to " + $Computer.NewName
        Write-Log -Message $Message
    }
})

[void]$MainForm.ShowDialog()