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

Import-Module "$PSScriptRoot\Modules\Write-Log.psm1"
$Credential = Get-Credential
$CSVFile = $null

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
[System.Windows.Forms.Application]::EnableVisualStyles()

$RenameComputersForm             = New-Object system.Windows.Forms.Form
$RenameComputersForm.ClientSize  = New-Object System.Drawing.Point(669,301)
$RenameComputersForm.Text        = "Rename Computers Tool"
$RenameComputersForm.TopMost     = $true

$OldNameLabel                    = New-Object system.Windows.Forms.Label
$OldNameLabel.Text               = "Old Name"
$OldNameLabel.AutoSize           = $true
$OldNameLabel.Width              = 25
$OldNameLabel.Height             = 10
$OldNameLabel.Location           = New-Object System.Drawing.Point(25,70)
$OldNameLabel.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$NewNameLabel                    = New-Object system.Windows.Forms.Label
$NewNameLabel.Text               = "New Name"
$NewNameLabel.AutoSize           = $true
$NewNameLabel.Width              = 25
$NewNameLabel.Height             = 10
$NewNameLabel.Location           = New-Object System.Drawing.Point(25,105)
$NewNameLabel.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$OldNameTextBox                  = New-Object system.Windows.Forms.TextBox
$OldNameTextBox.Multiline        = $false
$OldNameTextBox.Width            = 377
$OldNameTextBox.Height           = 25
$OldNameTextBox.Location         = New-Object System.Drawing.Point(117,64)
$OldNameTextBox.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$NewNameTextBox                  = New-Object system.Windows.Forms.TextBox
$NewNameTextBox.Multiline        = $false
$NewNameTextBox.Width            = 377
$NewNameTextBox.Height           = 25
$NewNameTextBox.Location         = New-Object System.Drawing.Point(117,100)
$NewNameTextBox.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$SingleRenameButton              = New-Object system.Windows.Forms.Button
$SingleRenameButton.Text         = "Rename"
$SingleRenameButton.Width        = 114
$SingleRenameButton.Height       = 30
$SingleRenameButton.Location     = New-Object System.Drawing.Point(527,76)
$SingleRenameButton.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$SingleRenameLabel               = New-Object system.Windows.Forms.Label
$SingleRenameLabel.Text          = "Single Rename"
$SingleRenameLabel.AutoSize      = $true
$SingleRenameLabel.Width         = 25
$SingleRenameLabel.Height        = 10
$SingleRenameLabel.Location      = New-Object System.Drawing.Point(279,21)
$SingleRenameLabel.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$BulkRenameLabel                 = New-Object system.Windows.Forms.Label
$BulkRenameLabel.Text            = "Bulk Rename"
$BulkRenameLabel.AutoSize        = $true
$BulkRenameLabel.Width           = 25
$BulkRenameLabel.Height          = 10
$BulkRenameLabel.Location        = New-Object System.Drawing.Point(287,145)
$BulkRenameLabel.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$CSVFileLabel                    = New-Object system.Windows.Forms.Label
$CSVFileLabel.Text               = "CSV File"
$CSVFileLabel.AutoSize           = $true
$CSVFileLabel.Width              = 25
$CSVFileLabel.Height             = 10
$CSVFileLabel.Location           = New-Object System.Drawing.Point(25,197)
$CSVFileLabel.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$FilePathTextBox                 = New-Object system.Windows.Forms.TextBox
$FilePathTextBox.Multiline       = $false
$FilePathTextBox.Width           = 377
$FilePathTextBox.Height          = 25
$FilePathTextBox.Enabled         = $false
$FilePathTextBox.Location        = New-Object System.Drawing.Point(116,191)
$FilePathTextBox.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$SelectFileButton1               = New-Object system.Windows.Forms.Button
$SelectFileButton1.Text          = "Select File"
$SelectFileButton1.Width         = 114
$SelectFileButton1.Height        = 30
$SelectFileButton1.Location      = New-Object System.Drawing.Point(527,184)
$SelectFileButton1.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$BulkRenameButton                = New-Object system.Windows.Forms.Button
$BulkRenameButton.Text           = "Bulk Rename"
$BulkRenameButton.Width          = 156
$BulkRenameButton.Height         = 30
$BulkRenameButton.Location       = New-Object System.Drawing.Point(256,241)
$BulkRenameButton.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$RenameComputersForm.controls.AddRange(@($OldNameLabel,$NewNameLabel,$OldNameTextBox,$NewNameTextBox,$SingleRenameButton,$SingleRenameLabel,$BulkRenameLabel,$CSVFileLabel,$FilePathTextBox,$SelectFileButton1,$BulkRenameButton))

$SingleRenameButton.Add_Click({
    $OldName = $OldNameTextBox.text
    $NewName = $NewNameTextBox.text
    try {
        Rename-Computer -ComputerName $OldName -NewName $NewName -DomainCredential $Credential
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
            Rename-Computer -ComputerName $OldName -NewName $NewName -DomainCredential $Credential
        }
        catch {
            $Message ="Faild to rename " + $Computer.OldName + " computer `n" + $_.Exception
            continue
        }
        $Message = "Successfully renamed " + $Computer.OldName + " computer to " + $Computer.NewName
        Write-Log -Message $Message
    }
})

[void]$RenameComputersForm.ShowDialog()