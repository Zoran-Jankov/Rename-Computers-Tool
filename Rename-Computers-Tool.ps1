<#
.SYNOPSIS
PowerShell Tool for renaming AD computers with graphical user interface.

.DESCRIPTION
PowerShell Tool for renaming network computers with graphical user interface that can be used to rename single computer or to
perform bulk rename drawing data from a `.csv` file. The `.csv` file must have ComputerName and NewName columns for tool to work.

.NOTES
Version:        1.1
Author:         Zoran Jankov
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

$SearchBase = "OU=Korisnici,OU=Centrala,DC=uni,DC=net"

#-----------------------------------------------------------[Functions]------------------------------------------------------------

<#
.SYNOPSIS
Writes a log entry to console, log file and report file.

.DESCRIPTION
Creates a log entry with timestamp and message passed thru a parameter Message, and saves the log entry to log file. Depending on
the NoTimestamp parameter, log entry can be written with or without a timestamp. Format of the timestamp is
"yyyy.MM.dd. HH:mm:ss:fff", and this function adds " - " after timestamp and before the main message.

.PARAMETER Message
A string message to be written as a log entry

.PARAMETER NoTimestamp
A switch parameter if present timestamp is disabled in log entry

.EXAMPLE
Write-Log -Message "A log entry"

.EXAMPLE
Write-Log "A log entry"

.EXAMPLE
Write-Log -Message "===========" -NoTimestamp

.EXAMPLE
"A log entry" | Write-Log

.NOTES
Version:        2.2
Author:         Zoran Jankov
#>
function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
                   Position = 0,
                   ValueFromPipeline = $true,
                   ValueFromPipelineByPropertyName = $true,
                   HelpMessage = "A string message to be written as a log entry")]
        [string]
        $Message,

        [Parameter(Mandatory = $false,
                   Position = 1,
                   ValueFromPipeline = $true,
                   ValueFromPipelineByPropertyName = $true,
                   HelpMessage = "A switch parameter if present timestamp is disabled in log entry")]
        [switch]
        $NoTimestamp = $false
    )

    begin {
        $Desktop = [Environment]::GetFolderPath("Desktop")
        $LogFile = "$Desktop\Log.log"
        if (-not (Test-Path -Path $LogFile)) {
            New-Item -Path $LogFile -ItemType File
        }
    }

    process {
        if (-not($NoTimestamp)) {
            $Timestamp = Get-Date -Format "yyyy.MM.dd. HH:mm:ss:fff"
            $LogEntry = "$Timestamp - $Message"
        }
        else {
            $LogEntry = $Message
        }
        Add-content -Path $LogFile -Value $LogEntry
    }
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$MainForm                        = New-Object system.Windows.Forms.Form
$MainForm.ClientSize             = New-Object System.Drawing.Point(670,300)
$MainForm.Text                   = "Rename Computers Tool"
$MainForm.TopMost                = $true
$MainForm.FormBorderStyle        = 'Fixed3D'
$MainForm.MaximizeBox            = $false
$MainForm.ShowIcon               = $false
$MainForm.StartPosition          = [System.Windows.Forms.FormStartPosition]::CenterScreen
$MainForm.ForeColor              = "#FFFFFF"
$MainForm.BackColor              = "#1167B1"
$MainForm.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',11)

$SingleRenameLabel               = New-Object system.Windows.Forms.Label
$SingleRenameLabel.Text          = "Single Rename"
$SingleRenameLabel.AutoSize      = $false
$SingleRenameLabel.Width         = 670
$SingleRenameLabel.Height        = 40
$SingleRenameLabel.Location      = New-Object System.Drawing.Point(0,0)
$SingleRenameLabel.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif',15)
$SingleRenameLabel.TextAlign     = "MiddleCenter"
$SingleRenameLabel.BackColor     = "#3D3D3D"

$BulkRenameLabel                 = New-Object system.Windows.Forms.Label
$BulkRenameLabel.Text            = "Bulk Rename"
$BulkRenameLabel.AutoSize        = $false
$BulkRenameLabel.Width           = 25
$BulkRenameLabel.Height          = 10
$BulkRenameLabel.Location        = New-Object System.Drawing.Point(285,165)

$ComputerNameLabel               = New-Object system.Windows.Forms.Label
$ComputerNameLabel.Text          = "   Computer"
$ComputerNameLabel.AutoSize      = $false
$ComputerNameLabel.Width         = 100
$ComputerNameLabel.Height        = 25
$ComputerNameLabel.Location      = New-Object System.Drawing.Point(0,60)
$ComputerNameLabel.TextAlign     = "MiddleLeft"
$ComputerNameLabel.BackColor     = "#3D3D3D"

$ComputerNameComboBox            = New-Object system.Windows.Forms.ComboBox
$ComputerNameComboBox.Width      = 380
$ComputerNameComboBox.Location   = New-Object System.Drawing.Point(100,60)
$ComputerNameComboBox.AutoCompleteMode = 'SuggestAppend'
$ComputerNameComboBox.AutoCompleteSource = 'ListItems'

$NewNameLabel                    = New-Object system.Windows.Forms.Label
$NewNameLabel.Text               = "   New Name"
$NewNameLabel.AutoSize           = $false
$NewNameLabel.Width              = 100
$NewNameLabel.Height             = 25
$NewNameLabel.Location           = New-Object System.Drawing.Point(0,95)
$NewNameLabel.TextAlign     = "MiddleLeft"
$NewNameLabel.BackColor     = "#3D3D3D"

$NewNameTextBox                  = New-Object system.Windows.Forms.TextBox
$NewNameTextBox.Multiline        = $false
$NewNameTextBox.Width            = 380
$NewNameTextBox.Height           = 25
$NewNameTextBox.Location         = New-Object System.Drawing.Point(110,116)
$NewNameTextBox.MaxLength        = 15

$CSVFileLabel                    = New-Object system.Windows.Forms.Label
$CSVFileLabel.Text               = "CSV File"
$CSVFileLabel.AutoSize           = $true
$CSVFileLabel.Width              = 25
$CSVFileLabel.Height             = 10
$CSVFileLabel.Location           = New-Object System.Drawing.Point(25,210)

$FilePathTextBox                 = New-Object system.Windows.Forms.TextBox
$FilePathTextBox.Multiline       = $false
$FilePathTextBox.Width           = 380
$FilePathTextBox.Height          = 25
$FilePathTextBox.Enabled         = $false
$FilePathTextBox.Location        = New-Object System.Drawing.Point(115,205)

$SelectFileButton               = New-Object system.Windows.Forms.Button
$SelectFileButton.Text          = "Select File"
$SelectFileButton.Width         = 120
$SelectFileButton.Height        = 30
$SelectFileButton.Location      = New-Object System.Drawing.Point(525,200)
$SelectFileButton.BackColor = "#FFFFFF"
$SelectFileButton.ForeColor = "#000000"

$SingleRenameButton              = New-Object system.Windows.Forms.Button
$SingleRenameButton.Text         = "Rename"
$SingleRenameButton.Width        = 115
$SingleRenameButton.Height       = 30
$SingleRenameButton.Location     = New-Object System.Drawing.Point(525,80)

$BulkRenameButton                = New-Object system.Windows.Forms.Button
$BulkRenameButton.Text           = "Bulk Rename"
$BulkRenameButton.Width          = 160
$BulkRenameButton.Height         = 30
$BulkRenameButton.Location       = New-Object System.Drawing.Point(250,250)

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
    $SelectFileButton,
    $BulkRenameButton
))

$Computers = Get-ADComputer -Filter {Enabled -eq $true} -SearchBase $SearchBase | Sort-Object 'Name'
foreach ($Computer in $Computers) {
    $ComputerNameComboBox.Items.Add($Computer.Name);
}

$SingleRenameButton.Add_Click({
    $OldName = $ComputerNameComboBox.text
    $NewName = $NewNameTextBox.text
    $Message = "Are you sure you want to rename '$OldName' computer to '$NewName'?"
    $Choice =  [System.Windows.Forms.MessageBox]::Show(
        $Message, "Rename Computer", 
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($Choice -eq 'Yes') {
        Rename-Computer -ComputerName $OldName -NewName $NewName
        Get-ADComputer -Identity $NewName
        if ($?) {
            $Message = "Successfully renamed '$OldName' computer to '$NewName'"
            $Title = "Info"
            $Icon = [System.Windows.Forms.MessageBoxIcon]::Asterisk
        }
        else {
            $Message ="Faild to rename '$OldName' computer"
            $Title = "Error"
            $Icon = [System.Windows.Forms.MessageBoxIcon]::Error
        }
        Write-Log -Message $Message
        [System.Windows.Forms.MessageBox]::Show(
            $Message,
            $Title,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            $Icon
        )
    }
})

$SelectFileButton.Add_Click({
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    $OpenFileDialog.Filter = "CSV Files (*.csv)| *.csv*"
    $OpenFileDialog.ShowDialog() | Out-Null
    $FilePathTextBox.Text = $OpenFileDialog.Filename
})

$BulkRenameButton.Add_Click({
    $Message = "Are you sure you want to rename '$OldName' computer to '$NewName'?"
    $Choice =  [System.Windows.Forms.MessageBox]::Show(
        $Message, "Rename Computer", 
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($Choice -eq 'Yes') {
        $CSVFile = Get-Content -Path $OpenFileDialog.Filename
        $Info = ""
        foreach ($Computer in $CSVFile) {
            $OldName = $Computer.ComputerName
            $NewName = $Computer.NewName
            Rename-Computer -ComputerName $OldName -NewName $NewName
            Get-ADComputer -Identity $NewName
            if ($?) {
                $Message = "Successfully renamed '$OldName' computer to '$NewName'"
            }
            else {
                $Message ="Faild to rename '$OldName' computer"
            }
            Write-Log -Message $Message
            $Info += "$Message`r`n"
            [System.Windows.Forms.MessageBox]::Show(
                $Info, "Info",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Asterisk
            )
        }
    }
})
[void]$MainForm.ShowDialog()