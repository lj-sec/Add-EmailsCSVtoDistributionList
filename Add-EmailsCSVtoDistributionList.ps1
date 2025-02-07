#Requires -Version 5.1

<#
.SYNOPSIS
A PowerShell script featuring a GUI to add emails into Microsoft Exchange Online distribution list groups.
Requires PS 5.1.

.NOTES
Author: Logan Jackson
Date: 2025

.LINK
Website: https://lj-sec.github.io/
#>

Add-Type -AssemblyName PresentationFramework, PresentationCore, System.Windows.Forms, System.Drawing

# Welcome message
[System.Windows.MessageBox]::Show("A tool to automatically add emails from a csv onto an Exchange mailing list`nClick okay, then select the .csv.","Alert",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information) > $null

# Prompt user for a CSV file
Do {
    # Prove me wrong
    $nocsv=$false

    # File dialog
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.InitialDirectory = $env:USERPROFILE
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"

    # Collect response
    $dialogResult = $openFileDialog.ShowDialog()

    # Check if OK
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK)
    {
        # Grab the file and check if it is a CSV
        $selectedFile = $openFileDialog.FileName
        if ([System.IO.Path]::GetExtension($selectedFile) -ne ".csv")
        {
            [System.Windows.MessageBox]::Show("Incorrect file type, ensure this is a .csv","Alert",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Warning) > $null
            $nocsv = $true
        }
    }
    else
    {
        Exit 1
    }
} while($nocsv) # if nocsv was true, loop again

# Declare combobox, label, and form so they are accessible outside of the function
$script:comboBox = $null
$script:form = $null
$script:label = $null

# Prepare a new form, will be used twice to gather user input
function New-Form()
{
    $script:form = New-Object System.Windows.Forms.Form
    $script:form.Size = New-Object System.Drawing.Size(750,400)
    $script:form.Text = "Script"
    $script:form.StartPosition = "CenterScreen"
    $script:form.FormBorderStyle = "FixedDialog"
    $script:form.Topmost = $true

    $script:label = New-Object System.Windows.Forms.Label
    $script:label.Size = New-Object System.Drawing.Size(700,35)
    $script:label.Location = New-Object System.Drawing.Point(25,25)
    $script:label.Font = New-Object System.Drawing.Font("Arial", 11)
    $script:form.Controls.Add($script:label)

    $script:comboBox = New-Object System.Windows.Forms.ComboBox
    $script:comboBox.Size = New-Object System.Drawing.Size(675,30)
    $script:comboBox.Location = New-Object System.Drawing.Point(25,75)
    $script:comboBox.Font = New-Object System.Drawing.Font("Arial", 11)
    $script:comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $script:form.Controls.Add($script:comboBox)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Size = New-Object System.Drawing.Size(150,35)
    $okButton.Font = New-Object System.Drawing.Font("Arial", 11)
    $okButton.Location = New-Object System.Drawing.Point(25,250)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $script:form.AcceptButton = $okButton
    $script:form.Controls.Add($okButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Size = New-Object System.Drawing.Size(150,35)
    $cancelButton.Font = New-Object System.Drawing.Font("Arial", 11)
    $cancelButton.Location = New-Object System.Drawing.Point(200,250)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $script:form.CancelButton = $cancelButton
    $script:form.Controls.Add($cancelButton)
}

# Attempt to import the ExchangeOnlineManagement module, if it fails, attempt to install it
try
{
    Import-Module -Name ExchangeOnlineManagement -ErrorAction Stop
}
catch
{
    try
    {
        Get-PackageProvider -Name NuGet 1>$null -ErrorAction Stop
    }
    catch
    {
        [System.Windows.MessageBox]::Show("NuGet Package Manager is a requirement for script functionality, exiting script.","Error",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Error)
        Exit 1
    }
    $confirm = [System.Windows.MessageBox]::Show("ExchangeOnlineManagement Module not found, would you like to install now?`nThis is a requirement for script functionality.","Warning",[System.Windows.MessageBoxButton]::YesNo,[System.Windows.MessageBoxImage]::Warning)
    if($confirm -eq [System.Windows.MessageBoxResult]::Yes)
    {
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -WarningAction 'SilentlyContinue' -Force -AllowClobber
        Import-Module -Name ExchangeOnlineManagement
    }
    else
    {
        [System.Windows.MessageBox]::Show("ExchangeOnlineManagement Module is a requirement for script functionality, exiting script.","Error",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Error)
        Exit 1
    }
}

# Connect to exchange online
try
{
    Connect-ExchangeOnline -ShowBanner:$false -SkipLoadingCmdletHelp -ErrorAction Stop
}
catch
{
    [System.Windows.MessageBox]::Show("Could not connect to Exchange Online, exiting","Error",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Error) > $null
    Exit 1
}

[System.Windows.MessageBox]::Show("Connected. Press OK to load distribution lists. This may take a minute depending on the size of your organization.","Alert",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information) > $null

Write-Host "Loading (please allow some time)..."

# Grab the user's email and distribution lists that they own
$email = (Get-MailBox).Name
$managedDLs = Get-DistributionGroup -WarningAction 'SilentlyContinue' | Where-Object {$_.ManagedBy -icontains "$email"}

if ($managedDLs.Count -eq 0)
{
    [System.Windows.MessageBox]::Show("No managed distribution lists found. Exiting.","Error",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Error)
    Exit 1
}

# Import the csv and and headers from it
$emailsCsv = Import-Csv -LiteralPath $openFileDialog.FileName
$headers = $emailsCsv[0].PSObject.Properties.Name

if ($headers.Count -eq 0)
{
    [System.Windows.MessageBox]::Show("No headers in the CSV file found. Exiting.","Error",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Error)
    Exit 1
}

# Prepare a form
New-Form

# Update explanation label
$script:label.Text = "Select the column of the CSV containing emails:"

# Populate the combobox with headers
foreach($header in $headers)
{
    $script:comboBox.Items.Add($header) > $null
}

# Prompt user to choose header containing emails
$result = $script:form.ShowDialog()

# Check for OK
if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $selectedHeader = $script:comboBox.SelectedItem
}
else
{
    Exit 1
}

# Ensure that there was a header selected
if($null -eq $selectedHeader)
{
    [System.Windows.MessageBox]::Show("No headers in the CSV file selected. Exiting.","Error",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Error)
    Exit 1
}

# Prepare another form
New-Form

# Update explanation label
$script:label.Text = "Select the distribution group:"

# Clear out and repopulate the combobox
$script:comboBox.Items.Clear()
foreach($DL in $managedDLs)
{
    $script:comboBox.Items.Add($DL.Name) > $null
}

# Prompt user to choose DL
$result = $script:form.ShowDialog()

# Check for OK
if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $selectedDL = $script:comboBox.SelectedItem
}
else
{
    Exit 1
}

# Ensure there was a list selected
if($null -eq $selectedDl)
{
    [System.Windows.MessageBox]::Show("No distribution group selected. Exiting.","Error",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Error)
    Exit 1
}

# Prepare a warning message for if emails fail to be added
$warningMessage = "The following emails could not be added (check for typos or if they are already on the list):`n"

# Assume warning message will not show
$warning = $false
$numAdded = 0
$index = 0

# Loop to add emails to list
foreach($email in $emailsCsv)
{
    $index++
    $percent = [Math]::Round(($index/$emailsCsv.Length)*100, 2)
    Write-Progress -Activity "Adding $member..." -Status "$percent% Complete:" -PercentComplete $percent

    # Grab the email
    $member = $email."$($selectedHeader)"
    # Attempt to add, if it fails add the user to the warning message
    try
    {
        Add-DistributionGroupMember -Identity $selectedDL -Member $member -ErrorAction Stop
        $numAdded++
    }
    catch
    {
        $warningMessage += "$($email."$($selectedHeader)")`n"
        # Update the warning messaage boolean
        $warning = $true
    }
}

Disconnect-ExchangeOnline -Confirm:$false

# Check for warning message, if not give success message
if($warning)
{
    [System.Windows.MessageBox]::Show("$warningMessage`n$numAdded emails were successfully added.","Alert",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Warning) > $null
}
else
{
    [System.Windows.MessageBox]::Show("Success! All $numAdded emails added to list","Alert",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information) > $null
}