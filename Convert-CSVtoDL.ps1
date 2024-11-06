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
    $script:label.Size = New-Object System.Drawing.Size(700,30)
    $script:label.Location = New-Object System.Drawing.Point(25,25)
    $script:label.Font = New-Object System.Drawing.Font("Arial", 12)
    $script:form.Controls.Add($script:label)

    $script:comboBox = New-Object System.Windows.Forms.ComboBox
    $script:comboBox.Size = New-Object System.Drawing.Size(675,30)
    $script:comboBox.Location = New-Object System.Drawing.Point(25,75)
    $script:comboBox.Font = New-Object System.Drawing.Font("Arial", 12)
    $script:comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $script:form.Controls.Add($script:comboBox)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Size = New-Object System.Drawing.Size(100,30)
    $okButton.Font = New-Object System.Drawing.Font("Arial", 12)
    $okButton.Location = New-Object System.Drawing.Point(25,250)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $script:form.AcceptButton = $okButton
    $script:form.Controls.Add($okButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Size = New-Object System.Drawing.Size(100,30)
    $cancelButton.Font = New-Object System.Drawing.Font("Arial", 12)
    $cancelButton.Location = New-Object System.Drawing.Point(150,250)
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
    $confirm = [System.Windows.MessageBox]::Show("ExchangeOnlineManagement Module not found, would you like to install now?`nThis is a requirement for script functionality.","Warning",[System.Windows.MessageBoxButton]::YesNo,[System.Windows.MessageBoxImage]::Warning)
    if($confirm -eq "Yes")
    {
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -WarningAction 'SilentlyContinue' -Force
        Import-Module -Name ExchangeOnlineManagement
    }
    else
    {
        [System.Windows.MessageBox]::Show("ExchangeOnlineManagement Module is requirement for script functionality, exiting script.","Warning",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information)   
    }
}

# Connect to exchange online
try
{
    Connect-ExchangeOnline -ShowBanner:$false -SkipLoadingCmdletHelp -ErrorAction Stop
}
catch
{
    [System.Windows.MessageBox]::Show("Could not connect to Exchange Online, exiting","Alert",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Warning) > $null
    Exit 1
}

[System.Windows.MessageBox]::Show("Connected, loading distribution lists. This may take a minute depending on the size of your organization.`nPress OK to continue","Alert",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information) > $null

# Grab the user's email and distribution lists that they own
$email = (Get-MailBox).Name
$managedDLs = Get-DistributionGroup -WarningAction 'SilentlyContinue' | Where-Object {$_.ManagedBy -icontains "$email"}

# Import the csv and and headers from it
$emailsCsv = Import-Csv -LiteralPath $openFileDialog.FileName
$headers = $emailsCsv[0].PSObject.Properties.Name

# Prepare a form
New-Form

# Update explanation label
$script:label.Text = "Select the column of the CSV containing email addresses to add:"

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

# Prepare another form
New-Form

# Update explanation label
$script:label.Text = "Select the distribution group you would like to add these emails to:"

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

# Prepare a warning message for if emails fail to be added
$warningMessage = "The following emails could not be added (check for typos or if they are already on the list):`n"

# Assume warning message will not show
$warning = $false

# Loop to add emails to list
foreach($email in $emailsCsv)
{
    # Grab the email
    $member = $email."$($selectedHeader)"
    # Attempt to add, if it fails add the user to the warning message
    try
    {
        Add-DistributionGroupMember -Identity $selectedDL -Member $member -ErrorAction Stop
    }
    catch
    {
        $warningMessage += "$($email."$($selectedHeader)")`n"
        # Update the warning messaage bool
        $warning = $true
    }
}

Disconnect-ExchangeOnline -Confirm:$false

# Check for warning message, if not give success message
if($warning)
{
    [System.Windows.MessageBox]::Show("$warningMessage","Alert",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Warning) > $null
}
else
{
    [System.Windows.MessageBox]::Show("Success!","Alert",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information) > $null
}