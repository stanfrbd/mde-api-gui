<#
.SYNOPSIS
	Simple PowerShell GUI for Microsoft Defender for Endpoint API machine actions.

.DESCRIPTION
	Sample response tool that benefits from APIs and are using PowerShell as the tool of choice to perform actions in bulk. It doesn't require installation and can easily be adapted by anyone with some scripting experience. The tool currently accepts advanced hunting queries, computer names, and CSVs as device input methods. Once devices are selected, three types of actions can be performed:

	- Tagging devices
	- Performing Quick/Full AV scan, and
	- Performing Isolation/Release from Isolation

	An Azure AD AppID and Secret is required to connect to API and the tool needs the following App Permissions:

	- AdvancedQuery.Read.All
	- Machine.Isolate
	- Machine.ReadWrite.All
	- Machine.Scan

#>


#===========================================================[Classes]===========================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;
public class ProcessDPI {
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool SetProcessDPIAware();      
}
'@
$null = [ProcessDPI]::SetProcessDPIAware()


#===========================================================[Variables]===========================================================


$script:selectedmachines = @{}
$credspath = 'c:\temp\mdeuicreds.txt'
$helpQueryBox = "
AH Query: Specify advanced hunting query that will return DeviceId and DeviceName, e.g.: 'DeviceInfo | distinct DeviceId, DeviceName'`r`n
Computer Name(s): Specify one o more FQDN computer names, e.g.: 'computer1.contoso.com, computer2.contoso.com'`r`n
CSV: Specify path to CSV file with computer names, e.g.: 'C:\temp\Computers.csv'"

$UnclickableColour = "#8d8989"
$ClickableColour = "#ff7b00"
$TextBoxFont = 'Microsoft Sans Serif,10'

#===========================================================[WinForm]===========================================================


[System.Windows.Forms.Application]::EnableVisualStyles()


$MainForm = New-Object system.Windows.Forms.Form
$MainForm.SuspendLayout()
$MainForm.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
$MainForm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$MainForm.ClientSize = '950,800'
$MainForm.text = "MDE API GUI"
$MainForm.BackColor = "#ffffff"
$MainForm.TopMost = $false

$Title = New-Object system.Windows.Forms.Label
$Title.text = "1 - Connect with MDE API Credentials"
$Title.AutoSize = $true
$Title.width = 25
$Title.height = 10
$Title.location = New-Object System.Drawing.Point(20, 20)
$Title.Font = 'Microsoft Sans Serif,12,style=Bold'

$AppIdBoxLabel = New-Object system.Windows.Forms.Label
$AppIdBoxLabel.text = "App Id:"
$AppIdBoxLabel.AutoSize = $true
$AppIdBoxLabel.width = 25
$AppIdBoxLabel.height = 10
$AppIdBoxLabel.location = New-Object System.Drawing.Point(20, 50)
$AppIdBoxLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$AppIdBox = New-Object system.Windows.Forms.TextBox
$AppIdBox.multiline = $false
$AppIdBox.width = 314
$AppIdBox.height = 20
$AppIdBox.location = New-Object System.Drawing.Point(100, 50)
$AppIdBox.Font = $TextBoxFont
$AppIdBox.Visible = $true

$AppSecretBoxLabel = New-Object system.Windows.Forms.Label
$AppSecretBoxLabel.text = "App Secret:"
$AppSecretBoxLabel.AutoSize = $true
$AppSecretBoxLabel.width = 25
$AppSecretBoxLabel.height = 10
$AppSecretBoxLabel.location = New-Object System.Drawing.Point(20, 75)
$AppSecretBoxLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$AppSecretBox = New-Object system.Windows.Forms.TextBox
$AppSecretBox.multiline = $false
$AppSecretBox.width = 314
$AppSecretBox.height = 20
$AppSecretBox.location = New-Object System.Drawing.Point(100, 75)
$AppSecretBox.Font = $TextBoxFont
$AppSecretBox.Visible = $true
$AppSecretBox.PasswordChar = '*'

$TenantIdBoxLabel = New-Object system.Windows.Forms.Label
$TenantIdBoxLabel.text = "Tenant Id:"
$TenantIdBoxLabel.AutoSize = $true
$TenantIdBoxLabel.width = 25
$TenantIdBoxLabel.height = 10
$TenantIdBoxLabel.location = New-Object System.Drawing.Point(20, 100)
$TenantIdBoxLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$TenantIdBox = New-Object system.Windows.Forms.TextBox
$TenantIdBox.multiline = $false
$TenantIdBox.width = 314
$TenantIdBox.height = 20
$TenantIdBox.location = New-Object System.Drawing.Point(100, 100)
$TenantIdBox.Font = $TextBoxFont
$TenantIdBox.Visible = $true

$ConnectionStatusLabel = New-Object system.Windows.Forms.Label
$ConnectionStatusLabel.text = "Status:"
$ConnectionStatusLabel.AutoSize = $true
$ConnectionStatusLabel.width = 25
$ConnectionStatusLabel.height = 10
$ConnectionStatusLabel.location = New-Object System.Drawing.Point(20, 135)
$ConnectionStatusLabel.Font = 'Microsoft Sans Serif,10,style=Bold'

$ConnectionStatus = New-Object system.Windows.Forms.Label
$ConnectionStatus.text = "Not Connected"
$ConnectionStatus.AutoSize = $true
$ConnectionStatus.width = 25
$ConnectionStatus.height = 10
$ConnectionStatus.location = New-Object System.Drawing.Point(100, 135)
$ConnectionStatus.Font = 'Microsoft Sans Serif,10'

$SaveCredCheckbox = new-object System.Windows.Forms.checkbox
$SaveCredCheckbox.Location = New-Object System.Drawing.Point(200, 135)
$SaveCredCheckbox.AutoSize = $true
$SaveCredCheckbox.width = 60
$SaveCredCheckbox.height = 10
$SaveCredCheckbox.Text = "Save Credentials"
$SaveCredCheckbox.Font = 'Microsoft Sans Serif,10'
$SaveCredCheckbox.Checked = $false

$ConnectBtn = New-Object system.Windows.Forms.Button
$ConnectBtn.BackColor = "#ff7b00"
$ConnectBtn.text = "Connect"
$ConnectBtn.width = 90
$ConnectBtn.height = 30
$ConnectBtn.location = New-Object System.Drawing.Point(325, 130)
$ConnectBtn.Font = 'Microsoft Sans Serif,10'
$ConnectBtn.ForeColor = "#ffffff"
$ConnectBtn.Visible = $True

$TitleActions = New-Object system.Windows.Forms.Label
$TitleActions.text = "3 - Perform Action on selected devices"
$TitleActions.AutoSize = $true
$TitleActions.width = 25
$TitleActions.height = 10
$TitleActions.location = New-Object System.Drawing.Point(500, 20)
$TitleActions.Font = 'Microsoft Sans Serif,12,style=Bold'

$TagDeviceGroupBox = New-Object System.Windows.Forms.GroupBox
$TagDeviceGroupBox.Location = New-Object System.Drawing.Point(500, 40)
$TagDeviceGroupBox.width = 400
$TagDeviceGroupBox.height = 50
$TagDeviceGroupBox.Text = "Device tag"
$TagDeviceGroupBox.Font = 'Microsoft Sans Serif,10,style=Bold'

$DeviceTag = New-Object system.Windows.Forms.TextBox
$Devicetag.multiline = $false
$DeviceTag.width = 200
$DeviceTag.height = 25
$DeviceTag.location = New-Object System.Drawing.Point(20, 20)
$Devicetag.Font = 'Microsoft Sans Serif,10'
$DeviceTag.Visible = $true
$Devicetag.Enabled = $false

$TagDeviceBtn = New-Object system.Windows.Forms.Button
$TagDeviceBtn.BackColor = $UnclickableColour
$TagDeviceBtn.text = "Apply Tag"
$TagDeviceBtn.width = 110
$TagDeviceBtn.height = 30
$TagDeviceBtn.location = New-Object System.Drawing.Point(280, 15)
$TagDeviceBtn.Font = 'Microsoft Sans Serif,10'
$TagDeviceBtn.ForeColor = "#ffffff"
$TagDeviceBtn.Visible = $true
$TagDeviceBtn.Enabled = $false

# enlarge groupbox to fit new button
$TagDeviceGroupBox.Height = 90

$RemoveTagBtn = New-Object System.Windows.Forms.Button
$RemoveTagBtn.BackColor = $UnclickableColour
$RemoveTagBtn.Text = "Remove Tag"
$RemoveTagBtn.Width = 110
$RemoveTagBtn.Height = 30
$RemoveTagBtn.Location = New-Object System.Drawing.Point(280, 50)
$RemoveTagBtn.Font = 'Microsoft Sans Serif,10'
$RemoveTagBtn.ForeColor = "#ffffff"
$RemoveTagBtn.Visible = $true
$RemoveTagBtn.Enabled = $false

$TagDeviceGroupBox.Controls.AddRange(@($DeviceTag, $TagDeviceBtn, $RemoveTagBtn))

$ScanGroupBox = New-Object System.Windows.Forms.GroupBox
$ScanGroupBox.Location = New-Object System.Drawing.Point(500, 140)    # moved down to avoid overlap with TagDeviceGroupBox
$ScanGroupBox.width = 400
$ScanGroupBox.height = 50
$ScanGroupBox.Text = "Scan mode"
$ScanGroupBox.Font = 'Microsoft Sans Serif,10,style=Bold'

$ScanRadioButton1 = New-Object System.Windows.Forms.RadioButton
$ScanRadioButton1.Width = 80
$ScanRadioButton1.Height = 20
$ScanRadioButton1.location = New-Object System.Drawing.Point(20, 20)
$ScanRadioButton1.Checked = $false
$ScanRadioButton1.Enabled = $false
$ScanRadioButton1.Text = "Full Scan"
$ScanRadioButton1.Font = 'Microsoft Sans Serif,8'
 
$ScanRadioButton2 = New-Object System.Windows.Forms.RadioButton
$ScanRadioButton2.Width = 80
$ScanRadioButton2.Height = 20
$ScanRadioButton2.location = New-Object System.Drawing.Point(120, 20)
$ScanRadioButton2.Checked = $true
$ScanRadioButton2.Enabled = $false
$ScanRadioButton2.Text = "Quick Scan"
$ScanRadioButton2.Font = 'Microsoft Sans Serif,8'

$ScanDeviceBtn = New-Object system.Windows.Forms.Button
$ScanDeviceBtn.BackColor = $UnclickableColour
$ScanDeviceBtn.text = "AV Scan"
$ScanDeviceBtn.width = 110
$ScanDeviceBtn.height = 30
$ScanDeviceBtn.location = New-Object System.Drawing.Point(280, 15)
$ScanDeviceBtn.Font = 'Microsoft Sans Serif,10'
$ScanDeviceBtn.ForeColor = "#ffffff"
$ScanDeviceBtn.Visible = $true

$ScanGroupBox.Controls.AddRange(@($ScanRadioButton1, $ScanRadioButton2, $ScanDeviceBtn))

$IsolateGroupBox = New-Object System.Windows.Forms.GroupBox
$IsolateGroupBox.Location = New-Object System.Drawing.Point(500, 200)  # moved down to sit below ScanGroupBox
$IsolateGroupBox.Width = 400
$IsolateGroupBox.height = 90
$IsolateGroupBox.text = "Isolation"
$IsolateGroupBox.Font = 'Microsoft Sans Serif,10,style=Bold'

$IsolateRadioButton1 = New-Object System.Windows.Forms.RadioButton
$IsolateRadioButton1.width = 60
$IsolateRadioButton1.height = 20
$IsolateRadioButton1.location = New-Object System.Drawing.Point(20, 20)
$IsolateRadioButton1.Checked = $false
$IsolateRadioButton1.Enabled = $false
$IsolateRadioButton1.Text = "Full"
$IsolateRadioButton1.Font = 'Microsoft Sans Serif,8'
 
$IsolateRadioButton2 = New-Object System.Windows.Forms.RadioButton
$IsolateRadioButton2.width = 120
$IsolateRadioButton2.height = 20
$IsolateRadioButton2.location = New-Object System.Drawing.Point(120, 20)
$IsolateRadioButton2.Checked = $true
$IsolateRadioButton2.Enabled = $false
$IsolateRadioButton2.Text = "Selective"
$IsolateRadioButton2.Font = 'Microsoft Sans Serif,8'

$IsolateDeviceBtn = New-Object system.Windows.Forms.Button
$IsolateDeviceBtn.BackColor = $UnclickableColour
$IsolateDeviceBtn.text = "Isolate Device"
$IsolateDeviceBtn.width = 110
$IsolateDeviceBtn.height = 30
$IsolateDeviceBtn.location = New-Object System.Drawing.Point(280, 15)
$IsolateDeviceBtn.Font = 'Microsoft Sans Serif,10'
$IsolateDeviceBtn.ForeColor = "#ffffff"
$IsolateDeviceBtn.Visible = $true

$ReleaseFromIsolationBtn = New-Object system.Windows.Forms.Button
$ReleaseFromIsolationBtn.BackColor = $UnclickableColour
$ReleaseFromIsolationBtn.text = "Release Device"
$ReleaseFromIsolationBtn.width = 110
$ReleaseFromIsolationBtn.height = 30
$ReleaseFromIsolationBtn.location = New-Object System.Drawing.Point(280, 50)
$ReleaseFromIsolationBtn.Font = 'Microsoft Sans Serif,10'
$ReleaseFromIsolationBtn.ForeColor = "#ffffff"
$ReleaseFromIsolationBtn.Visible = $true

$IsolateGroupBox.Controls.AddRange(@($IsolateRadioButton1, $IsolateRadioButton2, $IsolateDeviceBtn, $ReleaseFromIsolationBtn))

$InputCsvFileBox = New-Object System.Windows.Forms.GroupBox
$InputCsvFileBox.width = 880
$InputCsvFileBox.height = 240
$InputCsvFileBox.location = New-Object System.Drawing.Point(20, 290)
$InputCsvFileBox.text = "2 - Select devices to perform action on (CSV)"
$InputCsvFileBox.Font = 'Microsoft Sans Serif,12,style=Bold'

$GetDevicesFromQueryBtn = New-Object System.Windows.Forms.Button
$GetDevicesFromQueryBtn.BackColor = $UnclickableColour
$GetDevicesFromQueryBtn.text = "Get Devices"
$GetDevicesFromQueryBtn.width = 180
$GetDevicesFromQueryBtn.height = 30
$GetDevicesFromQueryBtn.location = New-Object System.Drawing.Point(690, 190)
$GetDevicesFromQueryBtn.Font = 'Microsoft Sans Serif,10'
$GetDevicesFromQueryBtn.ForeColor = "#ffffff"
$GetDevicesFromQueryBtn.Visible = $true

$SelectedDevicesBtn = New-Object system.Windows.Forms.Button
$SelectedDevicesBtn.BackColor = $UnclickableColour
$SelectedDevicesBtn.text = "Selected Devices (" + $script:selectedmachines.Keys.count + ")"
$SelectedDevicesBtn.width = 150
$SelectedDevicesBtn.height = 30
$SelectedDevicesBtn.location = New-Object System.Drawing.Point(530, 190)
$SelectedDevicesBtn.Font = 'Microsoft Sans Serif,10'
$SelectedDevicesBtn.ForeColor = "#ffffff"
$SelectedDevicesBtn.Visible = $false

$ClearSelectedDevicesBtn = New-Object system.Windows.Forms.Button
$ClearSelectedDevicesBtn.BackColor = $UnclickableColour
$ClearSelectedDevicesBtn.text = "Clear Selection"
$ClearSelectedDevicesBtn.width = 150
$ClearSelectedDevicesBtn.height = 30
$ClearSelectedDevicesBtn.location = New-Object System.Drawing.Point(370, 190)
$ClearSelectedDevicesBtn.Font = 'Microsoft Sans Serif,10'
$ClearSelectedDevicesBtn.ForeColor = "#ffffff"
$ClearSelectedDevicesBtn.Visible = $false

# CSV file picker controls (shown when InputRadioButton3 is selected)
$CsvPathBox = New-Object system.Windows.Forms.TextBox
$CsvPathBox.multiline = $false
$CsvPathBox.width = 700
$CsvPathBox.height = 25
$CsvPathBox.location = New-Object System.Drawing.Point(20, 60)
$CsvPathBox.Font = $TextBoxFont
$CsvPathBox.ReadOnly = $true
$CsvPathBox.Enabled = $false

$BrowseCsvBtn = New-Object system.Windows.Forms.Button
$BrowseCsvBtn.BackColor = $UnclickableColour
$BrowseCsvBtn.text = "Browse..."
$BrowseCsvBtn.width = 90
$BrowseCsvBtn.height = 25
$BrowseCsvBtn.location = New-Object System.Drawing.Point(730, 60)
$BrowseCsvBtn.Font = 'Microsoft Sans Serif,9'
$BrowseCsvBtn.ForeColor = "#ffffff"
$BrowseCsvBtn.Visible = $false
$BrowseCsvBtn.Enabled = $false

# add a label with short description of what to do with the CSV (should have "Name" header and only hostnames) under the browse button
$CsvDescLabel = New-Object system.Windows.Forms.Label
$CsvDescLabel.text = "Select a CSV file with a 'Name' header (one single column) containing hostnames (one per line)."
$CsvDescLabel.width = 700
$CsvDescLabel.height = 40
$CsvDescLabel.location = New-Object System.Drawing.Point(20, 90)
$CsvDescLabel.Font = 'Microsoft Sans Serif,9'
# use a visible colour (black) on the white form background
$CsvDescLabel.ForeColor = "#000000"
$CsvDescLabel.Visible = $true

# OpenFileDialog for CSV selection
$OpenCsvDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenCsvDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
$OpenCsvDialog.Multiselect = $false

$CsvPathBox.Visible = $true
$BrowseCsvBtn.Visible = $true


$InputCsvFileBox.Controls.AddRange(@(
        $CsvPathBox,
        $BrowseCsvBtn,
        $CsvDescLabel,
        $GetDevicesFromQueryBtn,
        $SelectedDevicesBtn,
        $ClearSelectedDevicesBtn
    ))

$LogBoxLabel = New-Object system.Windows.Forms.Label
$LogBoxLabel.text = "4 - Logs:"
$LogBoxLabel.width = 394
$LogBoxLabel.height = 20
$LogBoxLabel.location = New-Object System.Drawing.Point(20, 600)
$LogBoxLabel.Font = 'Microsoft Sans Serif,12,style=Bold'
$LogBoxLabel.Visible = $true

$LogBox = New-Object system.Windows.Forms.TextBox
$LogBox.multiline = $true
$LogBox.width = 880
$LogBox.height = 100
$LogBox.location = New-Object System.Drawing.Point(20, 630)
$LogBox.ScrollBars = 'Vertical'
$LogBox.Font = $TextBoxFont
$LogBox.Visible = $true

$ExportLogBtn = New-Object system.Windows.Forms.Button
$ExportLogBtn.BackColor = '#FFF0F8FF'
$ExportLogBtn.text = "Export Logs"
$ExportLogBtn.width = 90
$ExportLogBtn.height = 30
$ExportLogBtn.location = New-Object System.Drawing.Point(20, 750)
$ExportLogBtn.Font = 'Microsoft Sans Serif,10'
$ExportLogBtn.ForeColor = "#ff000000"
$ExportLogBtn.Visible = $true

$GetActionsHistoryBtn = New-Object system.Windows.Forms.Button
$GetActionsHistoryBtn.BackColor = $UnclickableColour
$GetActionsHistoryBtn.text = "Get Actions History"
$GetActionsHistoryBtn.width = 150
$GetActionsHistoryBtn.height = 30
$GetActionsHistoryBtn.location = New-Object System.Drawing.Point(130, 750)
$GetActionsHistoryBtn.Font = 'Microsoft Sans Serif,10'
$GetActionsHistoryBtn.ForeColor = "#ffffff"
$GetActionsHistoryBtn.Visible = $true

$ExportActionsHistoryBtn = New-Object system.Windows.Forms.Button
$ExportActionsHistoryBtn.BackColor = $UnclickableColour
$ExportActionsHistoryBtn.text = "Export Actions History"
$ExportActionsHistoryBtn.width = 150
$ExportActionsHistoryBtn.height = 30
$ExportActionsHistoryBtn.location = New-Object System.Drawing.Point(300, 750)
$ExportActionsHistoryBtn.Font = 'Microsoft Sans Serif,10'
$ExportActionsHistoryBtn.ForeColor = "#ffffff"
$ExportActionsHistoryBtn.Visible = $true

$cancelBtn = New-Object system.Windows.Forms.Button
$cancelBtn.BackColor = '#FFF0F8FF'
$cancelBtn.text = "Cancel"
$cancelBtn.width = 90
$cancelBtn.height = 30
$cancelBtn.location = New-Object System.Drawing.Point(810, 750)
$cancelBtn.Font = 'Microsoft Sans Serif,10'
$cancelBtn.ForeColor = "#ff000000"
$cancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$MainForm.CancelButton = $cancelBtn
$MainForm.Controls.Add($cancelBtn)

#$MainForm.AutoScaleMode = 'dpi'

$MainForm.controls.AddRange(@($Title,
        $ConnectionStatusLabel, 
        $ConnectionStatus,
        $cancelBtn, 
        $AppIdBox, 
        $AppSecretBox,
        $TenantIdBox, 
        $AppIdBoxLabel, 
        $AppSecretBoxLabel, 
        $TenantIdBoxLabel, 
        $ConnectBtn, 
        $TitleActions, 
        $LogBoxLabel, 
        $LogBox, 
        $IsolateGroupBox,
        $SaveCredCheckbox,
        $ScanGroupBox,
        $InputCsvFileBox,
        $TagDeviceGroupBox,
        $ExportLogBtn,
        $GetActionsHistoryBtn,
        $ExportActionsHistoryBtn))


#===========================================================[Functions]===========================================================


#Authentication

function GetToken {
    $ConnectionStatus.ForeColor = "#000000"
    $ConnectionStatus.Text = 'Connecting...'
    $tenantId = $TenantIdBox.Text
    $appId = $AppIdBox.Text
    $appSecret = $AppSecretBox.Text
    $resourceAppIdUri = 'https://api.securitycenter.windows.com'
    $oAuthUri = "https://login.windows.net/$TenantId/oauth2/token"
    $authBody = [Ordered] @{
        resource      = "$resourceAppIdUri"
        client_id     = "$appId"
        client_secret = "$appSecret"
        grant_type    = 'client_credentials'
    }
    
    $authResponse = Invoke-RestMethod -Method Post -Uri $oAuthUri -Body $authBody -ErrorAction Stop
    $token = $authResponse.access_token
    $script:headers = @{
        'Content-Type' = 'application/json'
        Accept         = 'application/json'
        Authorization  = "Bearer $token"
    }
    
    if ($authresponse) {
        $ConnectionStatus.text = "Connected"
        $ConnectionStatus.ForeColor = "#7ed321"
        $LogBox.AppendText((get-date).ToString() + " Successfully connected to Tenant ID: " + $tenantId + [Environment]::NewLine)
        ChangeButtonColours -Buttons $GetDevicesFromQueryBtn, $SelectedDevicesBtn, $ClearSelectedDevicesBtn, $ExportActionsHistoryBtn, $GetActionsHistoryBtn, $BrowseCsvBtn
        EnableRadioButtons
        SaveCreds
        $Devicetag.Enabled = $true
        $CsvPathBox.Enabled = $true
        $BrowseCsvBtn.Enabled = $true
        return $headers
    }
    else {
        $ConnectionStatus.text = "Connection Failed"
        [System.Windows.Forms.MessageBox]::Show("ErrorMessage: " + $Error[0] , "Error")
        $ConnectionStatus.ForeColor = "#D0021B"
        $cancelBtn.text = "Close"
    }

}

function SaveCreds {
    if ($SaveCredCheckbox.Checked) {
        $securespassword = $AppSecretBox.Text | ConvertTo-SecureString -AsPlainText -Force
        $securestring = $securespassword | ConvertFrom-SecureString
        $creds = @($TenantIdBox.Text, $AppIdBox.Text, $securestring)
        $creds | Out-File $credspath
    }
}

function ChangeButtonColours {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $True)]
        $Buttons
    )
    $ButtonsToChangeColour = $Buttons

    foreach ( $Button in $ButtonsToChangeColour) {
        $Button.BackColor = $ClickableColour
    }
}

function EnableRadioButtons {
    $ButtonsToEnable = $ScanRadioButton1, $ScanRadioButton2, $IsolateRadioButton1, $IsolateRadioButton2

    foreach ( $Button in $ButtonsToEnable) {
        $Button.Enabled = $true
    }
}


function AddTagDevice {
    $script:selectedmachines.GetEnumerator() | foreach-object {
        Start-Sleep -Seconds 1
        $MachineId = $_.value
        $MachineTag = $DeviceTag.Text
        $body = @{
            "Value"  = $MachineTag;
            "Action" = "Add";
        }

        $url = "https://api.securitycenter.windows.com/api/machines/$MachineId/tags" 
        try { $webResponse = Invoke-WebRequest -Method Post -Uri $url -Headers $headers -Body ($body | ConvertTo-Json) -ContentType "application/json" -ErrorAction Stop }
        Catch {
            if ($_.ErrorDetails.Message) {
                [System.Windows.Forms.MessageBox]::Show("ErrorMessage: " + $_.ErrorDetails.Message , "Error")
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("Status: " + $webResponse.StatusCode)
            }
        }
        if ($null -ne $webResponse.statuscode) { 
            $LogBox.AppendText((get-date).ToString() + " Applying machine tag: " + $MachineTag + " Machine Name: " + $_.Key + " Status code: " + $webResponse.statuscode + [Environment]::NewLine) 
        }
        
    }
}

function RemoveTagDevice {
    $script:selectedmachines.GetEnumerator() | ForEach-Object {
        Start-Sleep -Seconds 1
        $MachineId = $_.Value
        $MachineTag = $DeviceTag.Text

        if ([string]::IsNullOrWhiteSpace($MachineTag)) {
            [System.Windows.Forms.MessageBox]::Show("Please specify a tag to remove.", "Information")
            return
        }

        $body = @{
            "Value"  = $MachineTag;
            "Action" = "Remove";
        }

        $url = "https://api.securitycenter.windows.com/api/machines/$MachineId/tags"
        try {
            $webResponse = Invoke-WebRequest -Method Post -Uri $url -Headers $headers -Body ($body | ConvertTo-Json) -ContentType "application/json" -ErrorAction Stop
        }
        catch {
            if ($_.ErrorDetails.Message) {
                [System.Windows.Forms.MessageBox]::Show("ErrorMessage: " + $_.ErrorDetails.Message , "Error")
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("Status: " + $webResponse.StatusCode)
            }
        }

        if ($null -ne $webResponse.statuscode) {
            $LogBox.AppendText((Get-Date).ToString() + " Removing machine tag: " + $MachineTag + " Machine Name: " + $_.Key + " Status code: " + $webResponse.statuscode + [Environment]::NewLine)
        }
    }
}



function ScanDevice {
    $script:selectedmachines.GetEnumerator() | foreach-object {
        Start-Sleep -Seconds 1
        $machineid = $_.Value
        if ($ScanRadioButton1.Checked) { $ScanMode = 'Full' } else { $ScanMode = 'Quick' }
        $body = @{
            "Comment"  = "AV Scan";
            "ScanType" = $ScanMode;
        }
        $url = "https://api.securitycenter.windows.com/api/machines/$machineid/runAntiVirusScan" 
        try { $webResponse = Invoke-WebRequest -Method Post -Uri $url -Headers $headers -Body ($body | ConvertTo-Json) -ContentType "application/json" -ErrorAction Stop }
        Catch {
            if ($_.ErrorDetails.Message) {
                [System.Windows.Forms.MessageBox]::Show("ErrorMessage: " + $_.ErrorDetails.Message , "Error")
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("Status: " + $webResponse.StatusCode)
            }
        }
        if ($null -ne $webResponse.statuscode) { $LogBox.AppendText((get-date).ToString() + " " + $ScanMode + " AV Scan on Machine Name: " + $_.Key + " Status code: " + $webResponse.statuscode + [Environment]::NewLine) }
    }
}

function IsolateDevice {
    $script:selectedmachines.GetEnumerator() | foreach-object {
        Start-Sleep -Seconds 1
        $machineid = $_.Value
        $IsolationType = 'Selective'
        if ($IsolateRadioButton1.Checked) { $IsolationType = 'Full' }
        $body = @{
            "Comment"       = "Isolating device";
            "IsolationType" = $IsolationType;
        }
        $url = "https://api.securitycenter.windows.com/api/machines/$machineid/isolate" 
        try { $webResponse = Invoke-WebRequest -Method Post -Uri $url -Headers $headers -Body ($body | ConvertTo-Json) -ContentType "application/json" -ErrorAction Stop }
        Catch {
            if ($_.ErrorDetails.Message) {
                #[System.Windows.Forms.MessageBox]::Show("ErrorMessage: " + $_.ErrorDetails.Message + $_.ErrorDetails, "Error")
                $LogBox.AppendText((get-date).ToString() + " ErrorMessage: " + $_.ErrorDetails.Message + $_.Exception.Response.StatusCode + [Environment]::NewLine)
                
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("Status: " + $webResponse.StatusCode)
            }
        }
        if ($null -ne $webResponse.statuscode) { $LogBox.AppendText((get-date).ToString() + " " + $IsolationType + " Isolation on: " + " Machine Name: " + $_.Key + " Status code: " + $webResponse.statuscode + [Environment]::NewLine) }
    }
}

function ReleaseFromIsolation {
    $script:selectedmachines.GetEnumerator() | foreach-object {
        Start-Sleep -Seconds 1
        $machineid = $_.Value
        $body = @{
            "Comment" = "Releasing device from isolation";
        }
        $url = "https://api.securitycenter.windows.com/api/machines/$machineid/unisolate" 
        try { $webResponse = Invoke-WebRequest -Method Post -Uri $url -Headers $headers -Body ($body | ConvertTo-Json) -ContentType "application/json" -ErrorAction Stop }
        Catch {
            if ($_.ErrorDetails.Message) {
                #[System.Windows.Forms.MessageBox]::Show("ErrorMessage: " + $_.ErrorDetails.Message + $_.ErrorDetails, "Error")
                $LogBox.AppendText("ErrorMessage: " + $_.ErrorDetails.Message + $_.Exception.Response.StatusCode + [Environment]::NewLine)
                
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("Status: " + $webResponse.StatusCode)
            }
        }
        if ($null -ne $webResponse.statuscode) { $LogBox.AppendText($IsolationType + " Releasing isolation on: " + " Machine Name: " + $_.Key + " Status code: " + $webResponse.statuscode + [Environment]::NewLine) }
    }
}


# This function is not present in GUI to avoid any unwanted changes to the environments
function OffboardDevice {
    $script:selectedmachines.GetEnumerator() | foreach-object {
        Start-Sleep -Seconds 1
        $machineid = $_.Value
        $body = @{
            "Comment" = "Offboarding machine using API";
        }
        $url = "https://api.securitycenter.windows.com/api/machines/$machineid/offboard" 
        try { $webResponse = Invoke-WebRequest -Method Post -Uri $url -Headers $headers -Body ($body | ConvertTo-Json) -ContentType "application/json" -ErrorAction Stop }
        Catch {
            if ($_.ErrorDetails.Message) {
                #[System.Windows.Forms.MessageBox]::Show("ErrorMessage: " + $_.ErrorDetails.Message + $_.ErrorDetails, "Error")
                $LogBox.AppendText("ErrorMessage: " + $_.ErrorDetails.Message + $_.Exception.Response.StatusCode + [Environment]::NewLine)
                
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("Status: " + $webResponse.StatusCode)
            }
        }
        if ($null -ne $webResponse.statuscode) { $LogBox.AppendText("Offboarding machine: " + [Environment]::NewLine + " Machine Name: " + $_.Key + " Status code: " + $webResponse.statuscode + [Environment]::NewLine) }
    }
}

function ViewSelectedDevices {
    $filtermachines = $script:selectedmachines | Out-GridView -Title "Select devices to perform action on:" -PassThru 
    $script:selectedmachines.clear()
    foreach ($machine in $filtermachines) {
        $script:selectedmachines.Add($machine.Name, $machine.Value)
    }
    $SelectedDevicesBtn.text = "Selected Devices (" + $script:selectedmachines.Keys.count + ")"
    if ($null -eq $script:selectedmachines.Keys.Count) {
        $SelectedDevicesBtn.Visible = $false
        $SelectedDevicesBtn.text = "Selected Devices (" + $script:selectedmachines.Keys.count + ")"
        $ClearSelectedDevicesBtn.Visible = $false
    }
    $LogBox.AppendText((get-date).ToString() + " Devices selected count: " + ($script:selectedmachines.Keys.count -join [Environment]::NewLine) + [Environment]::NewLine + ($script:selectedmachines.Keys -join [Environment]::NewLine) + [Environment]::NewLine)
}

function ClearSelectedDevices {
    $script:selectedmachines = @{}
    $ClearSelectedDevicesBtn.Visible = $false
    $SelectedDevicesBtn.Visible = $false
    $LogBox.AppendText((get-date).ToString() + " Devices selected count: " + $script:selectedmachines.Keys.count + [Environment]::NewLine)
}


function GetDevicesFromCsv {
    if ((Test-Path $CsvPathBox.Text) -and ($CsvPathBox.Text).EndsWith(".csv")) {
        $machines = Import-Csv -Path $CsvPathBox.Text
        $script:selectedmachines = @{}
        $LogBox.AppendText("Quering " + $machines.count + " machines from CSV file... Please wait" + [Environment]::NewLine)
        foreach ($machine in $machines) {
            # Add sleep if needed to avoid throttling
            # Start-Sleep -Seconds 1
            $MachineName = $machine.Name
            $url = "https://api.securitycenter.windows.com/api/machines/$MachineName"  
            $webResponse = Invoke-RestMethod -Method Get -Uri $url -Headers $headers -ErrorAction Stop
            $MachineId = $webResponse.id
            if (-not $script:selectedmachines.contains($machine.Name)) {
                $script:selectedmachines.Add($machine.Name, $MachineId)
            }
        }
        $filtermachines = $script:selectedmachines | Out-GridView -Title "Select devices to perform action on:" -PassThru 
        $script:selectedmachines.clear()
        foreach ($machine in $filtermachines) {
            $script:selectedmachines.Add($machine.Name, $machine.Value)
        }
        if ($script:selectedmachines.Keys.Count -gt 0) {
            ChangeButtonColours -Buttons $TagDeviceBtn, $RemoveTagBtn, $ScanDeviceBtn, $IsolateDeviceBtn, $ReleaseFromIsolationBtn
            $SelectedDevicesBtn.Visible = $true
            $TagDeviceBtn.Enabled = $true
            $RemoveTagBtn.Enabled = $true
            $SelectedDevicesBtn.text = "Selected Devices (" + $script:selectedmachines.Keys.count + ")"
            $ClearSelectedDevicesBtn.Visible = $true
        }
        $LogBox.AppendText((get-date).ToString() + " Devices selected count: " + ($script:selectedmachines.Keys.count -join [Environment]::NewLine) + [Environment]::NewLine + ($script:selectedmachines.Keys -join [Environment]::NewLine) + [Environment]::NewLine)
    } 
    else {
        [System.Windows.Forms.MessageBox]::Show($CsvPathBox.Text + " is not a valid CSV path." , "Error")
    }
}


function GetActionsHistory {
    $LogBox.AppendText("Getting machine actions list.." + [Environment]::NewLine)
    $url = "https://api-us.securitycenter.windows.com/api/machineactions" 
    try { $webResponse = Invoke-WebRequest -Method Get -Uri $url -Headers $headers -ErrorAction Stop }
    Catch {
        if ($_.ErrorDetails.Message) {
            [System.Windows.Forms.MessageBox]::Show("ErrorMessage: " + $_.ErrorDetails.Message , "Error")
        }
        else {
            $LogBox.AppendText((get-date).ToString() + " Status: " + $webResponse.StatusCode)
        }
    }
    $results = ($webResponse.Content | Convertfrom-json).value
    $LogBox.AppendText((get-date).ToString() + " Status: " + $webResponse.StatusCode + " Machine actions count: " + $results.count + [Environment]::NewLine)
    $LogBox.AppendText((get-date).ToString() + " Last 10 machine actions: " + ($results | Select-Object type, computerDnsName, status -First 10 | Out-string) + [Environment]::NewLine)
    $results | Out-GridView -Title "Actions History" -PassThru 
}

function ExportActionsHistory {
    $LogBox.AppendText("Getting machine actions list.." + [Environment]::NewLine)
    $url = "https://api-us.securitycenter.windows.com/api/machineactions" 
    try { $webResponse = Invoke-WebRequest -Method Get -Uri $url -Headers $headers -ErrorAction Stop }
    Catch {
        if ($_.ErrorDetails.Message) {
            [System.Windows.Forms.MessageBox]::Show("ErrorMessage: " + $_.ErrorDetails.Message , "Error")
        }
        else {
            $LogBox.AppendText((get-date).ToString() + " Status: " + $webResponse.StatusCode)
        }
    }
    $results = ($webResponse.Content | Convertfrom-json).value
    $LogBox.AppendText((get-date).ToString() + " Status: " + $webResponse.StatusCode + " Machine actions count: " + $results.count + [Environment]::NewLine)
    $results | Export-Csv -Path .\Response_Actions.csv -NoTypeInformation
    $LogBox.AppendText((get-date).ToString() + " Export file created: " + (Get-Item .\Response_Actions.csv).FullName + [Environment]::NewLine)
}


function ExportLog {
    $LogBox.Text | Out-file .\mde_ui_log.txt
    $LogBox.AppendText((get-date).ToString() + " Log file created: " + (Get-Item .\mde_ui_log.txt).FullName + [Environment]::NewLine)
}

#===========================================================[Script]===========================================================


if (test-path $credspath) {
    $creds = Get-Content $credspath
    $pass = $creds[2] | ConvertTo-SecureString
    $unsecurePassword = [PSCredential]::new(0, $pass).GetNetworkCredential().Password
    $TenantIdBox.Text = $creds[0]
    $AppIdBox.Text = $creds[1]
    $AppSecretBox.Text = $unsecurePassword
}


$ConnectBtn.Add_Click({ GetToken })

$TagDeviceBtn.Add_Click({ AddTagDevice })

$RemoveTagBtn.Add_Click({ RemoveTagDevice })

$ScanDeviceBtn.Add_Click({ ScanDevice })

$IsolateDeviceBtn.Add_Click({ IsolateDevice })

$ReleaseFromIsolationBtn.Add_Click({ ReleaseFromIsolation })

$GetDevicesFromQueryBtn.Add_Click({ GetDevicesFromCsv })

$BrowseCsvBtn.Add_Click({
        if ($OpenCsvDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $CsvPathBox.Text = $OpenCsvDialog.FileName
            # show selected path in log
            $LogBox.AppendText((get-date).ToString() + " Selected CSV: " + $CsvPathBox.Text + [Environment]::NewLine)
            # enable the GetDevicesFromQueryBtn so user can proceed
            $GetDevicesFromQueryBtn.BackColor = $ClickableColour
        }
    })

$SelectedDevicesBtn.Add_Click({ ViewSelectedDevices })

$ClearSelectedDevicesBtn.Add_Click({ ClearSelectedDevices })

$ExportLogBtn.Add_Click({ ExportLog })

$GetActionsHistoryBtn.Add_Click({ getActionsHistory })

$ExportActionsHistoryBtn.Add_Click({ ExportActionsHistory })

$MainForm.ResumeLayout()
[void]$MainForm.ShowDialog()