# reboot Script V3.0 
# This script will reboot any machines that have a pending reboot.

#remain hidden
# .Net methods for hiding/showing the console in the background
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
function Hide-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [Console.Window]::ShowWindow($consolePtr, 0)
}
Hide-Console

# Hash table to store registry checks
<#PSScriptInfo
 
.VERSION 1.11
 
.GUID fe3d3698-52fc-40e8-a95c-bbc67a507ed1
 
.AUTHOR Adam Bertram
 
.COMPANYNAME Adam the Automator, LLC
 
.COPYRIGHT
 
.DESCRIPTION This script tests various registry values to see if the local computer is pending a reboot.
 
.TAGS
 
.LICENSEURI
 
.PROJECTURI
 
.ICONURI
 
.EXTERNALMODULEDEPENDENCIES
 
.REQUIREDSCRIPTS
 
.EXTERNALSCRIPTDEPENDENCIES
 
.RELEASENOTES
 
.SYNOPSIS
    This script tests various registry values to see if the local computer is pending a reboot
.NOTES
    Inspiration from: https://gallery.technet.microsoft.com/scriptcenter/Get-PendingReboot-Query-bdb79542
.EXAMPLE
    PS> Test-PendingReboot -ComputerName localhost
     
    This example checks various registry values to see if the local computer is pending a reboot.
#>


function Test-RegistryKey {
    [OutputType('bool')]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Key
    )

    $ErrorActionPreference = 'Stop'

    if (Get-Item -Path $Key -ErrorAction Ignore) {
        $true
    }
}

function Test-RegistryValue {
    [OutputType('bool')]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Key,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Value
    )

    $ErrorActionPreference = 'Stop'

    if (Get-ItemProperty -Path $Key -Name $Value -ErrorAction Ignore) {
        $true
    }
}

function Test-RegistryValueNotNull {
    [OutputType('bool')]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Key,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Value
    )

    $ErrorActionPreference = 'Stop'

    if (($regVal = Get-ItemProperty -Path $Key -Name $Value -ErrorAction Ignore) -and $regVal.($Value)) {
        $true
    }
}

# Added "test-path" to each test that did not leverage a custom function from above since
# an exception is thrown when Get-ItemProperty or Get-ChildItem are passed a nonexistant key path
$tests = @(
    { Test-RegistryKey -Key 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending' }
    { Test-RegistryKey -Key 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootInProgress' }
    { Test-RegistryKey -Key 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired' }
    { Test-RegistryKey -Key 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending' }
    { Test-RegistryKey -Key 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\PostRebootReporting' }
    { Test-RegistryValueNotNull -Key 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Value 'PendingFileRenameOperations' }
    { Test-RegistryValueNotNull -Key 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Value 'PendingFileRenameOperations2' }
    { 
        # Added test to check first if key exists, using "ErrorAction ignore" will incorrectly return $true
        'HKLM:\SOFTWARE\Microsoft\Updates' | Where-Object { test-path $_ -PathType Container } | ForEach-Object {            
            (Get-ItemProperty -Path $_ -Name 'UpdateExeVolatile' -ErrorAction Ignore | Select-Object -ExpandProperty UpdateExeVolatile) -ne 0 
        }
    }
    { Test-RegistryValue -Key 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce' -Value 'DVDRebootSignal' }
    { Test-RegistryKey -Key 'HKLM:\SOFTWARE\Microsoft\ServerManager\CurrentRebootAttemps' }
    { Test-RegistryValue -Key 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon' -Value 'JoinDomain' }
    { Test-RegistryValue -Key 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon' -Value 'AvoidSpnSet' }
    {
        # Added test to check first if keys exists, if not each group will return $Null
        # May need to evaluate what it means if one or both of these keys do not exist
        ( 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName' | Where-Object { test-path $_ } | %{ (Get-ItemProperty -Path $_ ).ComputerName } ) -ne 
        ( 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName' | Where-Object { Test-Path $_ } | %{ (Get-ItemProperty -Path $_ ).ComputerName } )
    }
    {
        # Added test to check first if key exists
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Services\Pending' | Where-Object { 
            (Test-Path $_) -and (Get-ChildItem -Path $_) } | ForEach-Object { $true }
    }
)

foreach ($test in $tests) {
    Write-Verbose "Running scriptblock: [$($test.ToString())]"
    if (& $test) {
        $results = $true
        break
    }
}

function New-PopupMessage ($timeout,$message){
    Add-Type -AssemblyName system.windows.forms
    Add-Type -AssemblyName system.drawing
    

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "EFEX IT - Restart Required!"
    $form.Size = New-Object System.Drawing.Size(500,160)
    $form.StartPosition = 'CenterScreen'
    $form.BackColor = 'LightGreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MinimizeBox = $false
    $form.MaximizeBox = $false
    $form.ControlBox = $false

    $label = New-Object System.Windows.Forms.label
    $label.Text = $message
    $label.Size = New-Object System.Drawing.Size(450,80)
    $Label.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $label.Location = New-Object System.Drawing.Point(10,10)
    $Label.Font = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
    $form.Controls.Add($label)
    
    #add button to form - Restart now
    $restartButton = New-Object System.Windows.Forms.Button
    $restartButton.Location = New-Object System.Drawing.Point(310,90)
    $restartButton.Size = New-Object System.Drawing.Size(80,23)
    $restartButton.Text = 'Restart Now'
    $restartButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $restartButton
    $form.Controls.Add($restartButton)
    $restartButton.BackColor = 'White'

    #add button to form - Delay 1 Hour
    $delayButton = New-Object System.Windows.Forms.Button
    $delayButton.Location = New-Object System.Drawing.Point(400,90)
    $delayButton.Size = New-Object System.Drawing.Size(75,23)
    $delayButton.Text = 'Delay 1H'
    $delayButton.DialogResult = [System.Windows.Forms.DialogResult]::retry
    $delayButton.BackColor = 'White'
    $form.AcceptButton = $delayButton
    $form.Controls.Add($delayButton)
    
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = $timeout * 1000
    $timer.add_tick({$form.Close(); return })
    $timer.Start()
    
    $form.Topmost = $true
    $form.add_formclosed({$timer.Enabled = $false})
    $form.ShowDialog()
    
    $form.Dispose()
}

$timeout = 3600
$message = "Hi $($env:username), 
your computer is pending a restart. Please restart your computer now or delay for 1 hour. Note - Your computer will restart in 1 hour if no option is selected!"


if ($results -eq $true){

    $result = New-PopupMessage -message $message -timeout $timeout
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {f
            Write-Host "$($ENV:Username) clicked restart now"
            Restart-Computer -Force
        }
        elseif ($result -eq [System.Windows.Forms.DialogResult]::retry){
            Write-Host "$($ENV:Username) clicked Delay 1 Hour"
            Start-Sleep -Seconds $timeout
            Restart-Computer -Force
        }
        else{
            Write-Host "$($ENV:Username) did not select option within the timeout of $timeout seconds. Restarting..."
        }

}
Else {
    Write-Host "Computer $($ENV:ComputerName) does not require a reboot. Exiting Script"
    Exit
}
