
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


  function New-PopUpMessage($timeout,$message,$button1text,$button2text) {
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
    
    $path = "C:\_EFEX"
    $filepath = 'C:\_EFEX\efexlogo.jpg'
    $uri = "https://i.ibb.co/tLV5SGL/efexlogo.jpg"
    
    if (!(test-path -path $path)){
        New-Item -Path "C:\" -Name "_EFEX" -ItemType "Directory"
        Invoke-WebRequest -Uri $uri -OutFile $filepath
    }
    else {
    if (!(test-path -path $filepath)) {
        Invoke-WebRequest -Uri $uri -OutFile $filepath
    }
    }

    $file = (get-item 'C:\_EFEX\efexlogo.jpg')
    $ans = 0
    $img = [System.Drawing.Image]::Fromfile($file);
    
    # This tip from http://stackoverflow.com/questions/3358372/windows-forms-look-different-in-powershell-and-powershell-ise-why/3359274#3359274
    [System.Windows.Forms.Application]::EnableVisualStyles();
    $form = new-object Windows.Forms.Form
    $form.Text = "EFEX - Reboot Pending"
    $form.Width = 380;
    $form.Height = 250;
    $Form.StartPosition = "CenterScreen"
    $form.MinimizeBox = $false
    $form.MaximizeBox = $false
    $form.ControlBox = $false
    
    $pictureBox = new-object Windows.Forms.PictureBox
    $pictureBox.Width =  $img.Size.Width;
    $pictureBox.Height =  $img.Size.Height;
    $pictureBox.location = New-Object System.Drawing.Size(120,10)
    $pictureBox.Image = $img;
    $form.controls.add($pictureBox)
    
    $Button1 = New-Object System.Windows.Forms.Button
    $Button1.Location = New-Object System.Drawing.Size(100,160)
    $Button1.Size = New-Object System.Drawing.Size(80,23)
    $Button1.Text = $Button1Text
    $Form.Controls.Add($Button1)
    $Button1.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $Button1
    $form.Controls.Add($Button1)
    $Button1.BackColor = 'White'

    $Button2 = New-Object System.Windows.Forms.Button
    $Button2.Location = New-Object System.Drawing.Size(180,160)
    $Button2.Size = New-Object System.Drawing.Size(75,23)
    $Button2.Text = $Button2text
    $Button2.DialogResult = [System.Windows.Forms.DialogResult]::retry
    $Button2.BackColor = 'White'
    $form.AcceptButton = $Button2
    if($button2text -ne ""){
    $form.Controls.Add($Button2)
    }
    
    $Form.Controls.Add($DelayButton)
    
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = $timeout * 1000
    $timer.add_tick({$form.Close(); return })
    $timer.Start()
    
    $Label = New-Object System.Windows.Forms.Label
    $Label.Location = New-Object System.Drawing.Size(40,75)
    $label.TextAlign = "MiddleCenter"
    #$label.AutoSize = $false
    $Label.Size = New-Object System.Drawing.Size(280,80)
    $Label.Text = $message
    $Form.Controls.Add($Label)
    
    $form.Topmost = $true
    $form.add_formclosed({$timer.Enabled = $false})
    $form.ShowDialog()
    
    $form.Dispose()

}

# Variables for popup - Misc
   $countdown = 3
 
 # Variables for popup - Timers
    $timeout = 3600
    $timeout2 = 1800

# variable for popup - Messages
    

    
    #$message2 = "Hi $($env:username), 
    ##your computer is pending a restart. Please restart your computer now or delay for 1 hour. You have 2 postpones left!"
    
    #$message3 = "Hi $($env:username), 
    #your computer is pending a restart. Please restart your computer now or delay for 1 hour. You have 1 postpones left!"
    
    $message2 = "Hi $($env:username), 
    your computer is pending a restart. You have reached the maximum number of postpones. Please restart now or you can postpone for 30 more minutes.This popup will close and restart the machine in 30 minutes."
    
 
# Variable for popup - Buttons
    $button1 = "Restart Now"
    $button2 = "Postpone"
    $button1_1 = "Restart Now"
    $button2_2 = "Delay"

    
    if ($results -eq $true){
        do {
            $message = "Hi $($env:username), your computer is pending a restart. Please restart your computer now or delay for 1 hour. You have $countdown postpones left!"
            $result = New-PopupMessage -message $message -timeout $timeout -button1text $button1 -button2text $button2
                if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                    Write-Host "$($ENV:Username) clicked restart now"
                    $result = 0
                    Restart-Computer -Force
                }
                elseif ($result -eq [System.Windows.Forms.DialogResult]::retry){
                    Write-Host "$($ENV:Username) clicked Postpone 1 hour"
                    $result = 1
                    Start-Sleep -Seconds $timeout
                    $countdown--
                
                }
                else{
                    Write-Host "$($ENV:Username) did not select option within the timeout of $timeout seconds. Restarting..."
                    Write-Host "Popup timed out... waiting to reprompt"
                    Start-Sleep -seconds $timeout
                    $countdown--
                }
    
        } until ($countdown -eq 0 -or $result -eq 0)
    }
    
    Else {
        Write-Host "Computer $($ENV:ComputerName) does not require a reboot. Exiting Script"
        Exit
    }
     

    if($countdown -eq 0){
        New-PopupMessage -message $message2 -timeout $timeout2 -button1text $button1_1 -button2text $button2_2
        if($result -eq [System.Windows.Forms.DialogResult]::OK) {
            Write-Host "$($ENV:Username) clicked restart now"
            Restart-Computer -Force
        }
        Elseif($result -eq [System.Windows.Forms.DialogResult]::retry){
             write-host "$($ENV:Username) delayed for 30 mins"    
             Start-Sleep -Seconds $timeout2
             Restart-Computer -Force
        }
        else{
            Write-host "Popup timed out. Restarting in 30 minutes"
            Start-Sleep -seconds 60
            Restart-Computer -Force
        }
    }

