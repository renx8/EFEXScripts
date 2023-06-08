<#
EFEX Office Updater V1.0
This script will update Office 365 if the user accepts the efex branded pop up. If the user does not accept the script will exit.
#>

function get-efexpopup() {
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$file = (get-item 'C:\_EFEX\efexlogo.jpg')
$ans = 0
$img = [System.Drawing.Image]::Fromfile($file);

# This tip from http://stackoverflow.com/questions/3358372/windows-forms-look-different-in-powershell-and-powershell-ise-why/3359274#3359274
[System.Windows.Forms.Application]::EnableVisualStyles();
$form = new-object Windows.Forms.Form
$form.Text = "EFEX Update Alert"
$form.Width = 380;
$form.Height = 250;
$Form.StartPosition = "CenterScreen"

$Form.KeyPreview = $True
$Form.Add_KeyDown({
if ($_.KeyCode -eq "Enter"){
    $global:ans = 1;$Form.Close()
  }
})
$Form.Add_KeyDown({if ($_.KeyCode -eq "Escape")
    {$global:ans = 0;$Form.Close()}})

$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.Width =  $img.Size.Width;
$pictureBox.Height =  $img.Size.Height;
$pictureBox.location = New-Object System.Drawing.Size(120,10)
$pictureBox.Image = $img;
$form.controls.add($pictureBox)

$UpdateButton = New-Object System.Windows.Forms.Button
$UpdateButton.Location = New-Object System.Drawing.Size(100,160)
$UpdateButton.Size = New-Object System.Drawing.Size(75,23)
$UpdateButton.Text = "Update"
$Form.Controls.Add($UpdateButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(180,160)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$Form.Controls.Add($CancelButton)

$UpdateButton.Add_Click({
    $global:ans = 1
    $Form.Close()
 })

$CancelButton.Add_Click({
    $global:ans = 0
    $Form.Close()
})

$Label = New-Object System.Windows.Forms.Label
$Label.Location = New-Object System.Drawing.Size(40,75)
$label.TextAlign = "MiddleCenter"
#$label.AutoSize = $false
$Label.Size = New-Object System.Drawing.Size(280,80)
$Label.Text = "Hi, Efex would like to update your Office 365 applications. Please hit Update if you are ready to update or otherwise hit cancel to exit.
note - The update will force restart all of your Office 365 apps."
$Form.Controls.Add($Label)

$Form.Topmost = $True
$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()
Write-output $ans
    
}

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

# Call function efex-popup and check confitionals. If user hits update or presses enter, update proceeds. If user hits escape or cancel, script exits.
get-efexpopup

if ($ans -eq 1){
    cmd /c "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe" /update user displaylevel=false forceappshutdown=true
}
Else{
    Write-output "user cancelled update script. Exiting script"
}