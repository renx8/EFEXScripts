
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


### Varibales for popup box

$timeout = 10
$Button1Text = 'Update Now'
$Button2Text = 'Postpone'
$message = "Hi $($env:username),
your computer is running a vulnerable version of office. Please update office now or postpone for 1 hour. 
Note - updating office will restart office applications. Please save your work prior to updating."

$message2 = "You have postponed the maximum number of times. Please update now! If you postpone again, your office will update in 30 minutes."
$timeout2 = 5
$button1text2 = "Update Now"
$button2text2 = "Postpone"
         
$result = 1   
$counter = 0    
 ### Check user button pressed and update or postpone depending on selection.
         do {  
            $popup = New-PopUpMessage -message $message -timeout $timeout -button1text $button1text -button2text $Button2Text  
            if ($popup -eq [System.Windows.Forms.DialogResult]::OK) {
                Write-Host "$($ENV:Username) clicked Update now"
                $result = 0
                cmd /c "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe" /update user displaylevel=false forceappshutdown=true
            }
            elseif ($popup -eq [System.Windows.Forms.DialogResult]::retry){
                Write-Host "$($ENV:Username) clicked Postpone. Delaying for 1 hour..."
                $result =1
                Start-Sleep -Seconds $timeout
                $counter++
                
            }
            else{
                Write-Host "Popup timed out... waiting to reprompt"
                Start-Sleep -seconds $timeout
                $counter++
            }
            } Until ($result -eq 0 -or $counter -eq 3)
            
    if ($counter -eq 3){
        Write-Host "$($ENV:Username) has postponed 3 times. The next popup will force reboot"
        $popup = New-PopUpMessage -message $message2 -timeout $timeout2 -button1text $button1text2 -button2text $button2text2
            if($popup -eq [System.Windows.Forms.DialogResult]::OK){
                Write-Host "$($ENV:Username) clicked Update now"
                cmd /c "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe" /update user displaylevel=false forceappshutdown=true
            Elseif ($popup -eq [System.Windows.Forms.DialogResult]::retry){
                write-host "user postponed again... updating in 30 minutes"
                Start-Sleep -seconds $timeout2
                cmd /c "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe" /update user displaylevel=false forceappshutdown=true
            }
            else {
                write-host "Popup Timed out... updating in 30 minutes"
                Start-Sleep -seconds $timeout2
                cmd /c "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe" /update user displaylevel=false forceappshutdown=true
        }
    }
                
}     
