#set-executionpolicy remotesigned
import-module activedirectory
Add-Type -assembly System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Main form
$Form = New-Object System.Windows.Forms.Form
    $Form.Text = 'Indexing Info'
    $Form.Width = 280
    $Form.Height = 160
    $Form.StartPosition = 'CenterScreen'
    $Form.BackColor = 'LightGreen'
    $Form.ForeColor = 'Aquamarine'
    $Form.FormBorderStyle = 'FixedToolWindow'

# Label "PC:"
$LabelPC = New-Object System.Windows.Forms.Label
    $LabelPC.Text = "PC:"
    $LabelPC.Location = New-Object System.Drawing.Point(10, 6)
    $LabelPC.Autosize = $true
    $LabelPC.Font = 'Comic Sans MS, 10'
    $LabelPC.ForeColor = '#000080'
$Form.Controls.Add($LabelPC)

# Input field: PC
$TextBoxPC = New-Object System.Windows.Forms.Textbox
    $TextBoxPC.Location = New-Object System.Drawing.Point(45, 5)
    $TextBoxPC.Size = New-Object System.Drawing.Size(100, 10)
    $TextBoxPC.BackColor = '#F5F5DC'
    $TextBoxPC.ForeColor = '#000080'
    $TextBoxPC.Font = 'Comic Sans MS, 10'
$Form.Controls.Add($TextBoxPC)		                

# Button (Find)
$Button_Find2 = New-Object System.Windows.Forms.Button
    $Button_Find2.Location = New-Object System.Drawing.Point(150, 5)
    $Button_Find2.Size = New-Object System.Drawing.Size(90, 20)
    $Button_Find2.Text = 'Find'
    $Button_Find2.Font = 'Comic Sans MS, 12'
    $Button_Find2.BackColor = 'Green'
    $Button_Find2.ForeColor = '#000080'
$Form.Controls.Add($Button_Find2)

# Button (Rebuild)
$Reb = New-Object System.Windows.Forms.Button
    $Reb.Location = New-Object System.Drawing.Point(75, 90)
    $Reb.Size = New-Object System.Drawing.Size(90, 20)
    $Reb.Text = 'Rebuild'
    $Reb.Font = 'Comic Sans MS, 12'
    $Reb.ForeColor = '#000080'
    $Reb.BackColor = '#AF4035'
$Form.Controls.Add($Reb)

$ToolTip = New-Object System.Windows.Forms.ToolTip
    $ToolTip.BackColor = [System.Drawing.Color]::LightGoldenrodYellow
    $ToolTip.IsBalloon = $true

$MessageTextBox             = New-Object System.Windows.Forms.TextBox
$MessageTextBox2             = New-Object System.Windows.Forms.TextBox

$MessageTextBoxLabel        = New-Object System.Windows.Forms.Label
$MessageTextBoxLabel2        = New-Object System.Windows.Forms.Label

$User             = New-Object System.Windows.Forms.TextBox
$User2             = New-Object System.Windows.Forms.TextBox
$UserLabel        = New-Object System.Windows.Forms.Label
$UserLabel2      = New-Object System.Windows.Forms.Label

$null = ""

# colors of the main form
function RandomBacklight 
{` 
$random = New-Object System.Random 
switch ($random.Next(9)) 
    { 
      0 {$Form.BackColor = "LightBlue"} 
      1 {$Form.BackColor = "LightGreen"} 
      2 {$Form.BackColor = "LightPink"} 
      3 {$Form.BackColor = "Yellow"} 
      4 {$Form.BackColor = "Orange"} 
      5 {$Form.BackColor = "Brown"} 
      6 {$Form.BackColor = "Magenta"} 
      7 {$Form.BackColor = "White"} 
      8 {$Form.BackColor = "Gray"} 
    } 
}

# Actions
# Check the size of the indexing service file
$Button_Find2.Add_Click({

RandomBacklight

    $error.clear()
    $username = $null
    $username = $TextBoxPC.Text

$p1 =((Get-Item "\\$username\c$\ProgramData\Microsoft\Search\Data\Applications\Windows\Windows.edb").length/1GB)
$p1 = [math]::Round($p1, 2) 

$Ind = 'Indexing file size: ' + $p1 + ' Gb' 


$all = @($Ind +[System.Environment]::NewLine) 

$MessageTextBox.Text  = $null
$MessageTextBox.Text  = $all 

 })

# Rebuild (taskkill SearchIndexer, Delete file Windows.edb, Stop WSearch, Start WSearch)
$Reb.Add_Click({
    $error.clear()
    $username = $null
    $username = $TextBoxPC.Text

$Proc = Get-Process -Name SearchIndexer -ComputerName "$username"
taskkill /s \\$username /im searchindexer.exe /f
Start-Sleep -Seconds 2

$File = "\\$username\c$\ProgramData\Microsoft\Search\Data\Applications\Windows\Windows.edb"
Remove-item $File -force
Start-Sleep -Seconds 2

$ServiceOff = Get-WmiObject Win32_Service -Filter "Name = 'WSearch'" -ComputerName "$username"
$ServiceOff.StopService()
Start-Sleep -Seconds 5
$ServiceOn = Get-WmiObject Win32_Service -Filter "Name = 'WSearch'" -ComputerName "$username"
$ServiceOn.StartService()
 })

# Output field with information
$MessageTextBox =  New-Object System.Windows.Forms.TextBox
    $MessageTextBox.Text = $null
    $MessageTextBox.Location       = New-Object System.Drawing.Point(10,50)
    $MessageTextBox.MinimumSize = "240,30"
    $MessageTextBox.Multiline = $true
    $MessageTextBox.Font = 'Comic Sans MS, 10'
    $MessageTextBox.BackColor = '#F5F5DC'
    $MessageTextBox.ForeColor = '#000080'
        $ToolTip.SetToolTip($MessageTextBox, "Indexing info")

# Indexing info label
$MessageTextBoxLabel.Location  = New-Object System.Drawing.Point(10,30)
    $MessageTextBoxLabel.Text      = "Indexing info"
    $MessageTextBoxLabel.Font = 'Comic Sans MS, 10'
    $MessageTextBoxLabel.ForeColor = '#000080'
    $MessageTextBoxLabel.Autosize  = 1
        $ToolTip.SetToolTip($MessageTextBoxLabel, "Computer name")

$Form.Controls.Add($MessageTextBox)
$Form.Controls.Add($MessageTextBoxLabel)

# Form showing
$Form.ShowDialog()