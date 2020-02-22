Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Select a Computer'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(75,120)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(150,120)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
Add-Type -AssemblyName System.Windows.Forms

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '400,400'
$Form.text                       = "AD user extractor tool"
$Form.TopMost                    = $false

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Select the user"
$Label1.AutoSize                 = $true
$Label1.width                    = 96
$Label1.height                   = 5
$Label1.location                 = New-Object System.Drawing.Point(21,38)
$Label1.Font                     = 'Microsoft Sans Serif,10'

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "Click me!"
$Button1.width                   = 90
$Button1.height                  = 45
$Button1.location                = New-Object System.Drawing.Point(150,92)
$Button1.Font                    = 'Microsoft Sans Serif,10'

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $true
$TextBox1.width                  = 319
$TextBox1.height                 = 100
$TextBox1.location               = New-Object System.Drawing.Point(50,156)
$TextBox1.Font                   = 'Microsoft Sans Serif,10'

$ComboBox1                       = New-Object system.Windows.Forms.ComboBox
$ComboBox1.text                  = "Pick a Role" 
$comboBox1.Items.Add("jagat");$comboBox1.Items.Add("Admin")
$ComboBox1.width                 = 227
$ComboBox1.height                = 11
$ComboBox1.location              = New-Object System.Drawing.Point(123,35)
$ComboBox1.Font                  = 'Microsoft Sans Serif,10'


$Form.controls.AddRange(@($Label1,$Button1,$TextBox1,$ComboBox1))

$Button1.Add_Click({

$user =$ComboBox1.SelectedItem.ToString()
$TextBox1.Text = Get-ADPrincipalGroupMembership -Identity $user | select name |
  Format-Table -HideTableHeaders | out-string


})

$Form.ShowDialog()