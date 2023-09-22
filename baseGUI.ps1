# BASE GUI w/ NO FUNCTIONS

# GUI PACKAGES
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# FLAT/MODERN UI STYLE
[System.Windows.Forms.Application]::EnableVisualStyles()

# BASE FORM
$form = New-Object System.Windows.Forms.Form
$form.Text = 'SCCM Collection Tool'
$form.Size = New-Object System.Drawing.Size(800,670)
$form.FormBorderStyle = 'Fixed3D'
$form.MaximizeBox = $false
$form.StartPosition = 'CenterScreen'
$form.Add_Load({
	$form.Activate()
})

# LABELS, BUTTONS AND TEXT BOXES
$labelComp1 = New-Object System.Windows.Forms.Label
$labelComp1.Location = New-Object System.Drawing.Size(80,5)
$labelComp1.Size = New-Object System.Drawing.Size(300,30)
$labelComp1.Text = 'Computer with MISSING software:'
$labelComp1.Font = ‘Segoe UI,12’
$labelComp1.ForeColor = [System.Drawing.Color]::Black
$labelComp1.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($labelComp1)

$boxComp1 = New-Object System.Windows.Forms.TextBox
$boxComp1.Location = New-Object System.Drawing.Point(10,35)
$boxComp1.Size = New-Object System.Drawing.Size(380,30)
$boxComp1.Text = " "
$boxComp1.AutoSize = $false
$boxComp1.Font = ‘Segoe UI,11’
$boxComp1.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$form.Controls.Add($boxComp1)

$labelIP1 = New-Object System.Windows.Forms.Label
$labelIP1.Location = New-Object System.Drawing.Size(10,80)
$labelIP1.Size = New-Object System.Drawing.Size(300,25)
$labelIP1.Text = 'Detected IP: '
$labelIP1.Font = [System.Drawing.Font]::new("Segoe UI",11,[System.Drawing.FontStyle]::Bold)
$labelIP1.ForeColor = [System.Drawing.Color]::Black
$labelIP1.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($labelIP1)

$boxCollection1 = New-Object System.Windows.Forms.RichTextBox
$boxCollection1.Location = New-Object System.Drawing.Point(10,115)
$boxCollection1.Size = New-Object System.Drawing.Size(380,400)
$boxCollection1.AutoSize = $false
$boxCollection1.Multiline = $true
$boxCollection1.WordWrap = $true
$boxCollection1.Font = ‘Segoe UI,11’
$boxCollection1.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$form.Controls.Add($boxCollection1)

$labelComp2 = New-Object System.Windows.Forms.Label
$labelComp2.Location = New-Object System.Drawing.Size(460,5)
$labelComp2.Size = New-Object System.Drawing.Size(300, 30)
$labelComp2.Text = 'Computer with CORRECT software:'
$labelComp2.Font = ‘Segoe UI,12’
$labelComp2.ForeColor = [System.Drawing.Color]::Black
$labelComp2.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($labelComp2)

$boxComp2 = New-Object System.Windows.Forms.TextBox
$boxComp2.Location = New-Object System.Drawing.Point(400,35)
$boxComp2.Size = New-Object System.Drawing.Size(370,30)
$boxComp2.Text = " "
$boxComp2.AutoSize = $false
$boxComp2.Font = ‘Segoe UI,11’
$boxComp2.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$form.Controls.Add($boxComp2)

$labelIP2 = New-Object System.Windows.Forms.Label
$labelIP2.Location = New-Object System.Drawing.Size(400,80)
$labelIP2.Size = New-Object System.Drawing.Size(300,25)
$labelIP2.Text = 'Detected IP: '
$labelIP2.Font = [System.Drawing.Font]::new("Segoe UI",11,[System.Drawing.FontStyle]::Bold)
$labelIP2.ForeColor = [System.Drawing.Color]::Black
$labelIP2.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($labelIP2)

$boxCollection2 = New-Object System.Windows.Forms.RichTextBox
$boxCollection2.Location = New-Object System.Drawing.Point(400,115)
$boxCollection2.Size = New-Object System.Drawing.Size(370,400)
$boxCollection2.AutoSize = $false
$boxCollection2.Multiline = $true
$boxCollection2.WordWrap = $true
$boxCollection2.Font = ‘Segoe UI,11’
$boxCollection2.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$form.Controls.Add($boxCollection2)

$buttonCompare = New-Object System.Windows.Forms.Button
$buttonCompare.Location = New-Object System.Drawing.Point(320,560)
$buttonCompare.Size = New-Object System.Drawing.Size(150,30)
$buttonCompare.Text = 'COMPARE'
$buttonCompare.Font = ‘Segoe UI,12’
$buttonCompare.ForeColor = [System.Drawing.Color]::White
$buttonCompare.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonCompare.FlatAppearance.BorderColor = [System.Drawing.Color]::DeepSkyBlue
$buttonCompare.BackColor = [System.Drawing.Color]::DeepSkyBlue
$buttonCompare.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::DarkBlue
$buttonCompare.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::DodgerBlue
$form.Controls.Add($buttonCompare)
$buttonCompare.Add_Click($compare) #Adds function to button

$labelResults = New-Object System.Windows.Forms.Label
$labelResults.Location = New-Object System.Drawing.Point(10,600)
$labelResults.Size = New-Object System.Drawing.Size(900,30)
$labelResults.Text = ""
$labelResults.Font = [System.Drawing.Font]::new("Segoe UI",11,[System.Drawing.FontStyle]::Bold+[System.Drawing.FontStyle]::Italic)
$labelResults.ForeColor = [System.Drawing.Color]::Black
$labelResults.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($labelResults)

# POPUP MESSAGE WHEN USER CLOSES APP
$form.Add_Closing({param($sender,$e)
    $result = [System.Windows.Forms.MessageBox]::Show(`
        "Are you sure you want to exit?", `
        "Close", [System.Windows.Forms.MessageBoxButtons]::YesNo)
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
		return
		exit
	}
	
	if ($result -eq [System.Windows.Forms.DialogResult]::No) {
		$e.Cancel = $true
	}
})

# WHAT HAPPENS WHEN USER PRESSES ENTER
$form.AcceptButton = $buttonCompare

# SHOW FORM
$Null = $form.ShowDialog()