Import-Module ActiveDirectory

if (-not (Get-Module -Name AdmPwd.PS -ErrorAction SilentlyContinue)) {
    try {
        Import-Module AdmPwd.PS -ErrorAction Stop
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "LAPS PowerShell module not found. Please install it first.",
            "Module Missing",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        exit 1
    }
}

function DoesExists {
    param (
        [string]$ComputerName
    )
    try {
        $null = Get-ADComputer -Identity $ComputerName -ErrorAction Stop
        return $true
    }catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]{
        return $false
    }
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Dark Mode Color Palette
$darkBg = [System.Drawing.Color]::FromArgb(32, 33, 35)       # Main background
$darkText = [System.Drawing.Color]::FromArgb(234, 236, 240)  # Primary text
$darkAccent = [System.Drawing.Color]::FromArgb(98, 154, 255) # Accent color
$darkBorder = [System.Drawing.Color]::FromArgb(62, 68, 81)    # Borders
$darkButton = [System.Drawing.Color]::FromArgb(52, 58, 70)    # Button background
$darkButtonHover = [System.Drawing.Color]::FromArgb(62, 68, 81) # Button hover
$darkInputBg = [System.Drawing.Color]::FromArgb(43, 45, 48)   # Input fields
$darkReadOnlyBg = [System.Drawing.Color]::FromArgb(38, 40, 44) # Read-only fields

# Main form
$lapsForm = New-Object Windows.Forms.Form
$lapsForm.Text = "LAPS"
$lapsForm.Width = 420
$lapsForm.Height = 220
$lapsForm.BackColor = $darkBg
$lapsForm.ForeColor = $darkText
$lapsForm.Font = New-Object Drawing.Font("Segoe UI", 10)
$lapsForm.StartPosition = "CenterScreen"


$lapsForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None


# Add your own close/minimize buttons
$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Text = "X"
$closeButton.Location = New-Object System.Drawing.Point(360, 10)
$closeButton.Add_Click({ $lapsForm.Close() })
$lapsForm.Controls.Add($closeButton)


# PC Name Label
$lapsPCNameL = New-Object Windows.Forms.Label
$lapsPCNameL.Text = "PC Name"
$lapsPCNameL.Font = New-Object Drawing.Font("Segoe UI", 12, [Drawing.FontStyle]::Bold)
$lapsPCNameL.Location = New-Object Drawing.Point(15, 20)
$lapsPCNameL.AutoSize = $true
$lapsPCNameL.ForeColor = $darkAccent
$lapsForm.Controls.Add($lapsPCNameL)

# PC Name Textbox
$lapsPCNameT = New-Object Windows.Forms.TextBox
$lapsPCNameT.Font = New-Object Drawing.Font("Segoe UI", 12)
$lapsPCNameT.Location = New-Object Drawing.Point(15, 47)
$lapsPCNameT.Width = 250
$lapsPCNameT.Height = 25
$lapsPCNameT.BackColor = $darkInputBg
$lapsPCNameT.ForeColor = $darkText
$lapsPCNameT.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$lapsPCNameT.ShortcutsEnabled = $true
$lapsForm.Controls.Add($lapsPCNameT)

# Copy popup
$lapsCopyP = New-Object Windows.Forms.Label
$lapsCopyP.Font = New-Object Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Italic)
$lapsCopyP.Location = New-Object Drawing.Point(15, 82)
$lapsCopyP.AutoSize = $true
$lapsCopyP.ForeColor = $darkAccent
$lapsForm.Controls.Add($lapsCopyP)

# Submit Button
$lapsSubmit = New-Object Windows.Forms.Button
$lapsSubmit.Text = "Get Password"
$lapsSubmit.Font = New-Object Drawing.Font("Segoe UI", 10, [Drawing.FontStyle]::Bold)
$lapsSubmit.Location = New-Object Drawing.Point(275, 47)
$lapsSubmit.Width = 120
$lapsSubmit.Height = 28
$lapsSubmit.BackColor = $darkButton
$lapsSubmit.ForeColor = $darkText
$lapsSubmit.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$lapsSubmit.FlatAppearance.BorderColor = $darkBorder
$lapsSubmit.FlatAppearance.MouseOverBackColor = $darkButtonHover
$lapsSubmit.FlatAppearance.MouseDownBackColor = $darkAccent
$lapsSubmit.Add_Click({
    $lapsPwd.text = Get-LapsADPassword -Identity $lapsPCNameT.Text -asplaintext | Select -ExpandProperty Password
    Set-Clipboard -Value $lapsPwd.text
    $lapsCopyP.Text = "Password copied to clipboard!"
})
$lapsForm.Controls.Add($lapsSubmit)

# Password Field
$lapsPwd = New-Object Windows.Forms.TextBox
$lapsPwd.Font = New-Object Drawing.Font("Segoe UI", 12)
$lapsPwd.Location = New-Object Drawing.Point(15, 120)
$lapsPwd.Width = 380
$lapsPwd.Height = 25
$lapsPwd.BackColor = $darkReadOnlyBg
$lapsPwd.ForeColor = $darkText
$lapsPwd.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$lapsPwd.ReadOnly = $true
$lapsForm.Controls.Add($lapsPwd)

# Enter key handler
$lapsPCNameT.Add_KeyDown({
    if($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        if(DoesExists -ComputerName $lapsPCNameT.Text) {
            $OU = Get-ADComputer $lapsPCNameT.Text -Properties distinguishedname | Select -ExpandProperty distinguishedname
            if($OU -like "*CN=Computers,*") {
                $lapsPwd.Text = "PC in Default OU"
                $lapsCopyP.Text = ""
            } else {
                $lapsPwd.text = Get-LapsADPassword -Identity $lapsPCNameT.Text -asplaintext | Select -ExpandProperty Password
                Set-Clipboard -Value $lapsPwd.text
                $lapsCopyP.Text = "Password copied to clipboard!"
            }
        } else {
            $lapsPwd.Text = "PC not found in AD"
            $lapsCopyP.Text = ""
        }
    }
})

# CTRL+A handler
$lapsPCNameT.Add_KeyDown({
    if($_.Control -and $_.KeyCode -eq "A"){
        $_.SuppressKeyPress = $true
        $this.SelectAll()       
    }
})

# Show form
[System.Windows.Forms.Application]::EnableVisualStyles()
$lapsForm.Add_Shown($lapsForm.Activate())
$lapsForm.ShowDialog()
