# powershell.exe -ExecutionPolicy Bypass -File "Get-Laps.ps1"
Import-Module ActiveDirectory

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

# Main box
$lapsForm = New-Object Windows.Forms.Form
$lapsForm.Text = "LAPS"
$lapsForm.Width = 400
$lapsForm.Height = 200

# PC name text form
# First label for the textbox
$lapsPCNameL = New-Object Windows.Forms.label
$lapsPCNameL.Text = "PC Name"
$lapsPCNameL.font = New-Object Drawing.font("Arial",11,[Drawing.FontStyle]::Bold)
$lapsPCNameL.Location = New-Object Drawing.Point(10, 20)
$lapsPCNameL.AutoSize = $true
$lapsForm.Controls.add($lapsPCNameL)

# Now textbox itself
$lapsPCNameT = New-Object Windows.Forms.textbox
$lapsPCNameT.font = New-Object Drawing.font("Arial",11)
$lapsPCNameT.Location = New-Object Drawing.Point(10, 45)
$lapsPCNameT.Width = 200
$lapsPCNameT.height = 15
$lapsPCNameT.AutoSize = $true
$lapsPCNameT.ShortcutsEnabled = $true
# Copy popup
$lapsCopyP = New-Object Windows.Forms.label
$lapsCopyP.font = New-Object Drawing.font("Arial",11)
$lapsCopyP.Location = New-Object Drawing.Point(10, 74)
$lapsCopyP.AutoSize = $true
$lapsForm.Controls.add($lapsCopyP)

# Add enter as a keydown 
$lapsPCNameT.Add_KeyDown({
    if($_.KeyCode -eq [system.windows.forms.keys]::enter) {
        if(DoesExists -ComputerName $lapsPCNameT.Text) {
            $OU = Get-ADcomputer $lapsPCNameT.Text -Properties distinguishedname | Select -ExpandProperty distinguishedname
            if($OU -like "*CN=Computers,*") {
                $lapsPwd.Text = "PC in Default OU"
                $lapsCopyP.Text = ""
            }else {
                $lapsPwd.text = Get-LapsADPassword -Identity $lapsPCNameT.Text -asplaintext | Select -ExpandProperty Password
                Set-Clipboard -Value $lapsPwd.text
                $lapsCopyP.Text = "Password copied to clipboard!"
            }
        }else {
            $lapsPwd.Text = "PC not in AD"
            $lapsCopyP.Text = ""
        }
    }
})
# CTRL + A as a valid shortcut 
$lapsPCNameT.Add_KeyDown({
    if($_.Control -and $_.KeyCode -eq "A"){
        $_.SuppressKeyPress = $true
        $this.SelectAll()       
    }
})



$lapsForm.Controls.add($lapsPCNameT)



# Submit button
$lapsSubmit = New-Object Windows.forms.button
$lapsSubmit.Text = "Sumbit"
$lapsSubmit.font = New-Object Drawing.font("Arial",11,[Drawing.FontStyle]::Bold)
$lapsSubmit.Location = New-Object Drawing.Point(220, 40)
$lapsSubmit.Width = 50
$lapsSubmit.height = 15
$lapsSubmit.AutoSize = $true
$lapsSubmit.Add_click({
    $lapsPwd.text = Get-LapsADPassword -Identity $lapsPCNameT.Text -asplaintext | Select -ExpandProperty Password
    Set-Clipboard -Value $lapsPwd.text
}) 
$lapsForm.Controls.add($lapsSubmit)

# Password field
$lapsPwd = New-Object Windows.Forms.textbox
$lapsPwd.font = New-Object Drawing.font("Arial",11)
$lapsPwd.Location = New-Object Drawing.Point(10, 100)
$lapsPwd.Width = 300
$lapsPwd.height = 15
$lapsPwd.AutoSize = $true
$lapsPwd.ReadOnly = $true
$lapsForm.Controls.add($lapsPwd)



$lapsForm.Add_shown($lapsForm.Activate())
$lapsForm.ShowDialog()

