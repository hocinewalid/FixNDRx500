Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
$form = New-Object System.Windows.Forms.Form
$form.Text = "IMCEAEX to X.500 Converter"
$form.Size = New-Object System.Drawing.Size(400, 250)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$form.MaximizeBox = $false
$imceaexLabel = New-Object System.Windows.Forms.Label
$imceaexLabel.Location = New-Object System.Drawing.Point(10, 20)
$imceaexLabel.Size = New-Object System.Drawing.Size(100, 20)
$imceaexLabel.Text = "IMCEAEX:"
$form.Controls.Add($imceaexLabel)

$imceaexTextBox = New-Object System.Windows.Forms.TextBox
$imceaexTextBox.Location = New-Object System.Drawing.Point(120, 20)
$imceaexTextBox.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($imceaexTextBox)
$mailboxLabel = New-Object System.Windows.Forms.Label
$mailboxLabel.Location = New-Object System.Drawing.Point(10, 50)
$mailboxLabel.Size = New-Object System.Drawing.Size(100, 20)
$mailboxLabel.Text = "Mailbox:"
$form.Controls.Add($mailboxLabel)

$mailboxTextBox = New-Object System.Windows.Forms.TextBox
$mailboxTextBox.Location = New-Object System.Drawing.Point(120, 50)
$mailboxTextBox.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($mailboxTextBox)
$addButton = New-Object System.Windows.Forms.Button
$addButton.Location = New-Object System.Drawing.Point(120, 90)
$addButton.Size = New-Object System.Drawing.Size(120, 30)
$addButton.Text = "Add X.500 Address"
$addButton.Add_Click({
    $IMCEAEX = $imceaexTextBox.Text
    $Mailbox = $mailboxTextBox.Text

    if ($IMCEAEX -notmatch "^IMCEAEX") {
        [System.Windows.Forms.MessageBox]::Show("Sorry, your IMCEAEX string must begin with IMCEAEX", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $X500 = $IMCEAEX -replace "IMCEAEX-","X500:" -replace "_","/" -replace "\+20"," " -replace "\+28","(" -replace "\+29",")" -replace "\+2E","." -replace "%3D","=" -split "@" | Select-Object -First 1

    Try {
        Set-Mailbox -Identity $Mailbox -EmailAddresses @{add="$X500"}
        [System.Windows.Forms.MessageBox]::Show("X.500 address successfully added to the mailbox: $Mailbox", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    Catch {
        [System.Windows.Forms.MessageBox]::Show("An error occurred while adding the X.500 address to the mailbox: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$form.Controls.Add($addButton)
$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Location = New-Object System.Drawing.Point(250, 90)
$closeButton.Size = New-Object System.Drawing.Size(120, 30)
$closeButton.Text = "Close"
$closeButton.Add_Click({ $form.Close() })
$form.Controls.Add($closeButton)
$form.ShowDialog()
