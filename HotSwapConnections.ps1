#This script was designed by Steve Campbell and provided by PowerBI.tips
#BE WARNED this will alter Power BI files so please make sure you know what you are doing, and always back up your files!
#This is not supported by Microsoft and changes to future file structures could cause this code to break

#--------------- Released 8/23/2020 ---------------
#--- By Steve Campbell provided by PowerBI.tips --




################################################################################ 

###///Check version////###

################################################################################





#Current Version
$version = '1.1.2'

# Help Page
Function Open-HelpPage() {
    Start-Process 'https://powerbi.tips/2020/08/hot-swap-report-connections-external-tools/'
}

#Check for update
$response = Invoke-WebRequest -URI https://raw.githubusercontent.com/PowerBISteve/powerbiscripts/master/versionhistory
if ($response.Content.Trim() -ne $version.Trim()) {


    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $updater = New-Object System.Windows.Forms.Form
    $updater.Text = 'Update Available'
    $updater.Size = New-Object System.Drawing.Size(350, 200)
    $updater.StartPosition = 'CenterScreen'

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(15, 80)
    $okButton.Size = New-Object System.Drawing.Size(120, 40)
    $okButton.Text = 'Click to open Link'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $updater.AcceptButton = $okButton
    $updater.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(165, 80)
    $cancelButton.Size = New-Object System.Drawing.Size(120, 40)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $updater.CancelButton = $cancelButton
    $updater.Controls.Add($cancelButton)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(250, 120)
    $label.Text = 'There is a new update available on PowerBI.tips. Please update to maintain functionality.'
    $updater.Controls.Add($label)


    $updater.Topmost = $true
    $result = $updater.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        Open-HelpPage
        exit
    }
}



################################################################################ 

###///Get inputs from Power BI////###

################################################################################

#Input arguments from Power BI
$port = $args[0]
$cat = $args[1]



################################################################################ 

###///Construct Functions////###

################################################################################


#Function to Modify files
Function Disconnect-PBIX([string]$inputpath) {

    # Open zip and find the particular file (assumes only one inside the Zip file)
    Add-Type -assembly  System.IO.Compression.FileSystem
    $zipfile = ($inputpath).Substring(0, ($inputpath).Length - 4) + "zip"
    Rename-Item -Path $inputpath -NewName  $zipfile
    $zip = [System.IO.Compression.ZipFile]::Open($zipfile, "Update")

    $zip.Entries.Where({ $_.name -eq 'SecurityBindings' }) | ForEach-Object { $_.Delete() } -ErrorAction SilentlyContinue
    $zip.Entries.Where({ $_.name -eq 'Connections' }) | ForEach-Object { $_.Delete() } -ErrorAction SilentlyContinue

    # Write the changes and close the zip file
    $zip.Dispose()

    Rename-Item -Path $zipfile -NewName $inputpath  
    Invoke-Item $inputpath  
}



#Add Connections
Function Connect-PBIX([string]$inputpath) {

    $ConStr = '{"Version":1,"Connections":[{"Name":"EntityDataSource","ConnectionString":"Data Source=' + $port + ';Initial Catalog=' + $cat + ';Cube=Model","ConnectionType":"analysisServicesDatabaseLive"}]}'

    # Open zip and find the particular file (assumes only one inside the Zip file)
    Add-Type -assembly  System.IO.Compression.FileSystem
    $zipfile = ($inputpath).Substring(0, ($inputpath).Length - 4) + "zip"
    Rename-Item -Path $inputpath -NewName  $zipfile
    $zip = [System.IO.Compression.ZipFile]::Open($zipfile, "Update")


    $contents = $zip.Entries.Where({ $_.name -eq 'Connections' })
    if ($contents.Count -gt 0 ) {
    
        # Overwrite the contents file
        $desiredFile = [System.IO.StreamWriter]($contents).Open()
        $desiredFile.BaseStream.SetLength(0)
        $desiredFile.Write($ConStr)
        $desiredFile.Flush()
        $desiredFile.Close()

    }

    else {

        # Create new contents file and add it
        $TempPath = $env:USERPROFILE + "\temp\HotSwap\"

        if (!(Test-Path $TempPath))
        { New-Item -ItemType Directory -Force -Path $TempPath }
        $NewConFile = $TempPath + '\Connections'
        if (Test-Path $NewConFile) { Remove-Item $NewConFile }
        New-Item -Path $TempPath -Name "Connections" -ItemType "file" -Value $ConStr -Force
    
        [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $NewConFile, "Connections", "Optimal")
    
        Remove-Item  $NewConFile
    }


    $zip.Entries.Where({ $_.name -eq 'SecurityBindings' }) | ForEach-Object { $_.Delete() } -ErrorAction SilentlyContinue

    # Write the changes and close the zip file
    $zip.Dispose()

    Rename-Item -Path $zipfile -NewName $inputpath  
    Invoke-Item $inputpath  
}


#Choose pbix funtion
Function Get-FileName($initialDirectory) {
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "PBIX (*.pbix)| *.pbix"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}
#Error check function
Function Get-FileLockStatus([string]$filePath) {
    Rename-Item $filePath $filePath -ErrorVariable errs -ErrorAction SilentlyContinue
    return ($errs.Count -ne 0)
}

Function Show-AlreadyOpenMessage() {
    [System.Windows.MessageBox]::Show('File is already open - please close and try again')
}

#Final pick function
Function Open-PBIX() {
    try { $pathn = Get-FileName }
    catch { "Incompatible File" }


    #Check for errors
    if ([string]::IsNullOrEmpty($pathn)) { return } 

    elseif (Get-FileLockStatus($pathn)) {
        Show-AlreadyOpenMessage
        return 
    } 

    else { $pathn }
}

#Final pick and copy function
Function Copy-PBIX([string]$suffix) {
    try { $pathn = Get-FileName }
    catch { "Incompatible File" }

    #Check for errors
    if ([string]::IsNullOrEmpty($pathn)) { return } 

    elseif (Get-FileLockStatus($pathn)) {
        Show-AlreadyOpenMessage
        return
    }

    $pathnnew = ($pathn).toString().Replace('.pbix', $suffix + '.pbix')

    if ( [string]::IsNullOrEmpty($pathnnew) -Or (Test-Path $pathnnew) ) {
        do {
            $nn += 1
            $pathnnew = ($pathnnew).Substring(0, ($pathnnew).Length - 5) + "(" + ($nn) + ").pbix"
            if (Test-Path $pathnnew) {
                $pass = 0
            }
            else {
                $pass = 1
            }
        } until ($pass -eq 1)
    }
    else {}


    Copy-Item $pathn -Destination $pathnnew  -Force
    return $pathnnew
}





################################################################################ 

###///Create Form ////###

################################################################################



# Loading external assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


# Creation of the components.
$HotSwap = New-Object System.Windows.Forms.Form

$menuStrip1 = New-Object System.Windows.Forms.MenuStrip
$tabPage5 = New-Object System.Windows.Forms.TabPage
$textBox1 = New-Object System.Windows.Forms.TextBox
$label6 = New-Object System.Windows.Forms.Label
$warnlabel = New-Object System.Windows.Forms.Label
$warnlabel1 = New-Object System.Windows.Forms.Label
$label5 = New-Object System.Windows.Forms.Label
$tabPage3 = New-Object System.Windows.Forms.TabPage
$label2 = New-Object System.Windows.Forms.Label
$label3 = New-Object System.Windows.Forms.Label
$linkLabel1 = New-Object System.Windows.Forms.LinkLabel
$tabPage2 = New-Object System.Windows.Forms.TabPage
$reml_ovr = New-Object System.Windows.Forms.Button
$reml_cop = New-Object System.Windows.Forms.Button
$tabPage1 = New-Object System.Windows.Forms.TabPage
$con_ovr = New-Object System.Windows.Forms.Button
$con_cop = New-Object System.Windows.Forms.Button
$tabControl1 = New-Object System.Windows.Forms.TabControl
$label1 = New-Object System.Windows.Forms.Label
$label4 = New-Object System.Windows.Forms.Label
$linkLabel2 = New-Object System.Windows.Forms.LinkLabel
$linkLabel3 = New-Object System.Windows.Forms.LinkLabel
#
# menuStrip1
#
$menuStrip1.Location = New-Object System.Drawing.Point(0, 0)
$menuStrip1.Name = "menuStrip1"
$menuStrip1.Size = New-Object System.Drawing.Size(275, 24)
$menuStrip1.TabIndex = 3
$menuStrip1.Text = "menuStrip1"
#
# tabPage5
#
$tabPage5.Controls.Add($label1)
$tabPage5.Controls.Add($label6)
$tabPage5.Controls.Add($textBox1)
$tabPage5.Location = New-Object System.Drawing.Point(4, 22)
$tabPage5.Name = "tabPage5"
$tabPage5.Size = New-Object System.Drawing.Size(266, 193)
$tabPage5.TabIndex = 4
$tabPage5.Text = "Settings"
$tabPage5.UseVisualStyleBackColor = $true
#
# textBox1
#
$textBox1.Location = New-Object System.Drawing.Point(60, 75)
$textBox1.Name = "textBox1"
$textBox1.Size = New-Object System.Drawing.Size(113, 20)
$textBox1.TabIndex = 4
$textBox1.Text = "_Report"
#
# label6
#
$label6.AutoSize = $true
$label6.Location = New-Object System.Drawing.Point(8, 78)
$label6.Name = "label6"
$label6.Size = New-Object System.Drawing.Size(36, 13)
$label6.TabIndex = 5
$label6.Text = "Suffix:"
#
# tabPage3
#
$tabPage3.Controls.Add($label4)
$tabPage3.Controls.Add($linkLabel1)
$tabPage3.Controls.Add($label3)
$tabPage3.Controls.Add($label2)
$tabPage3.Controls.Add($label5)
$tabPage3.Location = New-Object System.Drawing.Point(4, 22)
$tabPage3.Name = "tabPage3"
$tabPage3.Size = New-Object System.Drawing.Size(266, 193)
$tabPage3.TabIndex = 2
$tabPage3.Text = "Help"
$tabPage3.UseVisualStyleBackColor = $true
#
# label2
#
$label2.Location = New-Object System.Drawing.Point(3, 3)
$label2.Name = "label2"
$label2.Size = New-Object System.Drawing.Size(250, 37)
$label2.TabIndex = 0
$label2.Text = "Connect: Open selected file and connect it to dataset in current file"
#
# label3
#
$label3.Location = New-Object System.Drawing.Point(3, 40)
$label3.Name = "label3"
$label3.Size = New-Object System.Drawing.Size(250, 37)
$label3.TabIndex = 1
$label3.Text = "Remove: Clear all connections and/or data models from selected file"

#
# linkLabel1
#
$linkLabel1.AutoSize = $true
$linkLabel1.Location = New-Object System.Drawing.Point(1, 120)
$linkLabel1.Name = "linkLabel1"
$linkLabel1.Size = New-Object System.Drawing.Size(237, 13)
$linkLabel1.TabIndex = 4
$linkLabel1.TabStop = $true
$linkLabel1.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8)
$linkLabel1.Text = "Click here for documentation on PowerBI.tips"
$LinkLabel1.add_Click({ Open-HelpPage })
#
# label5
#
$label5.AutoSize = $true
$label5.Location = New-Object System.Drawing.Point(1, 140)
$label5.Name = "label5"
$label5.Size = New-Object System.Drawing.Size(237, 13)
$label5.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 7.5)
$label5.TabIndex = 5
$label5.Text = "Hot Swap Report Connections Version " + $version
#
# warnlabel
#
$warnlabel.Location = New-Object System.Drawing.Point(3, 150)
$warnlabel.Name = "warnlabel"
$warnlabel.Size = New-Object System.Drawing.Size(255, 50)
$warnlabel.TabIndex = 0
$warnlabel.Text = "THIS MODIFIES FILES IN A WAY UNSOPPORTED BY MICROSOFT. ALWAYS BACK UP ALL FILES."



$warnlabel1.Location = New-Object System.Drawing.Point(3, 150)
$warnlabel1.Name = "warnlabel1"
$warnlabel1.Size = New-Object System.Drawing.Size(250, 50)
$warnlabel1.TabIndex = 9
$warnlabel1.Text = "THIS MODIFIES FILES IN A WAY UNSOPPORTED BY MICROSOFT. ALWAYS BACK UP ALL FILES."
#
# tabPage2
#



$tabPage2.Controls.Add($linkLabel3)
$tabPage2.Controls.Add($reml_cop)
$tabPage2.Controls.Add($reml_ovr)
$tabPage2.Controls.Add($warnlabel1)
$tabPage2.Location = New-Object System.Drawing.Point(4, 22)
$tabPage2.Name = "tabPage2"
$tabPage2.Padding = New-Object System.Windows.Forms.Padding(3)
$tabPage2.Size = New-Object System.Drawing.Size(266, 193)
$tabPage2.TabIndex = 1
$tabPage2.Text = "Remove"
$tabPage2.UseVisualStyleBackColor = $true
#
# reml_ovr
#
$reml_ovr.BackColor = [System.Drawing.Color]::Wheat
$reml_ovr.Location = New-Object System.Drawing.Point(54, 20)
$reml_ovr.Name = "reml_ovr"
$reml_ovr.Size = New-Object System.Drawing.Size(158, 50)
$reml_ovr.TabIndex = 1
$reml_ovr.Text = "Overwrite and remove live connections"
$reml_ovr.UseVisualStyleBackColor = $false
#
# reml_cop
#
$reml_cop.BackColor = [System.Drawing.Color]::Wheat
$reml_cop.Location = New-Object System.Drawing.Point(54, 90)
$reml_cop.Name = "reml_cop"
$reml_cop.Size = New-Object System.Drawing.Size(158, 50)
$reml_cop.TabIndex = 4
$reml_cop.Text = "Copy and remove live connections"
$reml_cop.UseVisualStyleBackColor = $false
#
# tabPage1
#


$tabPage1.Controls.Add($linkLabel2)
$tabPage1.Controls.Add($con_cop)
$tabPage1.Controls.Add($con_ovr)
$tabPage1.Controls.Add($warnlabel)
$tabPage1.Location = New-Object System.Drawing.Point(4, 22)
$tabPage1.Name = "tabPage1"
$tabPage1.Padding = New-Object System.Windows.Forms.Padding(3)
$tabPage1.Size = New-Object System.Drawing.Size(266, 193) 
$tabPage1.TabIndex = 0
$tabPage1.Text = "Connect"
$tabPage1.UseVisualStyleBackColor = $true
#
# con_ovr
#
$con_ovr.Location = New-Object System.Drawing.Point(54, 20)
$con_ovr.Name = "con_ovr"
$con_ovr.Size = New-Object System.Drawing.Size(158, 50)
$con_ovr.TabIndex = 4
$con_ovr.Text = "Overwrite and connect"
$con_ovr.UseVisualStyleBackColor = $true

#
# con_cop
#
$con_cop.Location = New-Object System.Drawing.Point(54, 90)
$con_cop.Name = "con_cop"
$con_cop.Size = New-Object System.Drawing.Size(158, 50)
$con_cop.TabIndex = 5
$con_cop.Text = "Copy and connect"
$con_cop.UseVisualStyleBackColor = $true
#
# tabControl1
#
$tabControl1.Controls.Add($tabPage1)
$tabControl1.Controls.Add($tabPage2)
$tabControl1.Controls.Add($tabPage5)
$tabControl1.Controls.Add($tabPage3)
$tabControl1.Location = New-Object System.Drawing.Point(0, 0)
$tabControl1.Name = "tabControl1"
$tabControl1.SelectedIndex = 0
$tabControl1.Size = New-Object System.Drawing.Size(274, 239)
$tabControl1.TabIndex = 4
#
# label1
#
$label1.Location = New-Object System.Drawing.Point(8, 26)
$label1.Name = "label1"
$label1.Size = New-Object System.Drawing.Size(239, 46)
$label1.TabIndex = 6
$label1.Text = "Type the suffix to add to the new filename when using any COPY activity"
#
# label4
#
$label4.Location = New-Object System.Drawing.Point(3, 77)
$label4.Name = "label4"
$label4.Size = New-Object System.Drawing.Size(250, 37)
$label4.TabIndex = 3
$label4.Text = "Overwrite: Modify original file`nCopy: Duplicate file, then modify"
#
# linkLabel2
#
$linkLabel2.AutoSize = $true
$linkLabel2.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 12.25)
$linkLabel2.Location = New-Object System.Drawing.Point(242, 6)
$linkLabel2.Name = "linkLabel2"
$linkLabel2.Size = New-Object System.Drawing.Size(18, 20)
$linkLabel2.TabIndex = 6
$linkLabel2.TabStop = $true
$linkLabel2.Text = "?"
$LinkLabel2.add_Click({ Open-HelpPage })
#
# linkLabel3
#
$linkLabel3.AutoSize = $true
$linkLabel3.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 12.25)
$linkLabel3.Location = New-Object System.Drawing.Point(242, 6)
$linkLabel3.Name = "linkLabel3"
$linkLabel3.Size = New-Object System.Drawing.Size(18, 20)
$linkLabel3.TabIndex = 7
$linkLabel3.TabStop = $true
$linkLabel3.Text = "?"
$LinkLabel3.add_Click({ Open-HelpPage })
#
# HotSwap
#
$HotSwap.ClientSize = New-Object System.Drawing.Size(275, 240)
$HotSwap.Controls.Add($tabControl1)
$HotSwap.Controls.Add($menuStrip1)
$iconBase64 = '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'
$iconBytes = [Convert]::FromBase64String($iconBase64)
$stream = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
$stream.Write($iconBytes, 0, $iconBytes.Length);
[void][System.Drawing.Image]::FromStream($stream, $true)
$HotSwap.Icon = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())
$HotSwap.MainMenuStrip = $menuStrip1
$HotSwap.MaximizeBox = $false
$HotSwap.MinimizeBox = $false
$HotSwap.Name = "HotSwap"
$HotSwap.Text = "Hot Swap Report Connections"


#####Add functions to buttons
$con_cop.Add_Click({
        $inputpth = Copy-PBIX $textBox1.Text
        if ([string]::IsNullOrEmpty($inputpth ))
        { return } 
        else {
            Connect-PBIX $inputpth
            $HotSwap.Close()
        } })

$con_ovr.Add_Click({
        $inputpth = Open-PBIX
        if ([string]::IsNullOrEmpty($inputpth ))
        { return }
        else {
            Connect-PBIX $inputpth
            $HotSwap.Close()
        } })

$reml_ovr.Add_Click({
        $inputpth = Open-PBIX
        if ([string]::IsNullOrEmpty($inputpth ))
        { return }
        else {
            Disconnect-PBIX $inputpth
            $HotSwap.Close()
        } })

$reml_cop.Add_Click({
        $inputpth = Copy-PBIX $textBox1.Text
        if ([string]::IsNullOrEmpty($inputpth ))
        { return } 
        else {
            Disconnect-PBIX $inputpth
            $HotSwap.Close()
        } })

function Invoke-HotSwap_OnFormClosing { 
    # $this parameter is equal to the sender (object)
    # $_ is equal to the parameter e (eventarg)

    # The CloseReason property indicates a reason for the closure :
    # if (($_).CloseReason -eq [System.Windows.Forms.CloseReason]::UserClosing)

    #Sets the value indicating that the event should be canceled.
    ($_).Cancel = $False
}

$HotSwap.Add_FormClosing({ Invoke-HotSwap_OnFormClosing })

$HotSwap.Add_Shown({ $HotSwap.Activate() })
[void][system.windows.forms.application]::run($HotSwap) #.ShowDialog()

# Release the Form
$HotSwap.Dispose()
