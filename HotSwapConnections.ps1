#This script was designed by Steve Campbell and provided by PowerBI.tips
#BE WARNED this will alter Power BI files so please make sure you know what you are doing, and always back up your files!
#This is not supported by Microsoft and changes to future file structures could cause this code to break

#--------------- Released 8/23/2020 ---------------
#--- By Steve Campbell provided by PowerBI.tips --




################################################################################ 

###///Check version////###

################################################################################



 

#Current Version
$version = '1.1.0'



#Check for update
$response = Invoke-WebRequest -URI https://raw.githubusercontent.com/PowerBISteve/powerbiscripts/master/versionhistory
if ($response.Content.Trim() -ne $version.Trim()){


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$updater = New-Object System.Windows.Forms.Form
$updater.Text = 'Update Available'
$updater.Size = New-Object System.Drawing.Size(350,200)
$updater.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(15,80)
$okButton.Size = New-Object System.Drawing.Size(120,40)
$okButton.Text = 'Click to open Link'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$updater.AcceptButton = $okButton
$updater.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(165,80)
$cancelButton.Size = New-Object System.Drawing.Size(120,40)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$updater.CancelButton = $cancelButton
$updater.Controls.Add($cancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(250,120)
$label.Text = 'There is a new update available on PowerBI.tips
Please update to maintain functionality'
$updater.Controls.Add($label)


$updater.Topmost = $true
$result = $updater.ShowDialog()
if($result -eq [System.Windows.Forms.DialogResult]::OK){
Start-Process https://powerbi.tips/2020/08/hot-swap-report-connections-external-tools/
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
Function Disconnect-PBIX([string]$inputpath){

    # Open zip and find the particular file (assumes only one inside the Zip file)
    Add-Type -assembly  System.IO.Compression.FileSystem
    $zipfile = ($inputpath).Substring(0,($inputpath).Length-4) + "zip"
    Rename-Item -Path $inputpath -NewName  $zipfile
    $zip =  [System.IO.Compression.ZipFile]::Open($zipfile,"Update")
    $contents = $zip.Entries.Where({$_.name -eq 'DataModel'})

    $zip.Entries.Where({$_.name -eq 'SecurityBindings'}) | % { $_.Delete() } -ErrorAction SilentlyContinue
    $zip.Entries.Where({$_.name -eq 'Connections'}) | % { $_.Delete() } -ErrorAction SilentlyContinue

    # Write the changes and close the zip file
    $zip.Dispose()

    Rename-Item -Path $zipfile -NewName $inputpath  
    Invoke-Item $inputpath  
}












#Add Connections
Function Connect-PBIX([string]$inputpath){

    $ConStr = '{"Version":1,"Connections":[{"Name":"EntityDataSource","ConnectionString":"Data Source=' + $port + ';Initial Catalog=' + $cat + ';Cube=Model","ConnectionType":"analysisServicesDatabaseLive"}]}'

    # Open zip and find the particular file (assumes only one inside the Zip file)
    Add-Type -assembly  System.IO.Compression.FileSystem
    $zipfile = ($inputpath).Substring(0,($inputpath).Length-4) + "zip"
    Rename-Item -Path $inputpath -NewName  $zipfile
    $zip =  [System.IO.Compression.ZipFile]::Open($zipfile,"Update")


    $contents = $zip.Entries.Where({$_.name -eq 'Connections'})
    if ($contents.Count -gt 0 ){
    
    # Overwrite the contents file
   
    $desiredFile = [System.IO.StreamWriter]($contents).Open()
    $desiredFile.BaseStream.SetLength(0)
    $desiredFile.Write($ConStr)
    $desiredFile.Flush()
    $desiredFile.Close()

    }

    else
    {

    # Create new contents file and add it
    $TempPath =$env:USERPROFILE + "\temp\HotSwap\"

    If(!(test-path $TempPath))
    {New-Item -ItemType Directory -Force -Path $TempPath}
    $NewConFile = $TempPath + '\Connections'
    if (Test-Path $NewConFile) { Remove-Item $NewConFile}
    New-Item -Path $TempPath -Name "Connections" -ItemType "file" -Value $ConStr -Force
    
    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip,$NewConFile,"Connections","Optimal") | Out-Null
    
    Remove-Item  $NewConFile
    }


    $zip.Entries.Where({$_.name -eq 'SecurityBindings'}) | % { $_.Delete() } -ErrorAction SilentlyContinue

    # Write the changes and close the zip file
    $zip.Dispose()

    Rename-Item -Path $zipfile -NewName $inputpath  
    Invoke-Item $inputpath  
    }


#Choose pbix funtion
Function Get-FileName($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "PBIX (*.pbix)| *.pbix"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}
#Error check function
function IsFileLocked([string]$filePath){
    Rename-Item $filePath $filePath -ErrorVariable errs -ErrorAction SilentlyContinue
    return ($errs.Count -ne 0)
}
#Final pick function
Function Pick-PBIX(){
    try {$pathn = Get-FileName}
    catch { "Incompatible File" }

    #Check for errors
    If([string]::IsNullOrEmpty($pathn )){            
        exit } 

    elseif ( IsFileLocked($pathn) ){
        [System.Windows.MessageBox]::Show('File is already open - please close and try again')
        exit } 

    else{ $pathn}
}
#Final pick and copy function
Function Pick-Copy-PBIX([string]$suffix){
       try {$pathn = Get-FileName}
       catch { "Incompatible File" }

       $pathnnew = ($pathn).toString().Replace('.pbix', $suffix + '.pbix')

       if ( [string]::IsNullOrEmpty($pathnnew) -Or (Test-Path $pathnnew) ){
           DO{

             $nn +=1
             $pathnnew = ($pathnnew).Substring(0,($pathnnew).Length-5) + "(" + ($nn) +  ").pbix"
             if (Test-Path $pathnnew)
               {$pass = 0}
             else
               {$pass = 1}
             } Until ($pass -eq 1)}
       else
           {}


       Copy-Item $pathn -Destination $pathnnew  -Force
       return $pathnnew
}





################################################################################ 

###///Create Form Form////###

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
$tabPage5.Size = New-Object System.Drawing.Size(266, 163)
$tabPage5.TabIndex = 4
$tabPage5.Text = "Settings"
$tabPage5.UseVisualStyleBackColor = $true
#
# textBox1
#
$textBox1.Location = New-Object System.Drawing.Point(50, 75)
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
$tabPage3.Location = New-Object System.Drawing.Point(4, 22)
$tabPage3.Name = "tabPage3"
$tabPage3.Size = New-Object System.Drawing.Size(266, 163)
$tabPage3.TabIndex = 2
$tabPage3.Text = "Help"
$tabPage3.UseVisualStyleBackColor = $true
#
# label2
#
$label2.Location = New-Object System.Drawing.Point(3, 13)
$label2.Name = "label2"
$label2.Size = New-Object System.Drawing.Size(250, 48)
$label2.TabIndex = 0
$label2.Text = "Connect to report will open the selected report live connected to current file"
#
# label3
#
$label3.Location = New-Object System.Drawing.Point(3, 61)
$label3.Name = "label3"
$label3.Size = New-Object System.Drawing.Size(250, 37)
$label3.TabIndex = 1
$label3.Text = "Remove connections will remove the connections or data model of the selected file" +
""

#
# linkLabel1
#
$linkLabel1.AutoSize = $true
$linkLabel1.Location = New-Object System.Drawing.Point(3, 142)
$linkLabel1.Name = "linkLabel1"
$linkLabel1.Size = New-Object System.Drawing.Size(237, 13)
$linkLabel1.TabIndex = 3
$linkLabel1.TabStop = $true
$linkLabel1.Text = "Click here for the documentation on PowerBI.tips"
$LinkLabel2.add_Click({[system.Diagnostics.Process]::start("https://powerbi.tips/2020/08/hot-swap-report-connections-external-tools/")})
#
# tabPage2
#
$tabPage2.Controls.Add($linkLabel3)
$tabPage2.Controls.Add($reml_cop)
$tabPage2.Controls.Add($reml_ovr)
$tabPage2.Location = New-Object System.Drawing.Point(4, 22)
$tabPage2.Name = "tabPage2"
$tabPage2.Padding = New-Object System.Windows.Forms.Padding(3)
$tabPage2.Size = New-Object System.Drawing.Size(266, 163)
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
$tabPage1.Location = New-Object System.Drawing.Point(4, 22)
$tabPage1.Name = "tabPage1"
$tabPage1.Padding = New-Object System.Windows.Forms.Padding(3)
$tabPage1.Size = New-Object System.Drawing.Size(266, 163)
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
$tabControl1.Controls.Add($tabPage3)
$tabControl1.Controls.Add($tabPage5)
$tabControl1.Location = New-Object System.Drawing.Point(0, 0)
$tabControl1.Name = "tabControl1"
$tabControl1.SelectedIndex = 0
$tabControl1.Size = New-Object System.Drawing.Size(274, 189)
$tabControl1.TabIndex = 4
#
# label1
#
$label1.Location = New-Object System.Drawing.Point(8, 26)
$label1.Name = "label1"
$label1.Size = New-Object System.Drawing.Size(239, 46)
$label1.TabIndex = 6
$label1.Text = "Type the suffix to add to the report when using any COPY activity"
#
# label4
#
$label4.Location = New-Object System.Drawing.Point(3, 109)
$label4.Name = "label4"
$label4.Size = New-Object System.Drawing.Size(250, 24)
$label4.TabIndex = 4
$label4.Text = "Overwrite will modify file, Copy will duplicate first"
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
$LinkLabel2.add_Click({[system.Diagnostics.Process]::start("https://powerbi.tips/2020/08/hot-swap-report-connections-external-tools")})
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
$LinkLabel3.add_Click({[system.Diagnostics.Process]::start("https://powerbi.tips/2020/08/hot-swap-report-connections-external-tools")})
#
# HotSwap
#
$HotSwap.ClientSize = New-Object System.Drawing.Size(275, 187)
$HotSwap.Controls.Add($tabControl1)
$HotSwap.Controls.Add($menuStrip1)
$iconBase64      = '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'
$iconBytes       = [Convert]::FromBase64String($iconBase64)
$stream          = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
$stream.Write($iconBytes, 0, $iconBytes.Length);
$iconImage       = [System.Drawing.Image]::FromStream($stream, $true)
$HotSwap.Icon       = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())
$HotSwap.MainMenuStrip = $menuStrip1
$HotSwap.MaximizeBox = $false
$HotSwap.MinimizeBox = $false
$HotSwap.Name = "HotSwap"
$HotSwap.Text = "Hot Swap Report Connections"


#####Add functions to buttons
$con_cop.Add_Click({
$inputpth = Pick-Copy-PBIX $textBox1.Text
Connect-PBIX $inputpth
$HotSwap.Close()
})

$con_ovr.Add_Click({
$inputpth = Pick-PBIX
Connect-PBIX $inputpth
$HotSwap.Close()
})

$reml_ovr.Add_Click({
$inputpth = Pick-PBIX
Disconnect-PBIX $inputpth
$HotSwap.Close()
})

$reml_cop.Add_Click({
$inputpth = Pick-Copy-PBIX $textBox1.Text
Disconnect-PBIX $inputpth
$HotSwap.Close()
})

function OnFormClosing_HotSwap{ 
	# $this parameter is equal to the sender (object)
	# $_ is equal to the parameter e (eventarg)

	# The CloseReason property indicates a reason for the closure :
	#   if (($_).CloseReason -eq [System.Windows.Forms.CloseReason]::UserClosing)

	#Sets the value indicating that the event should be canceled.
	($_).Cancel= $False
}

$HotSwap.Add_FormClosing( { OnFormClosing_HotSwap} )

$HotSwap.Add_Shown({$HotSwap.Activate()})
$ModalResult=$HotSwap.ShowDialog()

# Release the Form
$HotSwap.Dispose()
