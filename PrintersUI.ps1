
function GenerateForm {
########################################################################
# Created By: Rick Bankers
# Date Modified: 10/12/2015
########################################################################

#region Import the Assemblies
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
#endregion

#region Generated Form Objects
$form1 = New-Object System.Windows.Forms.Form
$progressBar1 = New-Object System.Windows.Forms.ProgressBar
$label4 = New-Object System.Windows.Forms.Label
$panel1 = New-Object System.Windows.Forms.Panel
$AddRemoveCheckBox = New-Object System.Windows.Forms.CheckBox
$ComputerName = New-Object System.Windows.Forms.TextBox
$label3 = New-Object System.Windows.Forms.Label
$label2 = New-Object System.Windows.Forms.Label
$CancelButton = New-Object System.Windows.Forms.Button
$HelpButton = New-Object System.Windows.Forms.Button
$listBox1 = New-Object System.Windows.Forms.ListBox
$buttonListQueues = New-Object System.Windows.Forms.Button
$RunButton = New-Object System.Windows.Forms.Button
$QueueName = New-Object System.Windows.Forms.TextBox
$label1 = New-Object System.Windows.Forms.Label
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion Generated Form Objects


#----------------------------------------------
# Variables
#----------------------------------------------
$PrintServers = @("printserv3")
$PDFHelpFile = ""	# Enter the path to the PDF help file

#----------------------------------------------
#Custom Functions
#----------------------------------------------

#==========================================================================
#Check Spooler GPO is Applied for Windows 7. Remote printer commands will fail if this isn't
#enabled.
#==========================================================================
Function gpoCheck([string]$remoteComputer) {
$osv = (Get-WmiObject -computer $remoteComputer -cl Win32_OperatingSystem).Version  
#Write-Host $remoteComputer $osv        
If ($osv.StartsWith("5")) { 
    # Windows XP is version 5
	Return $true #| out-file C:\stdout.csv -append -encoding ascii 
	}
Else {
    # Windows Vista and later
	try {
		$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $remoteComputer)
		$regkey = $reg.OpenSubkey("SOFTWARE\\Policies\\Microsoft\\Windows NT\\Printers")
		$serialkey = $regkey.GetValue("RegisterSpoolerRemoteRpcEndPoint")
		Return $true #| out-file C:\stdout.csv -append -encoding ascii
	} catch {
		Return $false #| out-file C:\stderr.csv -append -encoding ascii
	}}}

#==========================================================================
#Toggles the Add/Remove Printer Button and Label
#==========================================================================
Function toggleButton {
$listbox1.Items.Clear()
$ComputerName.Text = ""
$QueueName.Text = ""
Switch ($RunButton.Text)
{
"Add Printer(s)" {# $RunButton.BackColor = [System.Drawing.Color]::FromArgb(255,255,0,0)
					$RunButton.Text = "Delete Printer(s)"}
"Delete Printer(s)" {# $RunButton.BackColor = [System.Drawing.Color]::FromArgb(255,0,255,0)
					$RunButton.Text = "Add Printer(s)"}
}					
}

#==========================================================================
#Gets a list of client printers installed
#==========================================================================
Function listClientPrinters {
Switch ($ComputerName.Text.Trim())
{
	"" {Start-Process "rundll32.exe" -ArgumentList "PRINTUI.DLL PrintUIEntry /ge /f$env:temp\PrintersUI.txt" -NoNewWindow -Wait}
	default {
	If ((Test-Connection -Count 1 -ComputerName $ComputerName.Text.Trim() -Quiet) -and (gpoCheck $ComputerName.Text.Trim())){
	Start-Process "rundll32.exe" -ArgumentList "PRINTUI.DLL PrintUIEntry /ge /c\\$($ComputerName.Text.Trim()) /f$env:temp\PrintersUI.txt" -NoNewWindow -Wait
	}Else{
	[System.Windows.Forms.MessageBox]::Show("Connection ERROR!`n`n" + $ComputerName.Text.Trim().ToUpper() + " is unreachable or the print spooler GPO has not been applied to this PC.`n`nPlease check that is online and verify group polcies have been updated.", "PrintersUI")
	Return
	}}
}
$printers1 = Select-String -path "$env:temp\PrintersUI.txt" -pattern "Printer Name"

$listbox1.beginupdate()
	foreach($i in $printers1)
	{
		If ($i) {
		$i = $i -Split("Printer Name:")
		$listbox1.Items.add($i[1].Trim())
		}Else{
		$listbox1.Items.add("No Printers Found")}
	}
$listbox1.EndUpdate()
$RunButton.Enabled = $True
$RunButton.Text = "Delete Printer(s)"

}

#==========================================================================
#Gets a list of client printers installed
#==========================================================================
Function listADPrinters {
#[System.Windows.Forms.MessageBox]::Show("ADD")
#--------------------------------------------------------------------------
#Build List of Printers
$strCategory = "printQueue"
$objDomain = New-Object System.DirectoryServices.DirectoryEntry
$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.SearchRoot = $objDomain
$objSearcher.PageSize = 1000
$objSearcher.Filter = "(objectCategory=$strCategory)"

$colProplist = "printername”, "servername"
foreach ($i in $colPropList){$objSearcher.PropertiesToLoad.Add($i)}

$colResults = $objSearcher.FindAll()
$ProgressBar1.Minimum = 0
$ProgressBar1.Maximum = $colResults.Count
$ProgressBar1.Value = 0

# [System.Windows.Forms.MessageBox]::Show($colResults.Count)
 
#Filter Queue if requested
Switch ($QueueName.Text.Trim()) 
{
	"" {$strQueue = "\w"}
	default {$strQueue = $QueueName.Text.Trim()}
}

$printers = foreach ($objResult in $colResults)
    {$ProgressBar1.Value += 1
	$objPrinter = $objResult.Properties
	If ($objPrinter.servername -match $PrintServers -and $objPrinter.printername -match $strQueue) {
#	[System.Windows.Forms.MessageBox]::Show($objPrinter.servername)
		$Var1 = $objPrinter.printername
		$Var2 = $objPrinter.servername.Split(".")[0]
		"$Var1      ,\\$Var2\"
    }}

$printers1 = $printers | select -Unique | sort-object

$listbox1.beginupdate()
  foreach($i in $printers1)
        {
  $listbox1.Items.add($i)
        }
 $listbox1.EndUpdate()
 $RunButton.Enabled = $True
 $RunButton.Text = "Add Printer(s)"
 
}

#==========================================================================
#Delete Client Printers
#==========================================================================
Function delPrinters
{
	If ($listbox1.SelectedItems.Count -ge 1){
	#RUNDLL32 PRINTUI.DLL PrintUIEntry /gd /c\\%1 /n%2 /Gw
	$ProgressBar1.Minimum = 0
	$ProgressBar1.Maximum = $listbox1.SelectedItems.Count + 1
	$ProgressBar1.Value = 1
foreach ($objItem in $listbox1.SelectedItems)
            {$ProgressBar1.Value += 1
			
Switch ($ComputerName.Text.Trim())
{
	"" {#[System.Windows.Forms.MessageBox]::Show("PRINTUI.DLL PrintUIEntry /gd /n$objItem /Gw /q")
	Start-Process -FilePath "rundll32.exe" -ArgumentList "PRINTUI.DLL PrintUIEntry /gd /n$objItem /Gw /q" -NoNewWindow -Wait
	}
	Default {#[System.Windows.Forms.MessageBox]::Show("PRINTUI.DLL PrintUIEntry /ga /c\\$($ComputerName.Text.Trim()) /gd /n$objItem /Gw /q")
	Start-Process "rundll32.exe" -ArgumentList "PRINTUI.DLL PrintUIEntry /ga /c\\$($ComputerName.Text.Trim()) /gd /n$objItem /Gw /q" -NoNewWindow -Wait
	}
}
}
}Else{
[System.Windows.Forms.MessageBox]::Show("You must select at least one printer from the list.","PrintersUI")
Return
}
Start-Sleep -Seconds 3
[System.Windows.Forms.MessageBox]::Show("Printers Deleted!`nYou can verify by running delete list queues again.","PrintersUI")
}

#==========================================================================
#Add Client Printers
#==========================================================================
Function addPrinters {

If ($listbox1.SelectedItems.Count -ge 1){
	$ProgressBar1.Minimum = 0
	$ProgressBar1.Maximum = $listbox1.SelectedItems.Count + 1
	$ProgressBar1.Value = 1
foreach ($objItem in $listbox1.SelectedItems)
            {$ProgressBar1.Value += 1
			$arrPPath = ($objItem.Replace("      ","").Split(","))
Switch ($ComputerName.Text.Trim())
{
	"" {#[System.Windows.Forms.MessageBox]::Show("PRINTUI.DLL PrintUIEntry /ga /n$($arrPPath[1])$($arrPPath[0]) /Gw /q")
		Start-Process -FilePath "rundll32.exe" -ArgumentList "PRINTUI.DLL PrintUIEntry /ga /n$($arrPPath[1])$($arrPPath[0]) /Gw /q" -NoNewWindow -Wait
}
	Default{If ((Test-Connection -Count 1 -ComputerName $ComputerName.Text.Trim() -Quiet) -and (gpoCheck $ComputerName.Text.Trim())){
				$pc = $ComputerName.Text.Trim()
#				[System.Windows.Forms.MessageBox]::Show("PRINTUI.DLL PrintUIEntry /ga /c\\$($ComputerName.Text.Trim()) /n$($arrPPath[1])$($arrPPath[0]) /Gw /q")
				Start-Process "rundll32.exe" -ArgumentList "PRINTUI.DLL PrintUIEntry /ga /c\\$($ComputerName.Text.Trim()) /n$($arrPPath[1])$($arrPPath[0]) /Gw /q" -NoNewWindow -Wait
				}Else{lllxc v
				[System.Windows.Forms.MessageBox]::Show("Connection ERROR!`n`n" + $ComputerName.Text.Trim().ToUpper() + " is unreachable or the print spooler GPO has not been applied to this PC.`n`nPlease check that is online and verify group polcies have been updated.", "PrintersUI")
				Return
				}}
}
			}

#			(Get-service -ComputerName $ComputerName.Text.Trim() -Name Spooler).Stop()
#			[System.Windows.Forms.MessageBox]::Show("SERvis STOPED")
#			(Get-service -ComputerName $ComputerName.Text.Trim() -Name Spooler).Start()

}Else{
[System.Windows.Forms.MessageBox]::Show("You must select at least one printer from the list.","PrintersUI")
}
Start-Sleep -Seconds 3
[System.Windows.Forms.MessageBox]::Show("Printers have been added to the client!`nYou can verify by running delete printers, list queues.","PrintersUI")
}

#----------------------------------------------
#Generated Event Script Blocks
#----------------------------------------------
#Provide Custom Code for events specified in PrimalForms.
$buttonListQueues_OnClick= 
{
$listbox1.Items.Clear()
If ($AddRemoveCheckBox.Checked)
{
	listClientPrinters
}Else{
	listADPrinters}
}

$handler_label4_Click= 
{
#TODO: Place custom script here

}

$CancelButton_OnClick= 
{
$form1.Close()
}

$HelpButton_OnClick= 
{
If (Test-Path $PDFHelpFile) {
Start-Process -FilePath $PDFHelpFile
}Else{
[System.Windows.Forms.MessageBox]::Show("Help File Not Found!`nMissing:$PDFHelpFile", "PrintersUI")
}
}

$RunButton_OnClick= 
{
Switch ($RunButton.Text)
{
	"Delete Printer(s)" {delPrinters}
	"Add Printer(s)" {addPrinters}
}
}

$OnLoadForm_StateCorrection=
{#Correct the initial state of the form to prevent the .Net maximized form issue
	$form1.WindowState = $InitialFormWindowState
}

#----------------------------------------------
#region Generated Form Code
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 435
$System_Drawing_Size.Width = 407
$form1.ClientSize = $System_Drawing_Size
$form1.DataBindings.DefaultDataSourceUpdateMode = 0
$form1.FormBorderStyle = 1
$form1.MaximizeBox = $False
$form1.MinimizeBox = $False
$form1.Name = "form1"
$form1.Text = "PrintersUI 2.0"
$form1.TopMost = $True

$progressBar1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 11
$System_Drawing_Point.Y = 397
$progressBar1.Location = $System_Drawing_Point
$progressBar1.Name = "progressBar1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 380
$progressBar1.Size = $System_Drawing_Size
$progressBar1.TabIndex = 10

$form1.Controls.Add($progressBar1)

$label4.DataBindings.DefaultDataSourceUpdateMode = 0
$label4.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,0,3,0)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = -1
$System_Drawing_Point.Y = 420
$label4.Location = $System_Drawing_Point
$label4.Name = "label4"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 14
$System_Drawing_Size.Width = 130
$label4.Size = $System_Drawing_Size
$label4.TabIndex = 3
$label4.Text = "Created by Rick Bankers"
$label4.add_Click($handler_label4_Click)

$form1.Controls.Add($label4)


$panel1.BorderStyle = 1
$panel1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 62
$panel1.Location = $System_Drawing_Point
$panel1.Name = "panel1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 329
$System_Drawing_Size.Width = 379
$panel1.Size = $System_Drawing_Size
$panel1.TabIndex = 1

$form1.Controls.Add($panel1)

$AddRemoveCheckBox.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 15
$System_Drawing_Point.Y = 76
$AddRemoveCheckBox.Location = $System_Drawing_Point
$AddRemoveCheckBox.Name = "AddRemoveCheckBox"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 120
$AddRemoveCheckBox.Size = $System_Drawing_Size
$AddRemoveCheckBox.TabIndex = 8
$AddRemoveCheckBox.Text = "List Client Printers"
$AddRemoveCheckBox.UseVisualStyleBackColor = $True

$panel1.Controls.Add($AddRemoveCheckBox)

$ComputerName.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 100
$System_Drawing_Point.Y = 13
$ComputerName.Location = $System_Drawing_Point
$ComputerName.Name = "ComputerName"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 250
$ComputerName.Size = $System_Drawing_Size
$ComputerName.TabIndex = 0

$panel1.Controls.Add($ComputerName)

$label3.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 13
$label3.Location = $System_Drawing_Point
$label3.Name = "label3"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 18
$System_Drawing_Size.Width = 90
$label3.Size = $System_Drawing_Size
$label3.TabIndex = 11
$label3.Text = "Computer Name:"
$label3.TextAlign = 64

$panel1.Controls.Add($label3)

$label2.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 4
$System_Drawing_Point.Y = 44
$label2.Location = $System_Drawing_Point
$label2.Name = "label2"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 90
$label2.Size = $System_Drawing_Size
$label2.TabIndex = 11
$label2.Text = "Queue Name:"
$label2.TextAlign = 64

$panel1.Controls.Add($label2)


$CancelButton.DataBindings.DefaultDataSourceUpdateMode = 0
$CancelButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,0,3,1)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 210
$System_Drawing_Point.Y = 278
$CancelButton.Location = $System_Drawing_Point
$CancelButton.Name = "CancelButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 34
$System_Drawing_Size.Width = 140
$CancelButton.Size = $System_Drawing_Size
$CancelButton.TabIndex = 5
$CancelButton.Text = "Cancel"
$CancelButton.UseVisualStyleBackColor = $True
$CancelButton.add_Click($CancelButton_OnClick)

$panel1.Controls.Add($CancelButton)


$HelpButton.DataBindings.DefaultDataSourceUpdateMode = 0
$HelpButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,0,3,1)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 210
$System_Drawing_Point.Y = 178
$HelpButton.Location = $System_Drawing_Point
$HelpButton.Name = "HelpButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 34
$System_Drawing_Size.Width = 140
$HelpButton.Size = $System_Drawing_Size
$HelpButton.TabIndex = 6
$HelpButton.Text = "Help"
$HelpButton.UseVisualStyleBackColor = $True
$HelpButton.add_Click($HelpButton_OnClick)

$panel1.Controls.Add($HelpButton)

$listBox1.DataBindings.DefaultDataSourceUpdateMode = 0
$listBox1.FormattingEnabled = $True
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 15
$System_Drawing_Point.Y = 100
$listBox1.Location = $System_Drawing_Point
$listBox1.Name = "listBox1"
$listBox1.SelectionMode = 3
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 212
$System_Drawing_Size.Width = 180
$listBox1.Size = $System_Drawing_Size
$listBox1.TabIndex = 3

$panel1.Controls.Add($listBox1)


$buttonListQueues.DataBindings.DefaultDataSourceUpdateMode = 0
$buttonListQueues.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,0,3,1)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 210
$System_Drawing_Point.Y = 100
$buttonListQueues.Location = $System_Drawing_Point
$buttonListQueues.Name = "buttonListQueues"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 34
$System_Drawing_Size.Width = 140
$buttonListQueues.Size = $System_Drawing_Size
$buttonListQueues.TabIndex = 2
$buttonListQueues.Text = "List Queues"
$buttonListQueues.UseVisualStyleBackColor = $True
$buttonListQueues.add_Click($buttonListQueues_OnClick)

$panel1.Controls.Add($buttonListQueues)


$RunButton.DataBindings.DefaultDataSourceUpdateMode = 0
$RunButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,0,3,1)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 210
$System_Drawing_Point.Y = 140
$RunButton.Location = $System_Drawing_Point
$RunButton.Name = "RunButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 32
$System_Drawing_Size.Width = 140
$RunButton.Size = $System_Drawing_Size
$RunButton.TabIndex = 4
$RunButton.Text = "Add Printers(s)"
$RunButton.UseVisualStyleBackColor = $True
$RunButton.add_Click($RunButton_OnClick)

$panel1.Controls.Add($RunButton)

$QueueName.BackColor = [System.Drawing.Color]::FromArgb(255,255,255,255)
$QueueName.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 100
$System_Drawing_Point.Y = 44
$QueueName.Location = $System_Drawing_Point
$QueueName.Name = "QueueName"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 250
$QueueName.Size = $System_Drawing_Size
$QueueName.TabIndex = 1

$panel1.Controls.Add($QueueName)


$label1.BorderStyle = 1
$label1.DataBindings.DefaultDataSourceUpdateMode = 0
$label1.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,0,3,1)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 9
$label1.Location = $System_Drawing_Point
$label1.Name = "label1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 50
$System_Drawing_Size.Width = 379
$label1.Size = $System_Drawing_Size
$label1.TabIndex = 0
$label1.Text = "This utility permanently adds or removes printers from local or remote PCs. Please contact the help desk with any questons."
$label1.TextAlign = 32

$form1.Controls.Add($label1)

#endregion Generated Form Code

#Save the initial state of the form
$InitialFormWindowState = $form1.WindowState
#Init the OnLoad event to correct the initial state of the form
$form1.add_Load($OnLoadForm_StateCorrection)
#Show the Form
$form1.ShowDialog()| Out-Null

} #End Function

#Call the Function
GenerateForm
