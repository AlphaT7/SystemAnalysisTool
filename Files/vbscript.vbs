
'// S Y S T E M    I N F O R M A T I O N

'Displays the script version
 
document.getElementById("htaversion").innerHTML = "About || Version " & SystemAnalysis.version
document.getElementById("aboutversion").innerHTML = "About || Version " & SystemAnalysis.version

' // --- SET THE WMI OBJECT --- //

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
                & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
 
' // --- M A N U F A C T U R E R    I N F O --- //
 
Set x = objWMIService.ExecQuery _
                ("Select * from Win32_ComputerSystem")
 
For Each y In x
document.getElementById("manufacturer").innerHTML= y.Manufacturer
document.getElementById("model").innerHTML= y.Model
 
Next
 
' // --- O P E R A T I N G    S Y S T E M    I N F O --- //
 
Set x = objWMIService.ExecQuery _
                ("Select * from Win32_OperatingSystem")
 
For Each y In x

document.getElementById("os").innerHTML= y.Caption
Dim sp
sp = y.CSDVersion
If (sp <> "") Then
document.getElementById("servicepack").innerHTML= y.CSDVersion
Else
document.getElementById("servicepack").innerHTML= "No service pack found"
End If
document.getElementById("ostype").innerHTML= GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'").AddressWidth & " bit"

Next
 
' // --- B I O S    I N F O --- //
 
Set x = objWMIService.ExecQuery _
                ("Select * from Win32_BIOS")
 
For Each y In x

document.getElementById("serialnumber").innerHTML= y.SerialNumber
document.getElementById("biosversion").innerHTML= y.SMBIOSBIOSVersion

Dim biosage
Dim biosdate
Dim rawbiosdate
Dim realbiosage

rawbiosdate = y.ReleaseDate
biosdate = WMIDateStringToDate(rawbiosdate)
biosage = DateDiff("d",biosdate,Date)/365
realbiosage = DateDiff("d",biosdate,Date) \ 365  & " yrs " & DateDiff("d",biosdate,Date) Mod 365 & " days"

document.getElementById("biosdate").innerHTML= biosdate

Next

'//Check if bios is older than 5 years to see if should install SD

Dim currentyear
currentyear = DatePart("yyyy",Date)
If (biosage >= 5) Then

document.getElementById("biosage").innerHTML= "5 yrs or Older"
document.getElementById("biosage").className= "biosstop"
document.getElementById("biosage").title= realbiosage

ElseIf (biosage < 5) Then
document.getElementById("biosage").innerHTML= "Newer than 5 yrs &#10004;"
document.getElementById("biosage").className= "biosgood"
document.getElementById("biosage").title= realbiosage
End If

'//Check if bios is older than 1 year to see if needs a check disk

If (biosage > 1) Then

document.getElementById("checkdisk").innerHTML= "Older than 1 yr, check disk required"
document.getElementById("checkdisk").className= "checkdiskcheck"
document.getElementById("checkdisk").title= realbiosage

ElseIf (biosage <= 1) Then

document.getElementById("checkdisk").innerHTML= "Newer than 1 yr &#10004;"
document.getElementById("checkdisk").className= "checkdiskgood"
document.getElementById("checkdisk").title= realbiosage
End If

Function WMIDateStringToDate(dtmWMIDate)

    If Not IsNull(dtmWMIDate) Then

WMIDateStringToDate = (Mid(dtmWMIDate,5,2) & "/" & Mid(dtmWMIDate,7,2) & "/" & Left(dtmWMIDate,4))

    End If

End Function

 
' // --- R E C O V E R Y    T O O L    B U T T O N S --- //

' // Displays the manufacturer recovery information, and highlightes the button selected
Sub displayPanel(manufacturer)

Dim i 
Dim buttons
Dim button_id
Dim panel_id

i = 0
buttons = Array("default","dell","hp","lenovo","toshiba","asus","acer","gateway","sony","samsung","emachines")

Do While i <= 10
button_id = buttons(i) &  "_rbutton"
panel_id = buttons(i) & "_options"
If buttons(i) <> manufacturer Then
document.getElementById(button_id).classNAME = "recovery_button_dim"
document.getElementById(panel_id).style.display = "none"
Else
document.getElementById(button_id).classNAME = "recovery_button_lit"
document.getElementById(panel_id).style.display = "block"
End If
i = i + 1
Loop
End Sub


'// Windows 8 Recovery Tool shortcut
Sub Windows8Recovery
        Set objShell = CreateObject("Wscript.Shell")
        strRDPath = objShell.ExpandEnvironmentStrings( "%SystemRoot%" ) & "\sysnative\RecoveryDrive.exe"
        Set objShellExec = CreateObject("Shell.Application")
objShellExec.ShellExecute "powershell.exe", strRDPath, "", "runas", 1                
End Sub

' // --- S Y S T E M    T O O L    B U T T O N S --- //

Sub aboutOpen

	document.getElementById("outer_credits_bg").style.display = "block"
	document.getElementById("credits").style.display = "block"

End Sub

Sub aboutClose

	document.getElementById("outer_credits_bg").style.display = "none"
	document.getElementById("credits").style.display = "none"

End Sub

'Set objWMIService = GetObject("winmgmts:" _
'                & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Sub displayRecovery
document.getElementById("systeminfo").style.display="none"
document.getElementById("systemrecovery").style.display="block"
document.getElementById("systeminfobutton").innerHTML="System Info"
document.getElementById("systemrecoverybutton").innerHTML="- System Backup -"
End Sub

Sub displayInfo
document.getElementById("systemrecovery").style.display="none"
document.getElementById("systeminfo").style.display="block"
document.getElementById("systemrecoverybutton").innerHTML="System Backup"
document.getElementById("systeminfobutton").innerHTML="- System Info -"
End Sub
 
Sub MsInfo
'// Open MsInfo32
Set oShell = CreateObject("Wscript.Shell")
oShell.Run "MsInfo32.Exe"
End Sub
 
Sub DManagement
'// Open Computer Management
Set oShell = CreateObject("Wscript.Shell")
oShell.Run "diskmgmt.msc"
End Sub

Sub UEFI
'// Run NYL UEFI check program
Set objShell = CreateObject("WScript.Shell")
Set objExec = objShell.Exec("CheckUefi_nyl_64bit\CheckUefi_nyl_64bit.exe")
End Sub
 
 
Sub OPAL
'// Check if C:\ exists
Set fso = CreateObject("Scripting.FileSystemObject")
   
If (fso.FolderExists("c:\\sd\")) Then

'// Run CMD as Admin
Set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "cmd.exe", "/k c: & cd\sd & hwemngr.exe -l", "", "runas", 1

Else

'// Copy SD folder to C:\\ and then open CMD and run script
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CopyFolder "SD", "c:\"

'// Run CMD as Admin
Set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "cmd.exe", "/k c: & cd\sd & hwemngr.exe -l", "", "runas", 1

End If
End Sub
 
Sub ARPrograms
'// Open Add/Remove Programs
Set oShell = CreateObject("Wscript.Shell")
oShell.Run "control appwiz.cpl"
End Sub
 
Sub CMD
'// Run CMD as Admin and display wmic csproduct list full
Set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "cmd.exe", "/k wmic csproduct list full"
End Sub

Sub DeviceManager()
'// Open Device Manager
Set oShell = CreateObject("Wscript.Shell")
oShell.Run "devmgmt.msc"
End Sub

Sub disableSleep()
Set oShell = CreateObject("Wscript.Shell")
oShell.Run "Files\powersettings.bat"
End Sub

Sub DisableAutomaticRepair()
Set objShell = CreateObject("WScript.Shell")
strCommand = objShell.ExpandEnvironmentStrings("%comspec%")

Set objShellApp = CreateObject("Shell.Application")
objShellApp.ShellExecute strCommand, "/K bcdedit /set {default} recoveryenabled No", , "RunAs", 1
End Sub

Sub windowsUpdates()
Set oShell = CreateObject("Wscript.Shell")
oShell.Run "wuapp.exe"
End Sub

'// C H E C K    I F    L A P T O P    O R    D E S K T O P

If IsLaptop( "." ) Then
   document.getELementById("systemtype").innerHTML = "Laptop"
Else
    document.getELementById("systemtype").innerHTML = "Desktop"
End If


Function IsLaptop( myComputer )
' This Function checks if a computer has a battery pack.
' One can assume that a computer with a battery pack is a laptop.
'
' Argument:
' myComputer   [string] name of the computer to check,
'                       or "." for the local computer
' Return value:
' True if a battery is detected, otherwise False
'
' Written by Rob van der Woude
' http://www.robvanderwoude.com
    On Error Resume Next
    Set objWMIService = GetObject( "winmgmts://" & myComputer & "/root/cimv2" )
    Set colItems = objWMIService.ExecQuery( "Select * from Win32_Battery", , 48 )
    IsLaptop = False
    For Each objItem in colItems
        IsLaptop = True
    Next
    If Err Then Err.Clear
    On Error Goto 0
End Function
