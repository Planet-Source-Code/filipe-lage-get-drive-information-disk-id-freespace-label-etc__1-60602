<div align="center">

## Get drive information \(Disk ID, FreeSpace, Label, etc\)


</div>

### Description

These functions will provide you quick, fail-safe information about a specific disk letter you specify (ex: "C:\" )<br>

GetDriveSerialID - Returns the serial number of a drive partition (if available)<br>

GetDriveFreeSpace - Returns the free space of the specified drive (if available)<br>

GetDriveSize - Returns the total drive space (if available)<br>

GetDriveUsedSpace - Returns the used disk space of the specified drive (if available)<br>

GetDriveLabel - Returns the volume label of the specified drive<br>

Some functions are declared as variant instead of longs, avoiding the 2GB limit.<br>
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Filipe Lage](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/filipe-lage.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/filipe-lage-get-drive-information-disk-id-freespace-label-etc__1-60602/archive/master.zip)





### Source Code

I know that this isn't new, and this information could be optained in many different ways (using API, etc).<br>
I'm just providing this info for beginers.<br>
<br>
Instructions:<br>
Put this code in a new form and add one button (named "Command1")<br>
<br>------------ ~ --------------<br><br>
Public Function GetDriveSerialID(diskletter As String) As String<br>
On Error Resume Next<br>
Set c = CreateObject("scripting.filesystemobject")<br>
GetDriveSerialID = Hex(c.drives(Left(diskletter, 1)).serialnumber)<br>
Set c = Nothing<br>
End Function<br>
<br>
Public Function GetDriveFreeSpace(diskletter As String) As Variant<br>
On Error Resume Next<br>
Set c = CreateObject("scripting.filesystemobject")<br>
GetDriveFreeSpace = 0 ' default<br>
GetDriveFreeSpace = c.drives(Left(diskletter, 1)).freespace<br>
Set c = Nothing<br>
End Function<br>
<br>
Public Function GetDriveSize(diskletter As String) As Variant<br>
On Error Resume Next<br>
Set c = CreateObject("scripting.filesystemobject")<br>
GetDriveSize = 0 ' default<br>
GetDriveSize = c.drives(Left(diskletter, 1)).totalsize<br>
Set c = Nothing<br>
End Function<br>
<br>
Public Function GetDriveUsedSpace(diskletter As String) As Variant<br>
On Error Resume Next<br>
GetDriveUsedSpace = GetDriveSize(diskletter) - GetDriveFreeSpace(diskletter)<br>
End Function<br>
<br>
Public Function GetDriveLabel(diskletter As String) As String<br>
On Error Resume Next<br>
Set c = CreateObject("scripting.filesystemobject")<br>
GetDriveLabel = c.drives(Left(diskletter, 1)).volumename<br>
Set c = Nothing<br>
End Function<br>
<br>
Private Sub Command1_Click()<br>
UseHD = "C:"<br>
Debug.Print "Volume Label: " & GetDriveLabel("c")<br>
Debug.Print "Disk serial number: " & GetDriveSerialID("c")<br>
Debug.Print "Free space: " & GetDriveFreeSpace("s") & " bytes"<br>
Debug.Print "Total drive size: " & GetDriveSize("c") & " bytes"<br>
Debug.Print "Total used space: " & GetDriveUsedSpace("c") & " bytes"<br>
End Sub<br>
<br>
------------ ~ --------------<br>
<br>
Vote if you wish. I appreciate it<br>
// FCLage<br>

