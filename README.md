<div align="center">

## Disk Information Class


</div>

### Description

This class enables you to get various information about the disks / drives in your PC.

It uses API's to retrive the information : Drive type, volume label, available space, space used etc.
 
### More Info
 
There is a single public method called GetAllDriveInfo which accepts a drive letter as a parameter (A:). It calls all the private methods in the class and populates all the properties.

In order to determine the availabale drives, simply read the AllDrives property which is populated when you instanciate the class.

Properties for all the drive information.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Alon Hirsch](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/alon-hirsch.md)
**Level**          |Advanced
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/alon-hirsch-disk-information-class__1-38448/archive/master.zip)

### API Declarations

Some - see the code


### Source Code

```
VERSION 1.0 CLASS
Attribute VB_Name = "clsDiskSpace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ***********************************************************************
'
' CLASS : clsDiskSpace.cls
'
' PURPOSE : functions for accessing disk / file information
'
' WRITTEN BY : Alon Hirsch
'
' COMPANY : Debtpack (Pty) Ltd. - Development
'
' DATE : 10 May 2002
'
' ***********************************************************************
Option Explicit
DefInt A-Z
Private m_cDiskSize As Currency
Private m_cDiskUsed As Currency
Private m_cDiskFree As Currency
Private m_fFreePercent As Single
Private m_lSerial As Long
Private m_sVolume As String
Private m_sFileSystem As String
Private m_sAllDrives As String
Private m_sDriveType As String
Private Const FS_CASE_IS_PRESERVED = &H2
Private Const FS_CASE_SENSITIVE = &H1
Private Const FS_UNICODE_STORED_ON_DISK = &H4
Private Const FS_PERSISTENT_ACLS = &H8
Private Const FS_FILE_COMPRESSION = &H10
Private Const FS_VOL_IS_COMPRESSED = &H8000
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Sub GetAllDriveInfo(ByVal sDrive As String)
 ' ensure that there is a \ at the end of the drive
 If Right$(sDrive, 1) <> "\" Then sDrive = sDrive & "\"
 GetDiskSpace sDrive
 GetVolumeInfo sDrive
 GetTypeOfDrive sDrive
End Sub
Private Sub GetDiskSpace(ByVal sDrive As String)
 ' this will calculate the drive specs for the drive and report total size,
 ' size used and size available as well as used %
 Dim lResult As Long
 Dim lSectorPerCluster As Long
 Dim lBytesPerSector As Long
 Dim lFreeClusters As Long
 Dim lTotalClusters As Long
 ' call the API and get the information
 lResult = GetDiskFreeSpace(sDrive, lSectorPerCluster, lBytesPerSector, lFreeClusters, _
        lTotalClusters)
 ' perform the various calculations required
 m_cDiskSize = CCur(lTotalClusters) * CCur(lSectorPerCluster) * CCur(lBytesPerSector)
 m_cDiskFree = CCur(lFreeClusters) * CCur(lSectorPerCluster) * CCur(lBytesPerSector)
 m_cDiskUsed = m_cDiskSize - m_cDiskFree
 If m_cDiskSize <> 0 Then
  m_fFreePercent = m_cDiskFree / m_cDiskSize * 100
 Else
  m_fFreePercent = 0
 End If
End Sub
Public Property Get cDiskSize() As Currency
 cDiskSize = m_cDiskSize
End Property
Public Property Get cDiskUsed() As Currency
 cDiskUsed = m_cDiskUsed
End Property
Public Property Get cDiskFree() As Currency
 cDiskFree = m_cDiskFree
End Property
Public Property Get fFreePercent() As Single
 fFreePercent = m_fFreePercent
End Property
Private Sub GetTypeOfDrive(ByVal sDrive As String)
 Select Case GetDriveType(sDrive)
  Case Is = 2
   m_sDriveType = "Removable"
  Case Is = 3
   m_sDriveType = "Fixed"
  Case Is = 4
   m_sDriveType = "Remote"
  Case Is = 5
   m_sDriveType = "CD-Rom"
  Case Is = 6
   m_sDriveType = "RAM Disk"
  Case Else
   m_sDriveType = "Unknown"
 End Select
End Sub
Private Sub GetVolumeInfo(ByVal sDrive As String)
 Dim sBuffer As String
 Dim sSysName As String
 Dim lResult As Long
 Dim lSysFlags As Long
 Dim lComponentLength As Long
 sBuffer = String$(256, 0)
 sSysName = String$(256, 0)
 lResult = GetVolumeInformation(sDrive, sBuffer, 255, m_lSerial, lComponentLength, lSysFlags, sSysName, 255)
 If lResult = 0 Then
  ' unable to get information
  m_sVolume = "Unable to retrieve information"
  m_sFileSystem = "Unable to retrieve information"
  m_lSerial = 0
 Else
  ' retrieve the information
  m_sVolume = Left$(sBuffer, InStr(sBuffer, Chr$(0)) - 1)
  m_sFileSystem = Left$(sSysName, InStr(sSysName, Chr$(0)) - 1)
 End If
End Sub
Public Property Get lSerial() As Long
 lSerial = m_lSerial
End Property
Public Property Get sAllDrives() As String
 sAllDrives = m_sAllDrives
End Property
Public Property Get sDriveType() As String
 sDriveType = m_sDriveType
End Property
Public Property Get sSerial() As String
 sSerial = Hex$(m_lSerial)
End Property
Public Property Get sVolume() As String
 sVolume = m_sVolume
End Property
Public Property Get sFileSystem() As String
 sFileSystem = m_sFileSystem
End Property
Private Sub Class_Initialize()
 ' determine which drives are available on this PC
 Dim sTemp As String
 Dim iPos As Integer
 sTemp = String$(2048, 0)
 Call GetLogicalDriveStrings(2047, sTemp)
 ' now build up the string containing a comma delimited version of all the drives
 m_sAllDrives = ""
 Do
  iPos = InStr(sTemp, Chr$(0))
  If iPos > 1 Then
   ' we have a drive letter - extract it from the buffer
   If m_sAllDrives = "" Then
    m_sAllDrives = Left$(sTemp, iPos - 1)
   Else
    m_sAllDrives = m_sAllDrives & "," & Left$(sTemp, iPos - 1)
   End If
   sTemp = Mid$(sTemp, iPos + 1)
  End If
 Loop Until iPos <= 1
End Sub
```

