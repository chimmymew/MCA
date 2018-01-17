Attribute VB_Name = "MCA_Module"

Option Explicit
Private Declare Function CreateFile _
 Lib "kernel32.dll" Alias "CreateFileA" ( _
 ByVal lpFileName As String, _
 ByVal dwDesiredAccess As Long, _
 ByVal dwShareMode As Long, _
 lpSecurityAttributes As SECURITY_ATTRIBUTES, _
 ByVal dwCreationDisposition As Long, _
 ByVal dwFlagsAndAttributes As Long, _
 ByVal hTemplateFile As Long) As Long
 
Private Declare Function CloseHandle _
 Lib "kernel32.dll" ( _
 ByVal hObject As Long) As Long
 
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
 
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public COMDesc() As String
Public COMPorts() As String


 
Public Function COMAvailable(iPortNum As Integer) As Boolean
    Dim hCOM As Long
    Dim ret As Long
    Dim sec As SECURITY_ATTRIBUTES
 
    'try to open the COM port
    hCOM = CreateFile("COM" & iPortNum & "", 0, FILE_SHARE_READ + FILE_SHARE_WRITE, _
     sec, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hCOM = -1 Then
        COMAvailable = False
    Else
        COMAvailable = True
        'close the COM port
        ret = CloseHandle(hCOM)
    End If
End Function


Public Function PortDescription() As String()
Dim WMIServ As Object
Dim Items As Object
Dim SubItems As Object
Dim PortList As String

    Set WMIServ = GetObject("winmgmts:\\.\root\cimv2")
    Set Items = WMIServ.ExecQuery("Select * from Win32_PnPEntity", , 48)

    For Each SubItems In Items
        If InStr(SubItems.Caption, "(COM") > 0 Then
            PortList = PortList & SubItems.Name & vbCrLf
        End If
    Next
    If Len(PortList) > 0 Then
        PortList = Left$(PortList, Len(PortList) - 2)
    End If
    PortDescription = Split(PortList, vbCrLf)
    
End Function

Public Function Ports() As String()
Dim WMIServ As Object
Dim Items As Object
Dim SubItems As Object
Dim PortNumbers As String
Dim NumStart As Integer
Dim NumEnd As Integer

    Set WMIServ = GetObject("winmgmts:\\.\root\cimv2")
    Set Items = WMIServ.ExecQuery("Select * from Win32_PnPEntity", , 48)

    For Each SubItems In Items
        If InStr(SubItems.Caption, "(COM") > 0 Then
            NumStart = InStr(SubItems.Caption, "(COM") + 4
            NumEnd = InStr(NumStart, SubItems.Caption, ")")
            PortNumbers = PortNumbers & Mid$(SubItems.Caption, NumStart, NumEnd - NumStart) & " "
            PortNumbers = PortNumbers & vbCrLf
        End If
    Next
    Ports = Split(PortNumbers, vbCrLf)
End Function


