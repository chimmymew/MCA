VERSION 5.00
Object = "{49F811F7-6005-4AAF-AE00-9D98766A6E26}#1.0#0"; "NTGraph.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frm_MCA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mew's Microsys mini Multi Channel Analyzer"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9690
   Icon            =   "frm_MCA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check3 
      Caption         =   "Oscilloscope mode"
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6180
      TabIndex        =   11
      Top             =   6180
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show dots distribution"
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6180
      TabIndex        =   10
      Top             =   5580
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog dialog 
      Left            =   9060
      Top             =   5820
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "csv"
      DialogTitle     =   "Export statistic"
      FileName        =   "stat001"
      Filter          =   "*.csv"
   End
   Begin VB.CommandButton cmd_Export 
      Caption         =   "Export"
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   8
      Top             =   5580
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Show bar graph"
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6180
      TabIndex        =   7
      Top             =   5880
      Width           =   1995
   End
   Begin VB.TextBox txt_interval 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Text            =   "2000"
      Top             =   5940
      Width           =   1275
   End
   Begin VB.CommandButton cmd_integrate 
      Caption         =   "Integrate"
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4860
      TabIndex        =   3
      Top             =   6060
      Width           =   1215
   End
   Begin VB.CommandButton cmd_Connect 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4860
      TabIndex        =   2
      Top             =   5580
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   900
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   5580
      Width           =   3855
   End
   Begin MSCommLib.MSComm mCom 
      Left            =   240
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   1
      InputLen        =   1
      OutBufferSize   =   1
      RThreshold      =   1
      BaudRate        =   115200
      EOFEnable       =   -1  'True
   End
   Begin NTGRAPHLib.NTGraph dgraph 
      Height          =   5475
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9555
      _Version        =   65536
      _ExtentX        =   16854
      _ExtentY        =   9657
      _StockProps     =   194
      Caption         =   "Scan"
      ShowGrid        =   -1  'True
      TrackMode       =   1
      FrameStyle      =   1
      ControlFrameColor=   12632256
      GridColor       =   2631720
      Cursor          =   1999
      FormatAxisLeft  =   "%.0f"
      FormatAxisBottom=   "%.0f"
      ElementLineColor=   0
      ElementPointColor=   0
      ElementName     =   "Element-0"
      BeginProperty TickFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty LabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty IdentFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PlotAreaPicture =   "frm_MCA.frx":0442
      ControlFramePicture=   "frm_MCA.frx":045E
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1320
      Top             =   5640
   End
   Begin VB.Label Label1 
      Caption         =   "Using"
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   5580
      Width           =   555
   End
   Begin VB.Label Label3 
      Caption         =   "mS"
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   5940
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Integration time interval"
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
      TabIndex        =   4
      Top             =   5940
      Width           =   1935
   End
End
Attribute VB_Name = "frm_MCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comportNo As Integer
Dim Comm_buffer As String
Dim My_buffer As String
Dim iData() As Single
Dim myFld() As String
Dim i As Integer
Dim myCount As Integer
Dim MaxCount As Integer
Dim MaxCH As Integer

Private Declare Function SetCommState Lib "kernel32" (ByVal hCommDev As Long, lpDCB As DCB) As Long
Private Declare Function GetCommState Lib "kernel32" (ByVal nCid As Long, lpDCB As DCB) As Long

Private Type DCB ' Win32API.TXT is incorrect here.
DCBlength As Long
BaudRate As Long
Bits1 As Long
wReserved As Integer
XonLim As Integer
XoffLim As Integer
ByteSize As Byte
Parity As Byte
StopBits As Byte
XonChar As Byte
XoffChar As Byte
ErrorChar As Byte
EofChar As Byte
EvtChar As Byte
wReserved2 As Integer
End Type

Private Sub cmd_Connect_Click()
On Error GoTo errdet:
If cmd_Connect.Caption = "Connect" Then
mCom.CommPort = Val(Combo1)
mCom.PortOpen = True
cmd_Connect.Caption = "Disconnect"
Else
mCom.PortOpen = False
cmd_Connect.Caption = "Connect"
End If
Exit Sub
errdet:
MsgBox Err.Description, vbCritical, "Cannot connect using " & Combo1
End Sub

Private Sub cmd_Export_Click()
On Error GoTo errdet:
dialog.ShowSave
Open dialog.FileName For Output As #1
Print #1, "Statistical table : "
Print #1, "Max count =" + Str(MaxCount) + " at CH =" + Str(MaxCH)
Print #1, "---------------------"
   For i = 0 To 511
    Print #1, Str(i) + "," + Str(iData(i))
    Next
Close #1
Exit Sub
errdet:
End Sub

Private Sub cmd_integrate_Click()
If cmd_integrate.Caption = "Integrate" Then
Timer1.Interval = txt_interval
Timer1.Enabled = True
cmd_integrate.Caption = "Stop"
Else
Timer1.Enabled = False
cmd_integrate.Caption = "Integrate"
End If
End Sub

Private Sub Form_Load()
Combo1.Clear
COMDesc = PortDescription
COMPorts = Ports
For comportNo = 0 To UBound(COMPorts) - 1
        Combo1.AddItem COMPorts(comportNo) + " " + COMDesc(comportNo)
Next
SetSpeed 115200
dgraph.XLabel = "Channel"
dgraph.YLabel = "Counts"
dgraph.XGridNumber = 20
dgraph.YGridNumber = 10
dgraph.SetRange 0, 511, 0, 256
End Sub

Private Sub mCom_OnComm()
On Error GoTo errdet:
Select Case mCom.CommEvent


       Case comEvReceive
                  
          My_buffer = My_buffer + mCom.Input
          If InStr(1, My_buffer, vbCr) > 0 Then
           Comm_buffer = My_buffer
           'Debug.Print My_buffer
           'frm_log.Text1.Text = frm_log.Text1.Text + Comm_buffer
            ' If InStr(1, My_buffer, "exit") > 0 Then mCom.Output = "xxxxxxxxxx"
            
            
            
            'If InStr(1, My_buffer, "..") > 0 Then dgraph.Caption = My_buffer
           If InStr(1, My_buffer, "%%") > 0 Then
                  myCount = -1
                  MaxCH = 0
                  MaxCount = 0
               For i = 4 To (Len(My_buffer) - 2) Step 2
                    myCount = myCount + 1
                     ReDim Preserve iData(myCount)
                  iData(myCount) = Asc(Right(Left(My_buffer, i), 1))
                  iData(myCount) = iData(myCount) + (Asc(Right(Left(My_buffer, i + 1), 1)) * 255)
               Next
               
               For i = 1 To myCount - 1
              If iData(i) > MaxCount Then
                  MaxCount = iData(i)
                  MaxCH = i
                  End If
            Next i
            
             dgraph.ClearGraph
            dgraph.AddElement
            dgraph.ElementLineColor = vbRed
            dgraph.ElementWidth = 1
            If Check1.Value = vbChecked Then
            dgraph.ElementPointSymbol = Dots
            dgraph.ElementPointColor = vbYellow
            dgraph.ElementSolidPoint = True
            End If
            dgraph.Caption = "Scan : Max count = " + Str(MaxCount) + " at Channel = " + Str(MaxCH)
            dgraph.XLabel = "Channel"
            dgraph.YLabel = "Counts"
               
               For i = 0 To myCount
                  
                  'dgraph.PlotXY i, 0, 1
                  dgraph.PlotXY i, iData(i), 1
                  If Check2.Value = vbChecked Then dgraph.PlotXY i, 0, 1
                  dgraph.AutoRange
                  'myGraph.PlotXY i, iData(i), 1
                  'myGraph.AutoRange
                  
            Next
               
               
           End If
            
            
            If InStr(1, My_buffer, "#") > 0 Then
            myFld = Split(My_buffer, ",")
              myCount = -1
              MaxCH = 0
              MaxCount = 0
             If UBound(myFld) > 8 Then
            For i = 1 To UBound(myFld) - 1
                  myCount = myCount + 1
                  ReDim Preserve iData(myCount)
                  iData(myCount) = Val(myFld(i))
             
            Next
            End If
            
            For i = 0 To UBound(iData)
              If iData(i) > MaxCount Then
                  MaxCount = iData(i)
                  MaxCH = i
                  End If
            Next i
            
            dgraph.ClearGraph
            dgraph.AddElement
            dgraph.ElementLineColor = vbRed
            dgraph.ElementWidth = 1
            If Check1.Value = vbChecked Then
            dgraph.ElementPointSymbol = Dots
            dgraph.ElementPointColor = vbYellow
            dgraph.ElementSolidPoint = True
            End If
            dgraph.Caption = "Scan : Max count = " + Str(MaxCount) + " at Channel = " + Str(MaxCH)
            dgraph.XLabel = "Channel"
            dgraph.YLabel = "Counts"
            
            'myGraph.ClearGraph
            'myGraph.AddElement
            'myGraph.ElementLineColor = vbRed
            'myGraph.Caption = "Scan from " + txt_begin + " to " + txt_end + " mV"
            
         
            For i = 0 To myCount
                  
                  'dgraph.PlotXY i, 0, 1
                  dgraph.PlotXY i, iData(i), 1
                  If Check2.Value = vbChecked Then dgraph.PlotXY i, 0, 1
                  dgraph.AutoRange
                  'myGraph.PlotXY i, iData(i), 1
                  'myGraph.AutoRange
                  
            Next
            
        
            End If
            
          ' If Comm_buffer <> Old_buffer Then Old_buffer = Comm_buffer
             ' Debug.Print Comm_Buffer;
               My_buffer = ""
         'Comm_buffer = mCom.Input
               
         'Current_Value = Asc(My_buffer)

              End If
              
       Case comEvSend:  'here put your condition that you want
       Case comEvCTS
       Case comEvDSR
       Case comEvCD
       Case comEvRing
       Case comEvEOF
       Case comBreak
       Case comCDTO
       Case comCTSTO
       Case comDCB
       Case comDSRTO
       Case comFrame
       Case comOverrun
       Case comRxOver
       Case comRxParity
       Case comTxFull
End Select

Exit Sub
errdet:

End Sub

Private Sub Timer1_Timer()
On Error GoTo errdet:
If mCom.PortOpen = True Then
    If Check3.Value = vbChecked Then
            mCom.Output = "O"
            Else
            mCom.Output = "I"
        End If
End If
errdet:
Exit Sub
End Sub

Private Sub SetSpeed(ByVal Speed As Long)
Dim MSCommDCB As DCB
Dim Port As Integer
Port = GetCommState(mCom.CommID, MSCommDCB)
MSCommDCB.BaudRate = Speed
Port = SetCommState(mCom.CommID, MSCommDCB)
End Sub

Private Sub SetLength(ByVal Length As Byte)
Dim MSCommDCB As DCB
Dim Port As Integer
Port = GetCommState(mCom.CommID, MSCommDCB)
MSCommDCB.ByteSize = Length
Port = SetCommState(mCom.CommID, MSCommDCB)
End Sub


