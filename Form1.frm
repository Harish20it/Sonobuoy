VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   Caption         =   "Form1"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8010
   DrawWidth       =   2
   FillColor       =   &H80000004&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Height          =   8295
      Left            =   13800
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   8235
      ScaleWidth      =   3795
      TabIndex        =   26
      Top             =   1320
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   480
      Picture         =   "Form1.frx":6C4B
      ScaleHeight     =   1635
      ScaleWidth      =   2955
      TabIndex        =   24
      Top             =   7800
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5280
      TabIndex        =   22
      Top             =   5205
      Width           =   390
   End
   Begin VB.CommandButton Command4 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5280
      TabIndex        =   21
      Top             =   4920
      Width           =   390
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3480
      TabIndex        =   20
      Top             =   4920
      Width           =   1560
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   14
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1140
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1080
      Top             =   1560
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   480
      Top             =   1305
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   8880
      TabIndex        =   11
      Top             =   7380
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View Record"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   7230
      TabIndex        =   10
      Top             =   7380
      Width           =   1635
   End
   Begin VB.OptionButton Option1 
      Caption         =   "&Current"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   2
      Left            =   8730
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Width           =   1500
   End
   Begin VB.OptionButton Option1 
      Caption         =   "&Voltage"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   1
      Left            =   7245
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6120
      Value           =   -1  'True
      Width           =   1500
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   3360
      TabIndex        =   6
      Top             =   3840
      Width           =   1740
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3480
      TabIndex        =   4
      Top             =   2520
      Width           =   1725
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "ANALYSIS OF SEA WATER ACTIVATED BATTERIES WITH                                 DATA LOGGER"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2055
      Left            =   240
      TabIndex        =   25
      Top             =   10200
      Width           =   11895
   End
   Begin VB.Line Line12 
      BorderWidth     =   5
      X1              =   240
      X2              =   12120
      Y1              =   9720
      Y2              =   9720
   End
   Begin VB.Line Line10 
      BorderWidth     =   5
      X1              =   12120
      X2              =   12120
      Y1              =   840
      Y2              =   9720
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   240
      X2              =   12120
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line13 
      X1              =   5880
      X2              =   5880
      Y1              =   840
      Y2              =   960
   End
   Begin VB.Line Line11 
      BorderWidth     =   5
      X1              =   240
      X2              =   240
      Y1              =   840
      Y2              =   9720
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      Height          =   255
      Left            =   1920
      TabIndex        =   23
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Dac Output"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   720
      TabIndex        =   19
      Top             =   4905
      Width           =   2430
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000A&
      Caption         =   "mA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   18
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      TabIndex        =   17
      Top             =   2640
      Width           =   165
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "---->   ms"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10440
      TabIndex        =   16
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6360
      TabIndex        =   15
      Top             =   2520
      Width           =   165
   End
   Begin VB.Line Line9 
      BorderWidth     =   5
      Index           =   0
      X1              =   5880
      X2              =   5880
      Y1              =   840
      Y2              =   9720
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   240
      X2              =   12120
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Graph"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   555
      Left            =   8715
      TabIndex        =   13
      Top             =   1305
      Width           =   1440
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "Readings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   555
      Left            =   2070
      TabIndex        =   12
      Top             =   1245
      Width           =   2145
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   7875
      Shape           =   4  'Rounded Rectangle
      Top             =   1170
      Width           =   3030
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1620
      Shape           =   4  'Rounded Rectangle
      Top             =   1155
      Width           =   3000
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   1695
      TabIndex        =   9
      Top             =   7335
      Width           =   105
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   7080
      X2              =   6960
      Y1              =   2640
      Y2              =   2520
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   6840
      X2              =   6960
      Y1              =   2640
      Y2              =   2520
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   10920
      X2              =   11040
      Y1              =   5880
      Y2              =   5760
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   10920
      X2              =   11040
      Y1              =   5640
      Y2              =   5760
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   6960
      X2              =   11040
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   6960
      X2              =   6960
      Y1              =   2520
      Y2              =   5760
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Current"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1440
      TabIndex        =   5
      Top             =   3960
      Width           =   1680
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Voltage"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1680
      TabIndex        =   3
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10320
      TabIndex        =   2
      Top             =   330
      Width           =   75
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   300
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "High Speed Data Acquiring System"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   630
      Left            =   1680
      TabIndex        =   0
      Top             =   210
      Width           =   8370
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form2.DBGrid1.Refresh
Form1.Hide
Form2.Show
End Sub

Private Sub Command2_Click()
If Command3.Caption = "&Stop" Then
mk_ch = MsgBox("Do You Want to Stop the Record & Exit", vbYesNo, "Corfirmation")
Command3.Caption = "&Stop"
    If mk_ch = vbYes Then
        End
    Else
        Exit Sub
    End If
Else
End
End If
End Sub

Private Sub Command3_Click()
If (Command3.Caption = "&Start") Then
Command1.Enabled = False
Timer2.Enabled = True
Command3.Caption = "&Stop"
Label6.Caption = "Recording is Started ..."
Else
Command1.Enabled = True
Timer2.Enabled = False
Command3.Caption = "&Start"
Label6.Caption = "Recording is Stopped"
End If
End Sub

Private Sub Command4_Click()
DacOut = DacOut + 1
If DacOut > 255 Then DacOut = 255
End Sub

Private Sub Command5_Click()
DacOut = DacOut - 1
If DacOut < 0 Then DacOut = 0
End Sub

Private Sub Form_Load()
sx = 6960
sy = 5760

plot = 1

MSComm1.PortOpen = True

MSComm1.Output = "{1B00}"
Sleep (100)
'FindComPort
Form2.Data1.DatabaseName = App.Path & "\eswari_data.mdb"
Form2.Data1.RecordSource = "hsd_table"
Form2.DBGrid1.Refresh
End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Option1_Click(Index As Integer)
plot = Index
If (Index = 1) Then Label9.Caption = "V"
If (Index = 2) Then Label9.Caption = "I"
grclr
End Sub

Private Sub Picture2_Click()

End Sub

Private Sub Timer1_Timer()
Me.Caption = Format(Date, "dd/mm/yyyy") & "                                               High Speed Data Acquiring System                                           " & Format(Time, "hh:mm:ss")
Label2.Caption = Format(Date, "dd/mm/yyyy")
Label3.Caption = Format(Time, "hh:mm:ss")
MSComm1.Output = "{40}"
Sleep (100)
temp = MSComm1.Input
If (temp <> "") Then
volt = ChrToVal(temp)
volt = Round(volt / 18)
Else
volt = 0
End If
Text1.Text = volt

MSComm1.Output = "{41}"
Sleep (100)
           temp = MSComm1.Input
If (temp <> "") Then
curr = ChrToVal(temp)
curr = curr / 2
Else
curr = 0
End If
Text2.Text = Format(curr, "0#")

ex = sx + 50
Select Case plot
Case 1
ey = 5760 - ((volt / 25) * 3240)
Case 2
ey = 5760 - ((curr / 100) * 3240)
End Select
Line (sx, sy)-(ex, ey), vbBlue
sx = ex
sy = ey
If sx > 11040 Then
grclr
End If

If Len(Hex(DacOut)) = 1 Then
    MSComm1.Output = "{5B0" & CStr(Hex(DacOut)) & "}"
    Sleep (100)
Else
    MSComm1.Output = "{5B" & CStr(Hex(DacOut)) & "}"
    Sleep (100)
End If

Text3.Text = DacOut

End Sub



Sub grclr()
Line (6960, 2520)-(11040, 5760), Form1.BackColor, BF
sx = 6960
sy = 5760
Line3.Refresh
Line4.Refresh
Line8.Refresh
Line5.Refresh
End Sub

Private Sub Timer2_Timer()

Form2.Data1.Recordset.AddNew
Form2.Data1.Recordset.Fields(0) = Format(Date, "dd:mm:yyyy")
Form2.Data1.Recordset.Fields(1) = Format(Time, "hh:mm:ss")
Form2.Data1.Recordset.Fields(2) = Text1.Text
Form2.Data1.Recordset.Fields(3) = Text2.Text
Form2.Data1.Recordset.Update
Form2.Data1.Refresh
Form2.DBGrid1.Refresh

End Sub
