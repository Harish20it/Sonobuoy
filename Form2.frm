VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form2.frx":0000
      Height          =   5775
      Left            =   360
      OleObjectBlob   =   "Form2.frx":0014
      TabIndex        =   3
      Top             =   720
      Width           =   9855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Back To Main"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   6720
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Delete All Records"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   1
      Top             =   6720
      Width           =   2415
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Command1_Click()
While Data1.Recordset.RecordCount <> 0
Data1.Recordset.MoveFirst
Data1.Recordset.Delete
Wend
DBGrid1.Refresh
End Sub

Private Sub Command2_Click()
Form2.Hide
Form1.Show
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\Eswari_data.mdb"
Data1.RecordSource = "hsd_table"
Data1.Refresh
DBGrid1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form2
Form1.Show
End Sub
