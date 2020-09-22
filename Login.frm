VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H00C06934&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1560
   ClientLeft      =   4290
   ClientTop       =   3405
   ClientWidth     =   3360
   ControlBox      =   0   'False
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "ICDBase.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Login"
      Top             =   1680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "New"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000080&
      ForeColor       =   &H00D3C098&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000080&
      ForeColor       =   &H00D3C098&
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label MemID 
      DataField       =   "MemID"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Pass 
      DataField       =   "Pass"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C06934&
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C06934&
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ontop As New clsOnTop
Private Sub Command1_Click()
    If Login.HScroll1.Value = 2 Then
        End
    End If
    If Login.Text1.Text = "" Then
        Login.HScroll1.Value = Login.HScroll1.Value + 1
        Exit Sub
    End If
    If Login.Text1.Text = "" Then
        Login.HScroll1.Value = Login.HScroll1.Value + 1
        Exit Sub
    End If
    Login.Data1.Recordset.FindFirst "memID = '" & Login.Text1.Text & "'"
    If Login.Pass.Caption = Login.Text2.Text Then
        ontop.MakeNormal hWnd
        MsgBox "Login Successful!"
        Login.MemID.Caption = ""
        Login.Pass.Caption = ""
        Login.Text1.Text = ""
        Login.Text2.Text = ""
        frmICDBaseForm.Show
        Me.Hide
        Exit Sub
        
    End If
    ontop.MakeNormal hWnd
    MsgBox "Login Unsuccessful!"
    Login.Text1.Text = ""
    Login.Text2.Text = ""
    Login.HScroll1.Value = Login.HScroll1.Value + 1
End Sub

Private Sub Command2_Click()
    ontop.MakeNormal hWnd
    If Login.Text1.Text = "" Then
        Exit Sub
    End If
    If Login.Text1.Text = "" Then
        Exit Sub
    End If
    Login.Data1.Recordset.AddNew
    Login.Data1.Recordset.Fields("memID") = "" & Login.Text1.Text & ""
    Login.Data1.Recordset.Fields("pass") = "" & Login.Text2.Text & ""
    Login.Data1.Recordset.Update
    Login.MemID.Caption = ""
    Login.Pass.Caption = ""
    Login.Text1.Text = ""
    Login.Text2.Text = ""
    MsgBox "Name and Password Have Been Logged", vbInformation
End Sub

Private Sub Command3_Click()
    End
End Sub


Private Sub Form_Load()

    Data1.DatabaseName = App.Path & "\ICDBase.mdb"
    ontop.MakeTopMost hWnd

End Sub

