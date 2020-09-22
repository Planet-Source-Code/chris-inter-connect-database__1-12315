VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSchedule 
   BackColor       =   &H00C06934&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inter Connect Scheduler"
   ClientHeight    =   5745
   ClientLeft      =   225
   ClientTop       =   975
   ClientWidth     =   11640
   ForeColor       =   &H00000080&
   Icon            =   "frmSchedule.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optPM 
      BackColor       =   &H00C06934&
      Caption         =   "PM"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   2480
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   480
   End
   Begin VB.TextBox txtHours 
      BackColor       =   &H00000080&
      ForeColor       =   &H00D3C098&
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtMinutes 
      BackColor       =   &H00000080&
      ForeColor       =   &H00D3C098&
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdCurrent 
      Caption         =   "&Current TaskTime"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "&Set Task Alarm"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   2290
      Width           =   2055
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit Task Scheduler"
      Height          =   315
      Left            =   9840
      TabIndex        =   16
      ToolTipText     =   "Exit and return to the main program section."
      Top             =   5400
      Width           =   1575
   End
   Begin VB.TextBox txtNotes 
      Alignment       =   2  'Center
      BackColor       =   &H00D3C098&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "Comment And Task Field"
      Top             =   3000
      Width           =   6255
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   "ICDBase.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TaskSchedule"
      Top             =   0
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.TextBox txtDateField 
      BackColor       =   &H00000080&
      DataField       =   "Date"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      ForeColor       =   &H00D3C098&
      Height          =   285
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Enter the date of your scheduled task."
      Top             =   5160
      Width           =   4335
   End
   Begin VB.TextBox txtDateTop 
      Alignment       =   2  'Center
      BackColor       =   &H00D3C098&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "Date Of Scheduled Task"
      Top             =   4920
      Width           =   4335
   End
   Begin VB.Frame fraScheduleControls 
      BackColor       =   &H00C06934&
      Caption         =   "Schedule Controls"
      ForeColor       =   &H00000080&
      Height          =   1185
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   4335
      Begin VB.CommandButton cmdNextTask 
         Caption         =   "&Next Task"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1500
         TabIndex        =   8
         ToolTipText     =   "Goto next task."
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdPreviousTask 
         Caption         =   "&Previous Task"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2830
         TabIndex        =   9
         ToolTipText     =   "Goto previous task."
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdRefreshTask 
         Caption         =   "&Refresh Task"
         Enabled         =   0   'False
         Height          =   375
         Left            =   160
         TabIndex        =   10
         ToolTipText     =   "Refresh the database."
         Top             =   620
         Width           =   1335
      End
      Begin VB.CommandButton cmdUpdateTask 
         Caption         =   "&Update Task"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2830
         TabIndex        =   12
         ToolTipText     =   "Update current task."
         Top             =   620
         Width           =   1335
      End
      Begin VB.CommandButton cmdDeleteTask 
         Caption         =   "&Delete Task"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1500
         TabIndex        =   11
         ToolTipText     =   "Delete current task."
         Top             =   620
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddTask 
         Caption         =   "&Add Task"
         Height          =   375
         Left            =   160
         TabIndex        =   7
         ToolTipText     =   "Add new task."
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox txtTaskField 
      BackColor       =   &H00000080&
      DataField       =   "Task"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      ForeColor       =   &H00D3C098&
      Height          =   2100
      Left            =   5280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      ToolTipText     =   "Enter your description or comments within this field (250 Character Max)."
      Top             =   3270
      Width           =   6255
   End
   Begin VB.TextBox txtTimeField 
      BackColor       =   &H00000080&
      DataField       =   "Time"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      ForeColor       =   &H00D3C098&
      Height          =   285
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Enter the time of your scheduled task."
      Top             =   4560
      Width           =   4335
   End
   Begin VB.TextBox txtTimeTop 
      Alignment       =   2  'Center
      BackColor       =   &H00D3C098&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Text            =   "Time Of Scheduled Task"
      Top             =   4320
      Width           =   4335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Bindings        =   "frmSchedule.frx":08CA
      Height          =   2805
      Left            =   5280
      TabIndex        =   6
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4948
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   3
   End
   Begin VB.OptionButton optAM 
      BackColor       =   &H00C06934&
      Caption         =   "AM"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   2480
      Width           =   735
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00C06934&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   360
      TabIndex        =   23
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label lblHour 
      BackColor       =   &H00C06934&
      Caption         =   "Hour"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1440
      TabIndex        =   22
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblMinute 
      BackColor       =   &H00C06934&
      Caption         =   "Minutes "
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1440
      TabIndex        =   21
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      X1              =   360
      X2              =   4800
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Image Image1 
      Height          =   180
      Left            =   0
      Picture         =   "frmSchedule.frx":08DE
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   5220
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   0
      Picture         =   "frmSchedule.frx":37D0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5145
   End
   Begin VB.Image Image3 
      Height          =   2775
      Left            =   4920
      Picture         =   "frmSchedule.frx":6F32
      Stretch         =   -1  'True
      Top             =   0
      Width           =   300
   End
   Begin VB.Image Image4 
      Height          =   2895
      Left            =   0
      Picture         =   "frmSchedule.frx":9DDC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   225
   End
End
Attribute VB_Name = "frmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Hrs
Dim Mnt
Dim AMPM
Dim SetAlarm
Private Sub Calendar1_Click()

End Sub

Private Sub cmdAddTask_Click()
    
    prompt$ = "Enter the new task."
    reply = MsgBox(prompt$, vbOKCancel, "Add Task")
    If reply = vbOK Then 'if the user clicks ok
    cmdPreviousTask.Enabled = True 'Enable the disabled database controls
    cmdNextTask.Enabled = True
    cmdRefreshTask.Enabled = True
    cmdDeleteTask.Enabled = True
    cmdUpdateTask.Enabled = True
    
    txtTimeField.Enabled = True 'Enable the textboxes
    txtDateField.Enabled = True
    txtTaskField.Enabled = True
    
    txtTimeField.SetFocus 'move cursor to time field box
    Data2.Recordset.AddNew
    End If
    
    
End Sub

Private Sub cmdCurrent_Click()
    
    MsgBox SetAlarm, , "Current Alarm Time"
    
End Sub

Private Sub cmdDeleteTask_Click()
    
    On Error GoTo ErrHandler
    MsgBox "Delete current scheduled task from database?", vbOKCancel + vbQuestion, "Delete Task"
    If vbOK Then
    Data2.Recordset.Delete 'Delete current record from the database
  
  End If
      
    Data2.Recordset.MoveNext
    If Data2.Recordset.EOF Then
        Data2.Recordset.MoveLast
  End If
  
    Exit Sub
    
ErrHandler:
 MsgBox Err.Description
   
   cmdAddTask.SetFocus 'Set the focus on the add button
   
End Sub

Private Sub cmdExit_Click()

    frmICDBaseForm.Show
    Me.Hide
    
End Sub

Private Sub cmdNextTask_Click()

     On Error GoTo ErrHandler
    
    If Not Data2.Recordset.EOF Then
       Data2.Recordset.MoveNext
    End If
    
    If Data2.Recordset.EOF And Data1.Recordset.RecordCount > 0 Then
        Data2.Recordset.MoveLast
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description
    
    cmdAddTask.SetFocus 'Set the focus on the add button
    
End Sub

Private Sub cmdPreviousTask_Click()

    On Error GoTo ErrHandler
    
    If Not Data2.Recordset.BOF Then
       Data2.Recordset.MovePrevious
    End If
    
    If Data2.Recordset.BOF And Data1.Recordset.RecordCount > 0 Then
        Data2.Recordset.MoveFirst
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description
    
    cmdAddTask.SetFocus 'Set the focus on the add button
    
End Sub

Private Sub cmdQuit_Click()

End Sub

Private Sub cmdRefreshTask_Click()

    On Error GoTo ErrHandler
    Data2.Refresh 'Refresh screen
    
    Exit Sub
    
ErrHandler:
   MsgBox Err.Description
   
    
    cmdAddTask.SetFocus 'Set the focus on the add button
    
End Sub

Private Sub cmdSet_Click()

    Hrs = txtHours.Text
    Mnt = txtMinutes.Text
    If optAM.Value = True Then
        AMPM = "AM"
    ElseIf optPM.Value = True Then
        AMPM = "PM"
    End If

    SetAlarm = Hrs + ":" + Mnt + ":00 " + AMPM
    
    If txtHours.Text = "" Then
    MsgBox "You didn't specify an hour.", vbInformation, "Hour Setting"
       
     Else
    MsgBox "Task alarm has been set.", vbInformation, "Alarm Confirmation"
    
    End If
    
    If txtMinutes.Text = "" Then
    MsgBox "You didn't specify any minutes.", vbInformation, "Minutes Setting"
       
    End If
    
End Sub

Private Sub cmdUpdateTask_Click()

    On Error GoTo ErrHandler
    Data2.UpdateRecord 'Update current record
    
    Exit Sub
    
ErrHandler:
   MsgBox Err.Description
   
   
   cmdAddTask.SetFocus 'Set the focus on the add button
   
End Sub

Private Sub fraCalendar_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub Form_Load()

    Data2.DatabaseName = App.Path & "\ICDBase.mdb"

End Sub

Private Sub Timer1_Timer()

    lblTime.Caption = Time   ' Update time display.
    If SetAlarm = lblTime.Caption Then
        'show message box
        
       MsgBox "It is time to coplete your task", vbInformation, "Task Alarm"
       
    End If
    
End Sub
