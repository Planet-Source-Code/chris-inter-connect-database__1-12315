VERSION 5.00
Begin VB.Form frmReminder 
   BackColor       =   &H00C06934&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Task Reminder"
   ClientHeight    =   3405
   ClientLeft      =   240
   ClientTop       =   1935
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
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
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtMinutes 
      BackColor       =   &H00000080&
      ForeColor       =   &H00D3C098&
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdCurrent 
      Caption         =   "&Current TaskTime"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   1980
      Width           =   2055
   End
   Begin VB.OptionButton optAM 
      BackColor       =   &H00C06934&
      Caption         =   "AM"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   2520
      Width           =   735
   End
   Begin VB.OptionButton optPM 
      BackColor       =   &H00C06934&
      Caption         =   "PM"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "&Set Task Alarm"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   2360
      Width           =   2055
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   2730
      Width           =   2055
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
      TabIndex        =   9
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label lblHour 
      BackColor       =   &H00C06934&
      Caption         =   "Hour"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblMinute 
      BackColor       =   &H00C06934&
      Caption         =   "Minutes "
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1440
      TabIndex        =   7
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
      Picture         =   "frmReminder.frx":0000
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   5220
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   0
      Picture         =   "frmReminder.frx":2EF2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5145
   End
   Begin VB.Image Image3 
      Height          =   3375
      Left            =   4920
      Picture         =   "frmReminder.frx":6654
      Stretch         =   -1  'True
      Top             =   0
      Width           =   300
   End
   Begin VB.Image Image4 
      Height          =   3375
      Left            =   0
      Picture         =   "frmReminder.frx":94FE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   225
   End
End
Attribute VB_Name = "frmReminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCurrent_Click()

    MsgBox SetAlarm, , "Current Alarm Time"

End Sub

Private Sub cmdQuit_Click()

    Me.Hide
   
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

Private Sub Timer1_Timer()
    
    lblTime.Caption = Time   ' Update time display.
    If SetAlarm = lblTime.Caption Then
        'show message box
        
       MsgBox "It is time to complete your task", vbInformation, "Task Alarm"
       
    End If
    
End Sub
