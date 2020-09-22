VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmDialup 
   BackColor       =   &H00C06934&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Call Contact"
   ClientHeight    =   1065
   ClientLeft      =   3705
   ClientTop       =   3315
   ClientWidth     =   2985
   Icon            =   "frmDialup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   Begin MSCommLib.MSComm MSComm1 
      Left            =   840
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton DialButton 
      Caption         =   "Dial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   120
      TabIndex        =   2
      Top             =   645
      Width           =   852
   End
   Begin VB.CommandButton QuitButton 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   2040
      TabIndex        =   1
      Top             =   645
      Width           =   852
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   1080
      TabIndex        =   0
      Top             =   645
      Width           =   852
   End
   Begin VB.Label Status 
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "To dial a number, click the Dial button"
      ForeColor       =   &H00D3C098&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmDialup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

DefInt A-Z

' This flag is set when the user chooses Cancel.
Dim CancelFlag

Private Sub CancelButton_Click()
    ' CancelFlag tells the Dial procedure to exit.
    CancelFlag = True
    CancelButton.Enabled = False
End Sub

Private Sub Dial(Number$)
    Dim DialString$, FromModem$, dummy

    ' AT is the Hayes compatible ATTENTION command and is required to send commands to the modem.
    ' DT means "Dial Tone." The Dial command uses touch tones, as opposed to pulse (DP = Dial Pulse).
    ' Numbers$ is the phone number being dialed.
    ' A semicolon tells the modem to return to command mode after dialing (important).
    ' A carriage return, vbCr, is required when sending commands to the modem.
    DialString$ = "ATDT" + Number$ + ";" + vbCr

    ' Communications port settings.
    ' Assuming that a mouse is attached to COM1, CommPort is set to 2
    MSComm1.CommPort = 2
    MSComm1.Settings = "9600,N,8,1"
    
    ' Open the communications port.
    On Error Resume Next
    MSComm1.PortOpen = True
    If Err Then
       MsgBox "COM2: not available. Change the CommPort property to another port."
       Exit Sub
    End If
    
    ' Flush the input buffer.
    MSComm1.InBufferCount = 0
    
    ' Dial the number.
    MSComm1.Output = DialString$
    
    ' Wait for "OK" to come back from the modem.
    Do
       dummy = DoEvents()
       ' If there is data in the buffer, then read it.
       If MSComm1.InBufferCount Then
          FromModem$ = FromModem$ + MSComm1.Input
          ' Check for "OK".
          If InStr(FromModem$, "OK") Then
             ' Notify the user to pick up the phone.
             Beep
             MsgBox "Please pick up the phone and either press Enter or click OK"
             Exit Do
          End If
       End If
        
       ' Did the user choose Cancel?
       If CancelFlag Then
          CancelFlag = False
          Exit Do
       End If
    Loop
    
    ' Disconnect the modem.
    MSComm1.Output = "ATH" + vbCr
    
    ' Close the port.
    MSComm1.PortOpen = False
End Sub

Private Sub DialButton_Click()
    Dim Number$, Temp$
    
    DialButton.Enabled = False
    QuitButton.Enabled = False
    CancelButton.Enabled = True
    
    ' Get the number to dial.
    Number$ = InputBox$("Enter phone number:", Number$)
        If Number$ = "" Then Exit Sub
    Temp$ = Status
    Status = "Dialing - " + Number$
    
    ' Dial the selected phone number.
    Dial Number$

    DialButton.Enabled = True
    QuitButton.Enabled = True
    CancelButton.Enabled = False

    Status = Temp$
End Sub

Private Sub Form_Load()
    ' Setting InputLen to 0 tells MSComm to read the entire
    ' contents of the input buffer when the Input property
    ' is used.
    MSComm1.InputLen = 0
    
End Sub

Private Sub QuitButton_Click()
    
    Me.Hide
    
End Sub


