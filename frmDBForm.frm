VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmICDBaseForm 
   BackColor       =   &H00C06934&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inter Connect Database"
   ClientHeight    =   8115
   ClientLeft      =   -75
   ClientTop       =   3150
   ClientWidth     =   12060
   Icon            =   "frmDBForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   12060
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   55
      Top             =   7785
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   582
      Style           =   1
      SimpleText      =   "Inter Connect Database Created By Christopher Palladino"
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
      Enabled         =   0   'False
   End
   Begin VB.Frame fraColorControl 
      BackColor       =   &H00C06934&
      Caption         =   "Color Controls"
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   2520
      TabIndex        =   62
      Top             =   6840
      Width           =   9245
      Begin VB.CommandButton cmdGrid 
         Caption         =   "&Grid Background"
         Height          =   375
         Left            =   7710
         TabIndex        =   37
         Top             =   610
         Width           =   1455
      End
      Begin VB.CommandButton cmdDateTimeColor 
         Caption         =   "Date\Time Color"
         Height          =   375
         Left            =   7710
         TabIndex        =   32
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdFontColor 
         Caption         =   "&Text Field Font Color"
         Height          =   375
         Left            =   5780
         TabIndex        =   31
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdFrameText 
         Caption         =   "&Frame Forecolor"
         Height          =   375
         Left            =   5780
         TabIndex        =   36
         Top             =   610
         Width           =   1935
      End
      Begin VB.CommandButton cmdInputTop 
         Caption         =   "&Label Background Color"
         Height          =   375
         Left            =   90
         TabIndex        =   33
         Top             =   610
         Width           =   1935
      End
      Begin VB.CommandButton cmdLabelColor 
         Caption         =   "&Frame Background Color"
         Height          =   375
         Left            =   3840
         TabIndex        =   35
         Top             =   610
         Width           =   1935
      End
      Begin VB.CommandButton cmdLabelText 
         Caption         =   "&Label Forecolor"
         Height          =   375
         Left            =   2020
         TabIndex        =   34
         Top             =   610
         Width           =   1815
      End
      Begin VB.CommandButton cmdTextBoxes 
         Caption         =   "&Text Field Color"
         Height          =   375
         Left            =   3840
         TabIndex        =   30
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdForeground 
         Caption         =   "&Main Forecolor"
         Height          =   375
         Left            =   2020
         TabIndex        =   29
         ToolTipText     =   "Change the color of the main programs forecolor."
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdBackground 
         Caption         =   "&Main Background Color"
         Height          =   375
         Left            =   90
         TabIndex        =   28
         ToolTipText     =   "Change the background color of the main program."
         Top             =   240
         Width           =   1935
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "frmDBForm.frx":08CA
      Height          =   1215
      Left            =   2520
      TabIndex        =   61
      Top             =   5640
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   2143
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   128
      ForeColor       =   12609844
      BackColorFixed  =   13877400
      ForeColorFixed  =   255
      BackColorSel    =   13877400
      ForeColorSel    =   12609844
      BackColorBkg    =   12609844
      GridColor       =   13877400
      GridColorFixed  =   13877400
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "ICDBase.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Main"
      Top             =   3480
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Frame fraControlCenter 
      BackColor       =   &H00C06934&
      Caption         =   "Control Center"
      ForeColor       =   &H00000080&
      Height          =   1620
      Left            =   2520
      TabIndex        =   57
      Top             =   3945
      Width           =   3975
      Begin VB.CommandButton cmdHelp 
         Caption         =   "&Help Control"
         Height          =   405
         Left            =   120
         TabIndex        =   63
         ToolTipText     =   "Display the help file."
         Top             =   670
         Width           =   3735
      End
      Begin VB.CommandButton cmdScheduleControl 
         Caption         =   "&Task Schedule Control"
         Height          =   405
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Access the scheduling page."
         Top             =   270
         Width           =   3735
      End
      Begin VB.CommandButton cmdAbout 
         Caption         =   "&About Control"
         Height          =   405
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "About this program and author."
         Top             =   1080
         Width           =   3735
      End
   End
   Begin VB.TextBox txtEmergencyNumber 
      BackColor       =   &H00000080&
      DataField       =   "EmergencyNumber"
      DataSource      =   "Data1"
      ForeColor       =   &H00D3C098&
      Height          =   285
      Left            =   9240
      TabIndex        =   16
      ToolTipText     =   "Enter persons emergency (family member or close friend) contact number."
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox txtEmergencyNumberTop 
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
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   53
      TabStop         =   0   'False
      Text            =   "Emergency Number"
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox txtWorkNumber 
      BackColor       =   &H00000080&
      DataField       =   "WorkNumber"
      DataSource      =   "Data1"
      ForeColor       =   &H00D3C098&
      Height          =   285
      Left            =   9240
      TabIndex        =   14
      ToolTipText     =   "Enter persons work phone number."
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txtWorkNumberTop 
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
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   52
      TabStop         =   0   'False
      Text            =   "Work Number"
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox txtClanRank 
      BackColor       =   &H00000080&
      DataField       =   "ClanRank"
      DataSource      =   "Data1"
      ForeColor       =   &H00D3C098&
      Height          =   285
      Left            =   9240
      TabIndex        =   26
      ToolTipText     =   "Enter persons clan (founder, co-founder, normal member or military ranking given by clan leader) rank."
      Top             =   5280
      Width           =   2535
   End
   Begin VB.TextBox txtClanRankTop 
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
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Text            =   "Clan Rank"
      Top             =   5040
      Width           =   2535
   End
   Begin VB.TextBox txtClanName 
      BackColor       =   &H00000080&
      DataField       =   "ClanName"
      DataSource      =   "Data1"
      ForeColor       =   &H00D3C098&
      Height          =   285
      Left            =   6600
      TabIndex        =   25
      ToolTipText     =   "Enter persons clan (name given to a group of players who play as a team) name."
      Top             =   5280
      Width           =   2535
   End
   Begin VB.TextBox txtClantop 
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
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      Text            =   "Clan Name"
      Top             =   5040
      Width           =   2535
   End
   Begin VB.TextBox txtGamingHandle 
      BackColor       =   &H00000080&
      DataField       =   "GamingHandle"
      DataSource      =   "Data1"
      ForeColor       =   &H00D3C098&
      Height          =   285
      Left            =   9240
      TabIndex        =   24
      ToolTipText     =   "Enter persons gaming (name used during gameplay) handle."
      Top             =   4680
      Width           =   2535
   End
   Begin VB.TextBox txtGamingHandleTop 
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
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Text            =   "Gaming Handle"
      Top             =   4440
      Width           =   2535
   End
   Begin VB.TextBox txtFavoriteGame 
      BackColor       =   &H00000080&
      DataField       =   "FavoriteGame"
      DataSource      =   "Data1"
      ForeColor       =   &H00D3C098&
      Height          =   285
      Left            =   6600
      TabIndex        =   23
      ToolTipText     =   "Enter persons favorite (normally played) game."
      Top             =   4680
      Width           =   2535
   End
   Begin VB.TextBox txtFavoritegameTop 
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
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      Text            =   "Favorite Game"
      Top             =   4440
      Width           =   2535
   End
   Begin VB.TextBox txtHomepageTitle 
      BackColor       =   &H00000080&
      DataField       =   "HomepageTitle"
      DataSource      =   "Data1"
      ForeColor       =   &H00D3C098&
      Height          =   285
      Left            =   6600
      TabIndex        =   21
      ToolTipText     =   "Enter persons homepage title."
      Top             =   4080
      Width           =   2535
   End
   Begin VB.TextBox txtHomepageTitleTop 
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
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Text            =   "Homepage Title"
      Top             =   3840
      Width           =   2535
   End
   Begin VB.TextBox txtIRCHandle 
      BackColor       =   &H00000080&
      DataField       =   "IRCHandle"
      DataSource      =   "Data1"
      ForeColor       =   &H00D3C098&
      Height          =   285
      Left            =   9240
      TabIndex        =   20
      ToolTipText     =   "Enter persons IRC handle."
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox txtIRCTop 
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
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Text            =   "IRC Handle"
      Top             =   3240
      Width           =   2535
   End
   Begin VB.TextBox txtAIMHandle 
      BackColor       =   &H00000080&
      DataField       =   "AIMHandle"
      DataSource      =   "Data1"
      ForeColor       =   &H00D3C098&
      Height          =   285
      Left            =   6600
      TabIndex        =   19
      ToolTipText     =   "Enter persons AIM handle."
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox txtAIMHandletop 
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
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Text            =   "AIM Handle"
      Top             =   3240
      Width           =   2535
   End
   Begin VB.TextBox txtICQNumber 
      BackColor       =   &H00000080&
      DataField       =   "ICQNumber"
      DataSource      =   "Data1"
      ForeColor       =   &H00D3C098&
      Height          =   285
      Left            =   9240
      TabIndex        =   18
      ToolTipText     =   "Enter persons ICQ number."
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox txtICQNumberTop 
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
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Text            =   "ICQ Number"
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox txtBeeper 
      BackColor       =   &H00000080&
      DataField       =   "BeeperNumber"
      DataSource      =   "Data1"
      ForeColor       =   &H00D3C098&
      Height          =   285
      Left            =   6600
      TabIndex        =   17
      ToolTipText     =   "Enter persons beeper number."
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox txtBeeperTop 
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
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Text            =   "Beeper Number"
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox txtCellular 
      BackColor       =   &H00000080&
      DataField       =   "CellularNumber"
      DataSource      =   "Data1"
      ForeColor       =   &H00D3C098&
      Height          =   285
      Left            =   6600
      TabIndex        =   15
      ToolTipText     =   "Enter persons cellular number."
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox txtCellPhoneTop 
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
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Text            =   "Cellular Number"
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox txtPhoneNumber 
      BackColor       =   &H00000080&
      DataField       =   "PhoneNumber"
      DataSource      =   "Data1"
      ForeColor       =   &H00D3C098&
      Height          =   285
      Left            =   6600
      TabIndex        =   13
      ToolTipText     =   "Enter persons phone number."
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txtPhoneNumberTop 
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
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Text            =   "Phone Number"
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox txtEmailAddress 
      BackColor       =   &H00000080&
      DataField       =   "EMailAddress"
      DataSource      =   "Data1"
      ForeColor       =   &H00D3C098&
      Height          =   285
      Left            =   9240
      TabIndex        =   12
      ToolTipText     =   "Enter persons e-mail address."
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox txtEmailTop 
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
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Text            =   "E-Mail Address"
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtHomepageAddress 
      BackColor       =   &H00000080&
      DataField       =   "HomepageURL"
      DataSource      =   "Data1"
      ForeColor       =   &H00D3C098&
      Height          =   285
      Left            =   9240
      TabIndex        =   22
      ToolTipText     =   "Enter persons homepage URL."
      Top             =   4080
      Width           =   2535
   End
   Begin VB.TextBox txtHomepageAddressTop 
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
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Text            =   "Homepage URL"
      Top             =   3840
      Width           =   2535
   End
   Begin VB.TextBox txtHomeAddress 
      BackColor       =   &H00000080&
      DataField       =   "HomeAddress"
      DataSource      =   "Data1"
      ForeColor       =   &H00D3C098&
      Height          =   285
      Left            =   6600
      TabIndex        =   11
      ToolTipText     =   "Enter persons home address."
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox txtHomeAddressTop 
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
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Text            =   "Home Address"
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtLastName 
      BackColor       =   &H00000080&
      DataField       =   "LastName"
      DataSource      =   "Data1"
      ForeColor       =   &H00D3C098&
      Height          =   285
      Left            =   9240
      TabIndex        =   10
      ToolTipText     =   "Enter persons last name."
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox txtLastNameTop 
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
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Text            =   "Last Name"
      Top             =   240
      Width           =   2535
   End
   Begin VB.TextBox txtFirstName 
      BackColor       =   &H00000080&
      DataField       =   "FirstName"
      DataSource      =   "Data1"
      ForeColor       =   &H00D3C098&
      Height          =   285
      Left            =   6600
      TabIndex        =   9
      ToolTipText     =   "Enter persons first name."
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox txtFirstNameTop 
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
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   54
      TabStop         =   0   'False
      Text            =   "First Name"
      Top             =   240
      Width           =   2535
   End
   Begin VB.Frame fraDatabaseControls 
      BackColor       =   &H00C06934&
      Caption         =   "Database Controls"
      ForeColor       =   &H00000080&
      Height          =   3210
      Left            =   2520
      TabIndex        =   56
      Top             =   160
      Width           =   3975
      Begin VB.CommandButton cmdClose 
         Caption         =   "E&xit Program"
         Height          =   405
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Exit this program, be sure to save all work before leaving."
         Top             =   2670
         Width           =   3735
      End
      Begin VB.CommandButton cmdPrevious 
         Cancel          =   -1  'True
         Caption         =   "&Previous Record"
         Height          =   405
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Goto previous database record."
         Top             =   1050
         Width           =   3735
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next Record"
         Height          =   405
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Goto next database record."
         Top             =   640
         Width           =   3735
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update Record"
         Height          =   405
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Update database record you just edited."
         Top             =   1860
         Width           =   3735
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh Record"
         Height          =   405
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Refresh current database record showing."
         Top             =   1460
         UseMaskColor    =   -1  'True
         Width           =   3735
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete Record"
         Height          =   405
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Delete current database record."
         Top             =   2270
         Width           =   3735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add Record"
         Height          =   405
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Add a new record to the database."
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame fraDateTime 
      BackColor       =   &H00C06934&
      Caption         =   "Current Date And Time"
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   120
      TabIndex        =   58
      Top             =   0
      Width           =   2310
      Begin VB.Label lblTime 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1175
         TabIndex        =   60
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         BackColor       =   &H00C06934&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   -35
         TabIndex        =   59
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10440
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1920
      Top             =   0
   End
   Begin VB.Image imgBanner 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   7450
      Left            =   120
      Picture         =   "frmDBForm.frx":08DE
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2310
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add Record"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuNext 
         Caption         =   "&Next Record"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuPrevious 
         Caption         =   "&Previous Record"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh Record"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "&Update Record"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete Record"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit Program"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuDashBottom0 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuTask 
         Caption         =   "&Task Schedule Control"
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu mnuAlarm 
         Caption         =   "&Task Alarm"
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu mnuDashBottom1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuMainBack 
         Caption         =   "&Main Background Color"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuLabelBack 
         Caption         =   "&Label Background Color"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuTextfield 
         Caption         =   "&Text Field Color"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuFrameBack 
         Caption         =   "&Frame Background Color"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuGrid 
         Caption         =   "&Grid Background Color"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnudash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMainfore 
         Caption         =   "&Main Forecolor"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuLabelfore 
         Caption         =   "&Label Forecolor"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuTextFont 
         Caption         =   "&Text Field Font Color"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuFrameFore 
         Caption         =   "&Frame Forecolor"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuDash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDate 
         Caption         =   "&Date\Time Color"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuDash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmail 
         Caption         =   "&E-Mail Contact"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuVisit 
         Caption         =   "&Visit Contacts Homepage"
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuDash6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHome 
         Caption         =   "&Call Contacts Phone Number"
      End
      Begin VB.Menu mnuWork 
         Caption         =   "&Call Contacts Work Number"
      End
      Begin VB.Menu mnuCell 
         Caption         =   "&Call Contacts Cellular Number"
      End
      Begin VB.Menu mnuEmergency 
         Caption         =   "&Call Contacts Emergency Number"
      End
      Begin VB.Menu mnuBeep 
         Caption         =   "&Call Contacts Beeper Number"
      End
      Begin VB.Menu mnuDashBottom2 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpFile 
         Caption         =   "&Help File"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuDashBottom3 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "frmICDBaseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function IsValidEmailAddress(AddressString As String) As Boolean

Dim sHost As String
Dim iPos As Integer
Dim sInvalidChars As String

If Len(Trim(AddressString)) = 0 Then
    IsValidEmailAddress = False
    Exit Function
End If

sInvalidChars = "!#$%^&*()=+{}[]|\;:'/?>,< "

    For iPos = 1 To Len(AddressString)
        If InStr(sInvalidChars, _
            Mid(AddressString, iPos, 1)) > 0 Then
            
            IsValidEmailAddress = False
            Exit Function
        End If
    Next


iPos = InStr(AddressString, "@")

If iPos = 0 Or Left(AddressString, 1) = "@" Then
    IsValidEmailAddress = False
    Exit Function
End If

sHost = Mid(AddressString, iPos + 1)
'can't have multiple "@" chars in the string
If InStr(sHost, "@") > 0 Then
    IsValidEmailAddress = False
    Exit Function
End If

IsValidEmailAddress = IsValidIPHost(UCase(sHost))


End Function
Private Function IsValidIPHost(HostString As String) As Boolean

Dim sHost As String
Dim bDottedQuad As Boolean
Dim sSplit() As String
Dim ictr As Integer
Dim bAns As Boolean
Dim sTopLevelDomains() As String

sHost = HostString

If InStr(sHost, ".") = 0 Then
    IsValidIPHost = False
    Exit Function
End If

sSplit = Split(sHost, ".")

If UBound(sSplit) = 3 Then
    bDottedQuad = True
    For ictr = 0 To 3
        If Not IsNumeric(sSplit(ictr)) Then
            bDottedQuad = False
            Exit For
        End If
    Next
    
    If bDottedQuad Then
        bAns = True
        For ictr = 0 To 3
            If ictr = 0 Then
            bAns = Val(sSplit(ictr)) <= 239
                If bAns = False Then Exit For
            Else
                bAns = Val(sSplit(ictr)) <= 255
                If bAns = False Then Exit For
            End If
        Next
        
        IsValidIPHost = bAns

        
        Exit Function
    End If
End If 'ubound(ssplit) = 3

    IsValidIPHost = isTopLevelDomain(sSplit(UBound(sSplit)))


End Function

Private Function isTopLevelDomain(DomainString As String) As Boolean
Dim asTopLevels() As String
Dim ictr As Integer
Dim iNumDomains As Integer
Dim bAns As Boolean

iNumDomains = 251
ReDim asTopLevels(iNumDomains - 1) As String


'Obtained from www.IANA.com.  Can and will change
asTopLevels(0) = "COM"
asTopLevels(1) = "ORG"
asTopLevels(2) = "NET"
asTopLevels(3) = "EDU"
asTopLevels(4) = "GOV"
asTopLevels(5) = "MIL"
asTopLevels(6) = "INT"
asTopLevels(7) = "AF"
asTopLevels(8) = "AL"
asTopLevels(9) = "DZ"
asTopLevels(10) = "AS"
asTopLevels(11) = "AD"
asTopLevels(12) = "AO"
asTopLevels(13) = "AI"
asTopLevels(14) = "AQ"
asTopLevels(15) = "AG"
asTopLevels(16) = "AR"
asTopLevels(17) = "AM"
asTopLevels(18) = "AW"
asTopLevels(19) = "AC"
asTopLevels(20) = "AU"
asTopLevels(21) = "AT"
asTopLevels(22) = "AZ"
asTopLevels(23) = "BS"
asTopLevels(24) = "BH"
asTopLevels(25) = "BD"
asTopLevels(26) = "BB"
asTopLevels(27) = "BY"
asTopLevels(28) = "BZ"
asTopLevels(29) = "BT"
asTopLevels(30) = "BJ"
asTopLevels(31) = "BE"
asTopLevels(32) = "BM"
asTopLevels(33) = "BO"
asTopLevels(34) = "BA"
asTopLevels(35) = "BW"
asTopLevels(36) = "BV"
asTopLevels(37) = "BR"
asTopLevels(38) = "IO"
asTopLevels(39) = "BN"
asTopLevels(40) = "BG"
asTopLevels(41) = "BF"
asTopLevels(42) = "BI"
asTopLevels(43) = "KH"
asTopLevels(44) = "CM"
asTopLevels(45) = "CA"
asTopLevels(46) = "CV"
asTopLevels(47) = "KY"
asTopLevels(48) = "CF"
asTopLevels(49) = "TD"
asTopLevels(50) = "CL"
asTopLevels(51) = "CN"
asTopLevels(52) = "CX"
asTopLevels(53) = "CC"
asTopLevels(54) = "CO"
asTopLevels(55) = "KM"
asTopLevels(56) = "CD"
asTopLevels(57) = "CG"
asTopLevels(58) = "CK"
asTopLevels(59) = "CR"
asTopLevels(60) = "CI"
asTopLevels(61) = "HR"
asTopLevels(62) = "CU"
asTopLevels(63) = "CY"
asTopLevels(64) = "CZ"
asTopLevels(65) = "DK"
asTopLevels(66) = "DJ"
asTopLevels(67) = "DM"
asTopLevels(68) = "DO"
asTopLevels(69) = "TP"
asTopLevels(70) = "EC"
asTopLevels(71) = "EG"
asTopLevels(72) = "SV"
asTopLevels(73) = "GQ"
asTopLevels(74) = "ER"
asTopLevels(75) = "EE"
asTopLevels(76) = "ET"
asTopLevels(77) = "FK"
asTopLevels(78) = "FO"
asTopLevels(79) = "FJ"
asTopLevels(80) = "FI"
asTopLevels(81) = "FR"
asTopLevels(82) = "GF"
asTopLevels(83) = "PF"
asTopLevels(84) = "TF"
asTopLevels(85) = "GA"
asTopLevels(86) = "GM"
asTopLevels(87) = "GE"
asTopLevels(88) = "DE"
asTopLevels(89) = "GH"
asTopLevels(90) = "GI"
asTopLevels(91) = "GR"
asTopLevels(92) = "GL"
asTopLevels(93) = "GD"
asTopLevels(94) = "GP"
asTopLevels(95) = "GU"
asTopLevels(96) = "GT"
asTopLevels(97) = "GG"
asTopLevels(98) = "GN"
asTopLevels(99) = "GW"
asTopLevels(100) = "GY"
asTopLevels(101) = "HT"
asTopLevels(102) = "HM"
asTopLevels(103) = "VA"
asTopLevels(104) = "HN"
asTopLevels(105) = "HK"
asTopLevels(106) = "HU"
asTopLevels(107) = "IS"
asTopLevels(108) = "IN"
asTopLevels(109) = "ID"
asTopLevels(110) = "IR"
asTopLevels(111) = "IQ"
asTopLevels(112) = "IE"
asTopLevels(113) = "IM"
asTopLevels(114) = "IL"
asTopLevels(115) = "IT"
asTopLevels(116) = "JM"
asTopLevels(117) = "JP"
asTopLevels(118) = "JE"
asTopLevels(119) = "JO"
asTopLevels(120) = "KZ"
asTopLevels(121) = "KE"
asTopLevels(122) = "KI"
asTopLevels(123) = "KP"
asTopLevels(124) = "KR"
asTopLevels(125) = "KW"
asTopLevels(126) = "KG"
asTopLevels(127) = "LA"
asTopLevels(128) = "LV"
asTopLevels(129) = "LB"
asTopLevels(130) = "LS"
asTopLevels(131) = "LR"
asTopLevels(132) = "LY"
asTopLevels(133) = "LI"
asTopLevels(134) = "LT"
asTopLevels(135) = "LU"
asTopLevels(136) = "MO"
asTopLevels(137) = "MK"
asTopLevels(138) = "MG"
asTopLevels(139) = "MW"
asTopLevels(140) = "MY"
asTopLevels(141) = "MV"
asTopLevels(142) = "ML"
asTopLevels(143) = "MT"
asTopLevels(144) = "MH"
asTopLevels(145) = "MQ"
asTopLevels(146) = "MR"
asTopLevels(147) = "MU"
asTopLevels(148) = "YT"
asTopLevels(149) = "MX"
asTopLevels(150) = "FM"
asTopLevels(151) = "MD"
asTopLevels(152) = "MC"
asTopLevels(153) = "MN"
asTopLevels(154) = "MS"
asTopLevels(155) = "MA"
asTopLevels(156) = "MZ"
asTopLevels(157) = "MM"
asTopLevels(158) = "NA"
asTopLevels(159) = "NR"
asTopLevels(160) = "NP"
asTopLevels(161) = "NL"
asTopLevels(162) = "AN"
asTopLevels(163) = "NC"
asTopLevels(164) = "NZ"
asTopLevels(165) = "NI"
asTopLevels(166) = "NE"
asTopLevels(167) = "NG"
asTopLevels(168) = "NU"
asTopLevels(169) = "NF"
asTopLevels(170) = "MP"
asTopLevels(171) = "NO"
asTopLevels(172) = "OM"
asTopLevels(173) = "PK"
asTopLevels(174) = "PW"
asTopLevels(175) = "PA"
asTopLevels(176) = "PG"
asTopLevels(177) = "PY"
asTopLevels(178) = "PE"
asTopLevels(179) = "PH"
asTopLevels(180) = "PN"
asTopLevels(181) = "PL"
asTopLevels(182) = "PT"
asTopLevels(183) = "PR"
asTopLevels(184) = "QA"
asTopLevels(185) = "RE"
asTopLevels(186) = "RO"
asTopLevels(187) = "RU"
asTopLevels(188) = "RW"
asTopLevels(189) = "KN"
asTopLevels(190) = "LC"
asTopLevels(191) = "VC"
asTopLevels(192) = "WS"
asTopLevels(193) = "SM"
asTopLevels(194) = "ST"
asTopLevels(195) = "SA"
asTopLevels(196) = "SN"
asTopLevels(197) = "SC"
asTopLevels(198) = "SL"
asTopLevels(199) = "SG"
asTopLevels(200) = "SK"
asTopLevels(201) = "SI"
asTopLevels(202) = "SB"
asTopLevels(203) = "SO"
asTopLevels(204) = "ZA"
asTopLevels(205) = "GS"
asTopLevels(206) = "ES"
asTopLevels(207) = "LK"
asTopLevels(208) = "SH"
asTopLevels(209) = "PM"
asTopLevels(210) = "SD"
asTopLevels(211) = "SR"
asTopLevels(212) = "SJ"
asTopLevels(213) = "SZ"
asTopLevels(214) = "SE"
asTopLevels(215) = "CH"
asTopLevels(216) = "SY"
asTopLevels(217) = "TW"
asTopLevels(218) = "TJ"
asTopLevels(219) = "TZ"
asTopLevels(220) = "TH"
asTopLevels(221) = "TG"
asTopLevels(222) = "TK"
asTopLevels(223) = "TO"
asTopLevels(224) = "TT"
asTopLevels(225) = "TN"
asTopLevels(226) = "TR"
asTopLevels(227) = "TM"
asTopLevels(228) = "TC"
asTopLevels(229) = "TV"
asTopLevels(230) = "UG"
asTopLevels(231) = "UA"
asTopLevels(232) = "AE"
asTopLevels(233) = "GB"
asTopLevels(234) = "US"
asTopLevels(235) = "UM"
asTopLevels(236) = "UY"
asTopLevels(237) = "UZ"
asTopLevels(238) = "VU"
asTopLevels(239) = "VE"
asTopLevels(240) = "VN"
asTopLevels(241) = "VG"
asTopLevels(242) = "VI"
asTopLevels(243) = "WF"
asTopLevels(244) = "EH"
asTopLevels(245) = "YE"
asTopLevels(246) = "YU"
asTopLevels(247) = "ZR"
asTopLevels(248) = "ZM"
asTopLevels(249) = "ZW"
asTopLevels(250) = "UK"

For ictr = 0 To iNumDomains - 1
    If asTopLevels(ictr) = DomainString Then
        bAns = True
        Exit For
    End If
Next

isTopLevelDomain = bAns

End Function

Private Sub cmdAbout_Click()
    
    frmAbout.Show 'show the about form
    frmSchedule.Visible = False 'if the schedule form is showing hide it
    Me.Hide 'hide this form
    
End Sub

Private Sub cmdAdd_Click()
    
    prompt$ = "Enter the new record."
    reply = MsgBox(prompt$, vbOKCancel, "Add Record")
    If reply = vbOK Then 'if the user clicks ok
        Data1.Refresh 'Refresh the database
        txtFirstName.SetFocus 'move cursor to First Name box
        Data1.Recordset.AddNew 'and get new record
    
    End If
    
End Sub

Private Sub cmdBackground_Click()

On Error GoTo ErrHandler 'Cancel error in property pages
    CommonDialog1.ShowColor 'Display the color dialog box
    frmICDBaseForm.BackColor = CommonDialog1.Color 'set the forms background color to the selected color
  Exit Sub
  
ErrHandler:
    MsgBox "Cancel button has been pressed!", vbExclamation
    
End Sub

Private Sub cmdClose_Click()

  On Error GoTo ErrHandler
  Dim Answer As Integer
  Answer = MsgBox("Are you sure you wish to exit? Remember to save all work before exiting.", vbOKCancel + vbQuestion, "Exit Window") 'make a message box pop up before the form exits displaying a message
    
    If Answer = vbOK Then 'if the user clicks ok then
    
    End 'shut the program
     
  End If
  
  Exit Sub
  
ErrHandler:
   MsgBox Err.Description
   
End Sub

Private Sub cmdDateTimeColor_Click()

On Error GoTo ErrHandler 'Cancel error in property pages
    CommonDialog1.ShowColor 'Display the color dialog box
    lblDate(0).ForeColor = CommonDialog1.Color 'set the Date And color to the selected color
    lblTime.ForeColor = CommonDialog1.Color    'set the Time And color to the selected color
  Exit Sub
  
ErrHandler:
    MsgBox "Cancel button has been pressed!", vbExclamation
    
End Sub

Private Sub cmdDelete_Click()

    On Error GoTo ErrHandler
    MsgBox "Delete current record from database?", vbOKCancel + vbQuestion, "Delete Record"
    If vbOK Then
    Data1.Recordset.Delete 'Delete current record from the database
  
  End If
      
    Data1.Recordset.MoveNext 'move tot he next record
    If Data1.Recordset.EOF Then 'f theres no record then
        Data1.Recordset.MoveLast 'move to the last one
  End If
  
    Exit Sub
    
ErrHandler:
 MsgBox Err.Description
   
   cmdAdd.SetFocus 'Set the focus on the add button
   
End Sub

Private Sub cmdFontColor_Click()

On Error GoTo ErrHandler 'Cancel error in property pages
    CommonDialog1.ShowColor 'Display the color dialog box
    txtFirstName.ForeColor = CommonDialog1.Color 'set the input fields text color to the selected color
    txtLastName.ForeColor = CommonDialog1.Color
    txtHomeAddress.ForeColor = CommonDialog1.Color
    txtEmailAddress.ForeColor = CommonDialog1.Color
    txtPhoneNumber.ForeColor = CommonDialog1.Color
    txtWorkNumber.ForeColor = CommonDialog1.Color
    txtCellular.ForeColor = CommonDialog1.Color
    txtEmergencyNumber.ForeColor = CommonDialog1.Color
    txtBeeper.ForeColor = CommonDialog1.Color
    txtICQNumber.ForeColor = CommonDialog1.Color
    txtAIMHandle.ForeColor = CommonDialog1.Color
    txtIRCHandle.ForeColor = CommonDialog1.Color
    txtHomepageTitle.ForeColor = CommonDialog1.Color
    txtHomepageAddress.ForeColor = CommonDialog1.Color
    txtFavoriteGame.ForeColor = CommonDialog1.Color
    txtGamingHandle.ForeColor = CommonDialog1.Color
    txtClanName.ForeColor = CommonDialog1.Color
    txtClanRank.ForeColor = CommonDialog1.Color
  Exit Sub
  
ErrHandler:
    MsgBox "Cancel button has been pressed!", vbExclamation
    
End Sub

Private Sub cmdForeground_Click()

On Error GoTo ErrHandler 'Cancel error in property pages
    CommonDialog1.ShowColor 'Display the color dialog box
    frmICDBaseForm.ForeColor = CommonDialog1.Color 'set the forms fore color to the selected color
  Exit Sub
  
ErrHandler:
    MsgBox "Cancel button has been pressed!", vbExclamation
    
End Sub

Private Sub cmdFrameText_Click()

On Error GoTo ErrHandler 'Cancel error in property pages
    CommonDialog1.ShowColor 'Display the color dialog box
    fraColorControl.ForeColor = CommonDialog1.Color 'set the frames text color to the selected color
    fraDatabaseControls.ForeColor = CommonDialog1.Color
    fraControlCenter.ForeColor = CommonDialog1.Color
    fraDateTime.ForeColor = CommonDialog1.Color
  Exit Sub
  
ErrHandler:
    MsgBox "Cancel button has been pressed!", vbExclamation
    
End Sub

Private Sub cmdGrid_Click()

On Error GoTo ErrHandler 'Cancel error in property pages
    CommonDialog1.ShowColor 'Display the color dialog box
    MSFlexGrid1.BackColorBkg = CommonDialog1.Color 'set the forms background color to the selected color
  Exit Sub
  
ErrHandler:
    MsgBox "Cancel button has been pressed!", vbExclamation
    
End Sub

Private Sub cmdHelp_Click()

    CommonDialog1.CancelError = True 'Set cancel to true
    On Error GoTo ErrHandler
    CommonDialog1.HelpCommand = cdlHelpForceFile 'Set the help command property
    CommonDialog1.HelpFile = (App.Path & "\Inter Connect Help\Inter Connect Database.HLP") 'specify the help file
    CommonDialog1.ShowHelp 'Display the help engine
    Exit Sub
    
ErrHandler:
    'user pressed the cancel button
    Exit Sub
    
End Sub

Private Sub cmdInputTop_Click()

On Error GoTo ErrHandler 'Cancel error in property pages
    CommonDialog1.ShowColor 'Display the color dialog box
    txtFirstNameTop.BackColor = CommonDialog1.Color 'set the label background color to the selected color
    txtLastNameTop.BackColor = CommonDialog1.Color
    txtHomeAddressTop.BackColor = CommonDialog1.Color
    txtEmailTop.BackColor = CommonDialog1.Color
    txtPhoneNumberTop.BackColor = CommonDialog1.Color
    txtWorkNumberTop.BackColor = CommonDialog1.Color
    txtCellPhoneTop.BackColor = CommonDialog1.Color
    txtEmergencyNumberTop.BackColor = CommonDialog1.Color
    txtBeeperTop.BackColor = CommonDialog1.Color
    txtICQNumberTop.BackColor = CommonDialog1.Color
    txtAIMHandletop.BackColor = CommonDialog1.Color
    txtIRCTop.BackColor = CommonDialog1.Color
    txtHomepageTitleTop.BackColor = CommonDialog1.Color
    txtHomepageAddressTop.BackColor = CommonDialog1.Color
    txtFavoritegameTop.BackColor = CommonDialog1.Color
    txtGamingHandleTop.BackColor = CommonDialog1.Color
    txtClantop.BackColor = CommonDialog1.Color
    txtClanRankTop.BackColor = CommonDialog1.Color
  Exit Sub
  
ErrHandler:
    MsgBox "Cancel button has been pressed!", vbExclamation
    
End Sub

Private Sub cmdLabelColor_Click()

On Error GoTo ErrHandler 'Cancel error in property pages
    CommonDialog1.ShowColor 'Display the color dialog box
    fraColorControl.BackColor = CommonDialog1.Color 'set the frames background color to the selected color
    fraDatabaseControls.BackColor = CommonDialog1.Color
    fraControlCenter.BackColor = CommonDialog1.Color
    fraDateTime.BackColor = CommonDialog1.Color
  Exit Sub
  
ErrHandler:
    MsgBox "Cancel button has been pressed!", vbExclamation
    
End Sub

Private Sub cmdLabelText_Click()

On Error GoTo ErrHandler 'Cancel error in property pages
    CommonDialog1.ShowColor 'Display the color dialog box
    txtFirstNameTop.ForeColor = CommonDialog1.Color 'set the Label text color to the selected color
    txtLastNameTop.ForeColor = CommonDialog1.Color
    txtHomeAddressTop.ForeColor = CommonDialog1.Color
    txtEmailTop.ForeColor = CommonDialog1.Color
    txtPhoneNumberTop.ForeColor = CommonDialog1.Color
    txtWorkNumberTop.ForeColor = CommonDialog1.Color
    txtCellPhoneTop.ForeColor = CommonDialog1.Color
    txtEmergencyNumberTop.ForeColor = CommonDialog1.Color
    txtBeeperTop.ForeColor = CommonDialog1.Color
    txtICQNumberTop.ForeColor = CommonDialog1.Color
    txtAIMHandletop.ForeColor = CommonDialog1.Color
    txtIRCTop.ForeColor = CommonDialog1.Color
    txtHomepageTitleTop.ForeColor = CommonDialog1.Color
    txtHomepageAddressTop.ForeColor = CommonDialog1.Color
    txtFavoritegameTop.ForeColor = CommonDialog1.Color
    txtGamingHandleTop.ForeColor = CommonDialog1.Color
    txtClantop.ForeColor = CommonDialog1.Color
    txtClanRankTop.ForeColor = CommonDialog1.Color
  Exit Sub
  
ErrHandler:
    MsgBox "Cancel button has been pressed!", vbExclamation
    
End Sub


Private Sub cmdNext_Click()

    On Error GoTo ErrHandler
    
    If Not Data1.Recordset.EOF Then 'check if theres a record after the current one
       Data1.Recordset.MoveNext 'if there is then goto it
    End If
    
    If Data1.Recordset.EOF And Data1.Recordset.RecordCount > 0 Then 'if theres no record after it
        Data1.Recordset.MoveLast 'move back to the last record
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description
    
    cmdAdd.SetFocus 'Set the focus on the add button
    
End Sub

Private Sub cmdPrevious_Click()

    On Error GoTo ErrHandler
    
    If Not Data1.Recordset.BOF Then 'check to see if there is a record before the current one
       Data1.Recordset.MovePrevious 'if there is then goto it
    End If
    
    If Data1.Recordset.BOF And Data1.Recordset.RecordCount > 0 Then 'if theres no record before it then
        Data1.Recordset.MoveFirst 'move to the first record
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description
    
    cmdAdd.SetFocus 'Set the focus on the add button
    
End Sub

Private Sub cmdRefresh_Click()
    On Error GoTo ErrHandler
    Data1.Refresh 'Refresh screen
    
    Exit Sub
    
ErrHandler:
   MsgBox Err.Description
   
   
    cmdAdd.SetFocus 'Set the focus on the add button
End Sub

Private Sub cmdScheduleControl_Click()

    frmSchedule.Show 'show the schedule form
    frmAbout.Visible = False 'if the about form is showing then hide it
    Me.Hide 'hide this form
       
End Sub

Private Sub cmdTextBoxes_Click()

On Error GoTo ErrHandler 'Cancel error in property pages
    CommonDialog1.ShowColor 'Display the color dialog box
    txtFirstName.BackColor = CommonDialog1.Color 'set the input fields background color to the selected color
    txtLastName.BackColor = CommonDialog1.Color
    txtHomeAddress.BackColor = CommonDialog1.Color
    txtEmailAddress.BackColor = CommonDialog1.Color
    txtPhoneNumber.BackColor = CommonDialog1.Color
    txtWorkNumber.BackColor = CommonDialog1.Color
    txtCellular.BackColor = CommonDialog1.Color
    txtEmergencyNumber.BackColor = CommonDialog1.Color
    txtBeeper.BackColor = CommonDialog1.Color
    txtICQNumber.BackColor = CommonDialog1.Color
    txtAIMHandle.BackColor = CommonDialog1.Color
    txtIRCHandle.BackColor = CommonDialog1.Color
    txtHomepageTitle.BackColor = CommonDialog1.Color
    txtHomepageAddress.BackColor = CommonDialog1.Color
    txtFavoriteGame.BackColor = CommonDialog1.Color
    txtGamingHandle.BackColor = CommonDialog1.Color
    txtClanName.BackColor = CommonDialog1.Color
    txtClanRank.BackColor = CommonDialog1.Color
  Exit Sub
  
ErrHandler:
    MsgBox "Cancel button has been pressed!", vbExclamation
    
End Sub

Private Sub cmdUpdate_Click()
    
    On Error GoTo ErrHandler
    
    Data1.UpdateRecord 'Update current record
    
    Exit Sub
    
ErrHandler:
   MsgBox Err.Description
   
   
   cmdAdd.SetFocus 'Set the focus on the add button
   
End Sub

Private Sub Form_Load()

    Data1.DatabaseName = App.Path & "\ICDBase.mdb"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    End
    
End Sub

Private Sub mnuAbout_Click()
    
    cmdAbout_Click 'acts as if you pressed the about button
    
End Sub

Private Sub mnuAdd_Click()
    cmdAdd_Click 'acts as if you pressed the add button
End Sub

Private Sub mnuAlarm_Click()

    frmReminder.Show
    
End Sub

Private Sub mnuBeep_Click()

    txtBeeper_DblClick

End Sub

Private Sub mnuCell_Click()

    txtCellular_DblClick

End Sub

Private Sub mnuDelete_Click()
    
    cmdDelete_Click 'acts as if you pressed the delete button
    
End Sub

Private Sub mnuEmail_Click()

txtEmailAddress_DblClick 'acts as if you double clicked on the the e-mail text box

End Sub

Private Sub mnuEmergency_Click()

   txtEmergencyNumber_DblClick

End Sub

Private Sub mnuExit_Click()
    
    cmdClose_Click 'acts as if you pressed the exit program button
    
End Sub

Private Sub mnuFrameBack_Click()
    
    cmdLabelColor_Click 'acts as if you pressed the label color button
    
End Sub

Private Sub mnuFrameFore_Click()
    
    cmdFrameText_Click 'acts as if you pressed the frame text button
    
End Sub

Private Sub mnuGrid_Click()
    
    cmdGrid_Click 'acts as if you pressed the grid button
    
End Sub

Private Sub mnuHelpFile_Click()
    
    cmdHelp_Click 'acts as if you pressed the help button
    
End Sub

Private Sub mnuHome_Click()

    txtPhoneNumber_DblClick

End Sub

Private Sub mnuLabelBack_Click()
    
    cmdInputTop_Click 'acts as if you pressed the input top button
    
End Sub

Private Sub mnuLabelfore_Click()
    
    cmdLabelText_Click 'acts as if you pressed the label text button
    
End Sub

Private Sub mnuMainBack_Click()
    
    cmdBackground_Click 'acts as if you pressed the background button
    
End Sub

Private Sub mnuMainfore_Click()
    
    cmdForeground_Click 'acts as if you pressed the foreground button
    
End Sub

Private Sub mnuNext_Click()
    
    cmdNext_Click 'acts as if you pressed the next button
    
End Sub

Private Sub mnuPrevious_Click()
    
    cmdPrevious_Click 'acts as if you pressed the previous button
    
End Sub

Private Sub mnuRefresh_Click()
    
    cmdRefresh_Click 'acts as if you pressed the refresh button
    
End Sub

Private Sub mnuTask_Click()
    
    cmdScheduleControl_Click 'acts as if you pressed the schedule control button
    
End Sub

Private Sub mnuTextfield_Click()
    
    cmdTextBoxes_Click 'acts as if you pressed the text boxes button
    
End Sub

Private Sub mnuTextFont_Click()
    
    cmdFontColor_Click 'acts as if you pressed the font color button
    
End Sub

Private Sub mnuUpdate_Click()
    
    cmdUpdate_Click 'acts as if you pressed the update button
    
End Sub

Private Sub mnuVisit_Click()

    txtHomepageAddress_DblClick 'acts as if you double clicked on the Homepage address textbox area

End Sub

Private Sub mnuWork_Click()

    txtWorkNumber_DblClick

End Sub

Private Sub Timer1_Timer()
    
    lblTime.Caption = Format$(Now, "h:mm:ss AM/PM") ' displays the time in the labels caption
    lblDate(0).Caption = Format$(Now, "m/dd/yyyy") 'displays the date in the labels caption
    
End Sub


Private Sub txtBeeper_DblClick()

    If txtBeeper.Text = "" Then
    MsgBox "You must enter the contacts Beeper number in order to call them", vbInformation, "Call Error"
    
  Else
  
    frmDialup.Show
    
  End If

End Sub

Private Sub txtCellular_DblClick()

    If txtCellular.Text = "" Then
    MsgBox "You must enter the contacts cellular phone number in order to call them", vbInformation, "Call Error"
    
  Else
  
    frmDialup.Show
    
  End If
  
End Sub

Private Sub txtEmailAddress_DblClick()

   Shell ("Start mailto:" & txtEmailAddress.Text), vbHide 'opens your mail program and places the text in the To: field

End Sub

Private Sub txtEmergencyNumber_DblClick()

    If txtEmergencyNumber.Text = "" Then
    MsgBox "You must enter the contacts emergency phone number in order to call them", vbInformation, "Call Error"
    
  Else
  
    frmDialup.Show
    
  End If


End Sub

Private Sub txtHomepageAddress_DblClick()

    Shell ("Start " & txtHomepageAddress.Text), vbHide 'opens your browser and places the url in the browsers address box

End Sub

Private Sub txtHomepageTitle_DblClick()

    txtHomepageAddress_DblClick

End Sub

Private Sub txtPhoneNumber_DblClick()

    If txtPhoneNumber.Text = "" Then
    MsgBox "You must enter the contacts home phone number in order to call them", vbInformation, "Call Error"
    
  Else
  
    frmDialup.Show
    
  End If
    

End Sub

Private Sub txtWorkNumber_DblClick()

    If txtWorkNumber.Text = "" Then
    MsgBox "You must enter the contacts work phone number in order to call them", vbInformation, "Call Error"
    
  Else
  
    frmDialup.Show
    
  End If

End Sub
