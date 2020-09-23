VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "PrivaShare ver. 1.3"
   ClientHeight    =   5325
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7350
   Icon            =   "PrivaShare.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   1111
      ButtonWidth     =   1376
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Connect"
            Key             =   "connect"
            ImageKey        =   "connect"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "search"
            ImageKey        =   "search"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Share"
            Key             =   "share"
            ImageKey        =   "share"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "No Share"
            Key             =   "noShare"
            ImageKey        =   "noShare"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Welcome"
            Key             =   "welcome"
            ImageKey        =   "welcome"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      TabIndex        =   9
      Top             =   4944
      Width           =   7344
      _ExtentX        =   12965
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   4921
            MinWidth        =   4762
            Text            =   "PrivaShare (c) 2001 Gene Hamilton"
            TextSave        =   "PrivaShare (c) 2001 Gene Hamilton"
            Object.ToolTipText     =   "That's me!"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4234
            MinWidth        =   4234
            Object.ToolTipText     =   "Number of current connections open."
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   1
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "7/29/02"
            Object.ToolTipText     =   "Todays date."
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Log:"
            TextSave        =   "Log:"
            Object.ToolTipText     =   "Logging Enabled/Disabled."
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4212
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   7368
      _ExtentX        =   12991
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   12632256
      ForeColor       =   12582912
      TabCaption(0)   =   "Chat"
      TabPicture(0)   =   "PrivaShare.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "MMControl1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtSend"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdSend"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdDrop"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "tvwConnects"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtOutput"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdAddFavorites"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdSendSound"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdRelay"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdPassive"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Download"
      TabPicture(1)   =   "PrivaShare.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command1"
      Tab(1).Control(1)=   "cmdRequestFile"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(3)=   "Frame1"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Share/Upload"
      TabPicture(2)   =   "PrivaShare.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(1)=   "Frame5"
      Tab(2).Control(2)=   "cmdUpload"
      Tab(2).Control(3)=   "Frame4"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Settings/Log"
      TabPicture(3)   =   "PrivaShare.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdSecure"
      Tab(3).Control(1)=   "Frame8"
      Tab(3).Control(2)=   "Frame7"
      Tab(3).Control(3)=   "txtNetIP"
      Tab(3).Control(4)=   "txtLocalIP"
      Tab(3).Control(5)=   "Frame2"
      Tab(3).Control(6)=   "txtName"
      Tab(3).Control(7)=   "Label4"
      Tab(3).Control(8)=   "Label2"
      Tab(3).Control(9)=   "Label6"
      Tab(3).ControlCount=   10
      TabCaption(4)   =   "Win Sounds"
      TabPicture(4)   =   "PrivaShare.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label8"
      Tab(4).Control(1)=   "Label9"
      Tab(4).Control(2)=   "Label10"
      Tab(4).Control(3)=   "File4"
      Tab(4).Control(4)=   "cmdSound"
      Tab(4).ControlCount=   5
      Begin VB.CommandButton cmdPassive 
         Caption         =   "Go Passive!"
         Enabled         =   0   'False
         Height          =   252
         Left            =   5640
         TabIndex        =   60
         ToolTipText     =   "No messages will be printed in chat window so you can play a game on this computer and relay your messages though another."
         Top             =   3000
         Width           =   1572
      End
      Begin VB.CommandButton cmdRelay 
         Caption         =   "Relay selected"
         Enabled         =   0   'False
         Height          =   252
         Left            =   4440
         TabIndex        =   59
         Top             =   3000
         Width           =   1212
      End
      Begin VB.CommandButton cmdSound 
         Caption         =   "Play sound for selected connection"
         Height          =   372
         Left            =   -71520
         TabIndex        =   57
         Top             =   1920
         Width           =   2772
      End
      Begin VB.FileListBox File4 
         Height          =   3015
         Left            =   -74760
         TabIndex        =   55
         Top             =   840
         Width           =   2772
      End
      Begin VB.CommandButton cmdSecure 
         BackColor       =   &H00CCB7B9&
         Caption         =   "Security"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -69240
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdSendSound 
         Caption         =   "Send .wav Recording to Selected"
         Height          =   255
         Left            =   4560
         TabIndex        =   51
         Top             =   3840
         Width           =   2655
      End
      Begin VB.CommandButton cmdAddFavorites 
         Caption         =   "&Add  Selected's IP"
         Enabled         =   0   'False
         Height          =   252
         Left            =   5640
         TabIndex        =   49
         ToolTipText     =   "You can add selected connection above to you favorites list."
         Top             =   2760
         Width           =   1572
      End
      Begin VB.Frame Frame8 
         Caption         =   "Preferences"
         Height          =   1215
         Left            =   -70680
         TabIndex        =   46
         Top             =   2760
         Width           =   2775
         Begin VB.CommandButton cmdLoadPreferences 
            Caption         =   "Load Preferences"
            Height          =   255
            Left            =   360
            TabIndex        =   48
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save Current Preferences"
            Height          =   255
            Left            =   360
            TabIndex        =   47
            Top             =   720
            Width           =   2055
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Logging options"
         Height          =   1575
         Left            =   -70680
         TabIndex        =   42
         Top             =   1080
         Width           =   2775
         Begin VB.CommandButton Command6 
            Caption         =   "Start Logging"
            Height          =   255
            Left            =   360
            TabIndex        =   45
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Stop Logging"
            Height          =   255
            Left            =   360
            TabIndex        =   44
            Top             =   720
            Width           =   2055
         End
         Begin VB.CommandButton cmdSaveLog 
            Caption         =   "Save && Clear Log"
            Height          =   255
            Left            =   360
            TabIndex        =   43
            Top             =   1200
            Width           =   2055
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "File Uploading"
         Height          =   975
         Left            =   -70080
         TabIndex        =   40
         Top             =   1800
         Width           =   1935
         Begin VB.CheckBox ChkUpload 
            Caption         =   "Permit uploads to this directory?"
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "File Sharing"
         Height          =   1095
         Left            =   -70080
         TabIndex        =   37
         Top             =   480
         Width           =   1935
         Begin VB.OptionButton optShare 
            Caption         =   "Share files"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optShare 
            Caption         =   "No file sharing"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   38
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.TextBox txtNetIP 
         Height          =   285
         Left            =   -71160
         TabIndex        =   35
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Show Hosts Files"
         Height          =   255
         Left            =   -70080
         TabIndex        =   34
         ToolTipText     =   "Look into connections file sharing directory."
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton cmdRequestFile 
         Caption         =   "Request Selected File"
         Height          =   255
         Left            =   -70080
         TabIndex        =   33
         ToolTipText     =   "Select file above, and press this button to download."
         Top             =   3600
         Width           =   2175
      End
      Begin VB.Frame Frame3 
         Caption         =   "Host download folder"
         Height          =   2535
         Left            =   -70200
         TabIndex        =   31
         Top             =   960
         Width           =   2415
         Begin VB.ListBox lstFiles 
            Height          =   2010
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "Upload Selected"
         Height          =   495
         Left            =   -70080
         TabIndex        =   30
         ToolTipText     =   "If connection allows uploading, you can send them a file."
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         Caption         =   "Folder to share Files from"
         Height          =   3375
         Left            =   -74880
         TabIndex        =   23
         Top             =   480
         Width           =   4575
         Begin VB.DriveListBox Drive2 
            Height          =   315
            Left            =   120
            TabIndex        =   26
            Top             =   2880
            Width           =   2055
         End
         Begin VB.DirListBox Dir2 
            Height          =   2340
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   2055
         End
         Begin VB.FileListBox File3 
            DragIcon        =   "PrivaShare.frx":04CE
            Height          =   2625
            Left            =   2280
            System          =   -1  'True
            TabIndex        =   24
            Top             =   360
            Width           =   2175
         End
      End
      Begin RichTextLib.RichTextBox txtOutput 
         Height          =   2172
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   4092
         _ExtentX        =   7223
         _ExtentY        =   3836
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"PrivaShare.frx":0910
      End
      Begin VB.TextBox txtLocalIP 
         Height          =   285
         Left            =   -72840
         TabIndex        =   19
         Top             =   720
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Caption         =   "Logging"
         Height          =   2895
         Left            =   -74880
         TabIndex        =   14
         Top             =   1080
         Width           =   3975
         Begin VB.ListBox lstLogging 
            Height          =   1815
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   3735
         End
         Begin VB.CheckBox ChkTransfers 
            Caption         =   "Log Transfers"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   1335
         End
         Begin VB.CheckBox ChkTime 
            Caption         =   "Log Time"
            Height          =   255
            Left            =   2760
            TabIndex        =   16
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox ChkIPs 
            Caption         =   "Log Name/IP"
            Height          =   255
            Left            =   1440
            TabIndex        =   15
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   -74760
         TabIndex        =   13
         Text            =   "Newbie"
         Top             =   720
         Width           =   1695
      End
      Begin MSComctlLib.TreeView tvwConnects 
         Height          =   2172
         Left            =   4440
         TabIndex        =   12
         ToolTipText     =   "These are your connections."
         Top             =   600
         Width           =   2772
         _ExtentX        =   4895
         _ExtentY        =   3836
         _Version        =   393217
         HideSelection   =   0   'False
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
      Begin VB.CommandButton cmdDrop 
         Caption         =   "Drop Selected"
         Enabled         =   0   'False
         Height          =   252
         Left            =   4440
         TabIndex        =   10
         ToolTipText     =   "This will drop connection to who's selected."
         Top             =   2760
         Width           =   1212
      End
      Begin VB.Frame Frame1 
         Caption         =   "Folder to Download Files into"
         Height          =   3375
         Left            =   -74760
         TabIndex        =   3
         Top             =   480
         Width           =   4575
         Begin VB.FileListBox File1 
            DragIcon        =   "PrivaShare.frx":0992
            Height          =   2625
            Left            =   2280
            System          =   -1  'True
            TabIndex        =   6
            Top             =   360
            Width           =   2175
         End
         Begin VB.DirListBox Dir1 
            Height          =   2340
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   2055
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Top             =   2880
            Width           =   2055
         End
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         ToolTipText     =   "Press to send the message you typ in above. Or just hit return."
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox txtSend 
         Height          =   525
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   3120
         Width           =   3732
      End
      Begin MCI.MMControl MMControl1 
         Height          =   375
         Left            =   6360
         TabIndex        =   50
         ToolTipText     =   "Press the circle to record 3 seconds of sound.  Play to hear it."
         Top             =   3360
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   661
         _Version        =   393216
         PlayEnabled     =   -1  'True
         RecordEnabled   =   -1  'True
         PrevVisible     =   0   'False
         NextVisible     =   0   'False
         PauseVisible    =   0   'False
         BackVisible     =   0   'False
         StepVisible     =   0   'False
         StopVisible     =   0   'False
         EjectVisible    =   0   'False
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin VB.Label Label10 
         Caption         =   "These are files in your windows\media folder. If both computers have the file you select it will be played at both ends."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -71520
         TabIndex        =   61
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label9 
         Caption         =   "Select files in window ending with .wav only."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -71520
         TabIndex        =   58
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "Windows multimedia sounds"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   -74760
         TabIndex        =   56
         Top             =   480
         Width           =   2772
      End
      Begin VB.Label Label7 
         Caption         =   "Use your mic input on sound card to send 3 second wav file."
         ForeColor       =   &H00000080&
         Height          =   372
         Left            =   4080
         TabIndex        =   52
         Top             =   3360
         Width           =   2292
      End
      Begin VB.Label Label4 
         Caption         =   "Internet IP"
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
         Left            =   -71160
         TabIndex        =   36
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Local IP"
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
         Left            =   -72840
         TabIndex        =   21
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Your Name"
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
         Left            =   -74760
         TabIndex        =   20
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Enter your text here:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Chat window"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Currently connected to"
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   0
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6480
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":0DD4
            Key             =   "shake"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":1228
            Key             =   "help1"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":167C
            Key             =   "help2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":1AD0
            Key             =   "view"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":1F24
            Key             =   "drop"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":2378
            Key             =   "lock"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":27CC
            Key             =   "secure"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":2C20
            Key             =   "fileClosed"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":3074
            Key             =   "fileOpen"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":34C8
            Key             =   "search"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":3A0C
            Key             =   "dropAll"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":3E60
            Key             =   "connect"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":42B4
            Key             =   "welcome"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":4708
            Key             =   "x"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":481C
            Key             =   "share"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrivaShare.frx":4C70
            Key             =   "noShare"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   6000
      Top             =   1800
   End
   Begin VB.ListBox lstConnect 
      Height          =   450
      Index           =   0
      Left            =   3600
      TabIndex        =   28
      Top             =   2040
      Width           =   975
   End
   Begin VB.FileListBox File2 
      Height          =   480
      Left            =   4680
      TabIndex        =   29
      Top             =   2040
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   5640
      Top             =   1800
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5640
      Top             =   2280
   End
   Begin MCI.MMControl MMControl2 
      Height          =   372
      Left            =   1200
      TabIndex        =   54
      Top             =   2880
      Width           =   2832
      _ExtentX        =   6244
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPort 
         Caption         =   "Port"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSearchNodes 
         Caption         =   "Search for file through your connections."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About PrivaShare"
      End
   End
   Begin VB.Menu mnuNodes 
      Caption         =   "Join"
      Visible         =   0   'False
      Begin VB.Menu JoinNode 
         Caption         =   "Join to this node"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'PrivaShare Application  ver. 1.3
'Written by Gene Hamilton (c) May 2001-2002
'Geno@localaccess.com
'
'For latest version or fixxes:
'http://www.gamerserver.com/privashare
'
'**************************************
'     VB6 .OCX components needed:
'mci32.ocx      194k    windows/system/
'mswinsck.ocx   106k    windows/system/
'**************************************



'
'*** This progy does not use a central server to find others. So the responsability for files shared is with the sharer only.
'Until I know how to get local internet ip, progy will use your network IP for identification. 8(
'
'This was a learning project so the code is repeated and messy in areas.
'
'Thanks to FreeVBCode.com for the many code examples on their site to learn from.
'IF you use any code, please give credit where credit is due.

Private Sub cmdAddFavorites_Click()

On Error GoTo noOneSelected4

Dim i As Integer

For i = 1 To intNum_Connections
        getSendersInfo i
        If tvwConnects.SelectedItem = strName Then
            'Put check routine here to see if already in list.
            '-------------------------------------------------
            If MsgBox("Add " & strName & " to favorite?", vbYesNo) = vbYes Then
                'Add to listboxes
                frmConnect.lstFavName.AddItem strName
                frmConnect.lstFavIP.AddItem strIP
                frmConnect.lstConnections.AddItem strName & vbTab & strIP
                saveFavorites
            End If
        End If
Next i
Exit Sub

noOneSelected4:
'No name was selected in tree view.  Siwwy wabbits.
MsgBox ("You must first select who you want to add.")
    
End Sub

Private Sub cmdDrop_Click()
    Dim i As Integer
    
On Error GoTo noOneSelectedDrop
    
    'look through Nodes.
    For i = 1 To intNum_Connections
        getSendersInfo i
        If tvwConnects.SelectedItem = strName Then
            If MsgBox("Disconnect from " & strName & "?", vbYesNo) = vbYes Then
                tvwConnects.Nodes.Remove strName & strIP
                Winsock1(i).Close '*********************
                'Set Name to nothing so you know later it's not used.
                Connect(i).Name = ""
                
                intNum_ConnectionsNow = intNum_ConnectionsNow - 1 'Decrease number of connections.
                'Connection was dropped.  That wasn't nice...
                
                'Set security to zero.
                intAccess(i) = 0
                
                'If no connections, disable buttons
                If intNum_ConnectionsNow = 0 Then
                    cmdSend.Enabled = False
                    cmdDrop.Enabled = False
                    txtName.Enabled = True
                End If
                
                txtOutput.Text = txtOutput.Text + vbCrLf + strName & " was dropped."
                txtOutput.SelStart = Len(txtOutput.Text)
                txtSend.SetFocus
                
                'Update log, IPs.
                If ChkIPs And blnLog Then
                    lstLogging.AddItem strName & " was dropped."
                End If
                
                'Update log, Date and time.
                If ChkTime And blnLog Then
                    lstLogging.AddItem Time
                End If
 
                Exit For
            End If
        End If
    Next i
    
    'Clear their connections.
    lstConnect(i).Clear
    
    sendConnectionsToAll i 'Update node list of all connections.
    
    Exit Sub
    
noOneSelectedDrop:
'No name was selected in tree view.  Siwwy wabbits.
MsgBox ("You must first select who you want to drop.")

End Sub

Private Sub cmdGetConnections_Click()

On Error GoTo noOneSelected2
    Dim i As Integer

    For i = 1 To intNum_Connections
        getSendersInfo i
        '***************************TEST********Check me later********************
        
        '***************************************************************************
        If tvwConnects.SelectedItem.key = strName & strIP Then
            'send request, and null takes up space, not used.
            sendToOne i, "requestContacts,null"
        End If
    Next i

    Exit Sub

noOneSelected2:
'No name was selected in tree view.  Siwwy wabbits.
MsgBox ("You must first select who you want to get directory from in Connections/Chat window.")

End Sub

Private Sub cmdPassive_Click()
    If Not passive Then
        passive = True
        cmdPassive.Caption = "Silent mode!"
    Else
        passive = False
        cmdPassive.Caption = "Go passive!"
    End If
    
End Sub

Private Sub cmdRelay_Click()

On Error GoTo noOneSelected4

Dim i As Integer

For i = 1 To intNum_Connections
        getSendersInfo i
        If tvwConnects.SelectedItem = strName Then
            'Put check routine here to see if already in list.
            '-------------------------------------------------
            If MsgBox("Relay messages for " & strName & "?", vbYesNo) = vbYes Then
                'Add to listboxes
                Connect(i).relay = True
            End If
        End If
Next i
Exit Sub

noOneSelected4:
'No name was selected in tree view.  Siwwy wabbits.
MsgBox ("You must first select who you want to relay messages for.")

End Sub

Private Sub cmdRequestFile_Click()

Dim i As Integer

On Error GoTo fileError

    For i = 1 To intNum_Connections
        getSendersInfo i
        '***********************TEST********Check me later************
        
        '************************************************************
        If tvwConnects.SelectedItem.key = strName & strIP Then
            'send request, and filename.
            sendToOne i, "requestFile," & lstFiles.Text
            'If logging file downloads, log it.
            If ChkTransfers And blnLog Then
                lstLogging.AddItem "File " & lstFiles.Text & "requested..."
            End If
        End If
    Next i
    
    Exit Sub
    
fileError:
    MsgBox ("There was an error while trying to download file.")
    
End Sub

Private Sub cmdSaveLog_Click()

Dim strLog
Dim strMonth As String
Dim strDay As String
Dim strYear As String
Dim intCounter As Integer
Dim i As Integer
intCounter = 1

'Get current date.
strLog = Date

'Get the day and year for name of log file.
strMonth = Mid(strLog, 1, InStr(1, strLog, "/") - 1)
strLog = Mid(strLog, InStr(1, strLog, "/") + 1, Len(strLog))
strDay = Mid(strLog, 1, InStr(1, strLog, "/") - 1)
strYear = Mid(strLog, InStr(1, strLog, "/") + 1, Len(strLog))

'Build filename string.
strLog = "Log_" & strMonth & "_" & strDay & "_" & strYear & "."

On Error GoTo writeLog

    'See if this log already exsists.
    For i = 1 To 150
        Open appPath & strLog & intCounter For Input As #1
        Close #1
        intCounter = intCounter + 1
    Next i
    
writeLog:

    'Open the log file, all lines of listbox.
    Open appPath & strLog & intCounter For Output As #1
    For i = 0 To lstLogging.ListCount - 1
        lstLogging.ListIndex = i
        Write #1, lstLogging.Text
    Next i
    Close #1
    lstLogging.Clear


End Sub

Private Sub cmdSecure_Click()
    frmSecurity.Show
End Sub

Private Sub cmdSend_Click()

    'If send button pressed, send message to all then output textbox.
    sendToEveryone
    txtOutput.Text = txtOutput.Text + vbCrLf + "Me: " + txtSend.Text
    txtOutput.SelStart = Len(txtOutput.Text)
    txtSend.SetFocus
    
End Sub

Private Sub cmdSave_Click()

    Open appPath & "preferences.cfg" For Output As #1
        Write #1, txtName.Text
        Write #1, intPort
        Write #1, ChkTransfers
        Write #1, ChkIPs
        Write #1, ChkTime
        Write #1, optShare(0).Value
        Write #1, optShare(1).Value
        Write #1, ChkUpload.Value
        Write #1, frmWelcome.txtWelcome.Text
        Write #1, Dir1.Path
        Write #1, Dir2.Path
    Close #1
    
End Sub

Private Sub cmdLoadPreferences_Click()

On Error GoTo savePreferences

    Dim strTemp As String
    Open appPath & "preferences.cfg" For Input As #1
    
        'Load name.
        Input #1, strTemp
        txtName.Text = strTemp
        
        'Load port.
        Input #1, strTemp
        intPort = strTemp
                
        'Log File transfers?
        Input #1, strTemp
        ChkTransfers.Value = strTemp
        
        'Save Name/IP in log?
        Input #1, strTemp
        ChkIPs.Value = strTemp
        
        'Save time in log?
        Input #1, strTemp
        ChkTime.Value = strTemp
        
        'Share files?
        Input #1, strTemp
        optShare(0).Value = strTemp
        Input #1, strTemp
        optShare(1).Value = strTemp
        
        'Permit uploads?
        Input #1, strTemp
        ChkUpload.Value = strTemp
        
        'Get Welcome string.
        Input #1, strWelcome
        frmWelcome.txtWelcome.Text = strWelcome
        
        'Set the download folder.
        Input #1, strTemp
        Dir1.Path = strTemp
        Dir1.Refresh
        'Set the share folder.
        Input #1, strTemp
        Dir2.Path = strTemp
        Dir2.Refresh
        
    Close #1
    Exit Sub
    
savePreferences:
    cmdSave_Click

End Sub

Private Sub loadFavorites()

On Error GoTo loadFavoritesError

    
    Open appPath & "favorites.cfg" For Input As #1
    
    Do While Not EOF(1)
    
        'Load name.
        Input #1, strName
        frmConnect.lstFavName.AddItem strName
        'Load IP.
        Input #1, strIP
        frmConnect.lstFavIP.AddItem strIP
        frmConnect.lstConnections.AddItem strName & vbTab & strIP
        
    Loop
    
    Close #1
    Exit Sub
    
loadFavoritesError:
    'No favorites have been saved yet.
    saveFavorites
    
End Sub

Private Sub loadSecurity()

    On Error GoTo loadSecurityError

    Open appPath & "security.cfg" For Input As #1
    
        'Load password.
        Input #1, strPassword
        
        'Load allowed times to try at password.
        Input #1, intStrikes
        
        'Is surity on?
        Input #1, blnSecure
        
    Close #1
    Exit Sub
    
loadSecurityError:
    'Sesurity file not written yet.
    saveSecurity

End Sub

Private Sub saveSecurity()

'On Error GoTo saveSecurityError

    
    Open appPath & "security.cfg" For Output As #1
    
        'Save password.
        Write #1, strPassword
        
        'Save allowed times to try at password.
        Write #1, intStrikes
        
        'Is surity on?
        Write #1, blnSecure
        
    Close #1
    Exit Sub
    
saveSecurityError:
    'Can't write file.
    MsgBox ("Error writing Security file")
End Sub

Private Sub cmdSendSound_Click()
    
On Error GoTo noOneSelected

Dim i As Integer

    For i = 1 To intNum_Connections
        getSendersInfo i
        '*****************TEST**********Check me later***********
        
        '*****************************************************(***
        If tvwConnects.SelectedItem.key = strName & strIP Then
            
            'Save the .wav file and send it.
            MMControl1.Command = "Save"
            MMControl1.Command = "Close"
            
            'Send the sound.
            cmdSendSound.Enabled = False
            setupsendFile i, "PS_SoundFile.wav"
        End If
    Next i
    
    
    Exit Sub
    
noOneSelected:
'No name was selected in tree view.  Siwwy wabbits.
MsgBox ("You must first select who you want to send sound to from Connections/Chat window.")
'cmdSendSound.Enabled = True
'MMControl1.Command = "Open"

End Sub

Private Sub cmdSound_Click()

Dim i As Integer

On Error GoTo fileError

    For i = 1 To intNum_Connections
        getSendersInfo i
        If tvwConnects.SelectedItem.key = strName & strIP Then
            'send windows sound.
            sendToOne i, "playWindowsSound," & File4.FileName
        End If
    Next i
    
    'Play sound locally
    MMControl2.Command = "close"
    MMControl2.FileName = "c:\windows\media\" & File4.FileName
    MMControl2.Command = "Open"
    MMControl2.Command = "Play"
    
Exit Sub
    
fileError:
    'MsgBox ("There was an error while trying to send sound.")
    

End Sub

Private Sub cmdUpload_Click()
    
On Error GoTo noOneSelected2
    Dim i As Integer

    For i = 1 To intNum_Connections
        getSendersInfo i
        '*********************test******************************
        
        '**********************************************************
        If tvwConnects.SelectedItem.key = strName & strIP Then
            'send request, and null takes up space, not used.
            sendToOne i, "uploadingFile," & File3.FileName
        End If
    Next i
    
    Exit Sub

noOneSelected2:
'No name was selected in tree view.  Siwwy wabbits.
MsgBox ("You must first select who to upload to inConnections/Chat window.")


End Sub

Private Sub Command1_Click()
    
lstFiles.Clear

On Error GoTo noOneSelected
    Dim i As Integer

    For i = 1 To intNum_Connections
        getSendersInfo i
        '*********************test**************************
        '**************************************************
        If tvwConnects.SelectedItem.key = strName & strIP Then
            'send request, and null takes up space, not used.
            sendToOne i, "requestDir,null"
        End If
    Next i
    
    Exit Sub

noOneSelected:
'No name was selected in tree view.  Siwwy wabbits.
MsgBox ("You must first select who you want to get directory from in Connections/Chat window.")

End Sub


Private Sub Command2_Click()
    'Close file when ready to send.
     MMControl1.Command = "Save"
    MMControl1.Command = "Close"
    
End Sub

Private Sub Command5_Click()

    'Stop sending text to logging listbox.
    blnLog = False
    lstLogging.AddItem "** Logging stopped: " & Date & " " & Time
    
    'Show log: off in status bar.
    StatusBar1.Panels(4).Text = "Log: Off"


End Sub

Private Sub Command6_Click()

    'Start sending text to logging listbox.
    blnLog = True
    lstLogging.AddItem "** Logging started: " & Date & " " & Time
    
    'Show log: on in status bar.
    StatusBar1.Panels(4).Text = "Log: On"
    
    'Enable save log button.
    cmdSaveLog.Enabled = True
 
End Sub

Private Sub Dir1_Change()

    'Chage directory looking to selected dir.
     File1.Path = Dir1.Path
     
End Sub

Private Sub Dir2_Change()
    'Change visible filelistbox used in requests.
    File3.Path = Dir2.Path
    'change hidden filelistbox used in requests.
    File2.Path = Dir2.Path
    
End Sub

Private Sub Drive1_Change()

    'Change directory window to match drive.
    Dir1.Path = Drive1.Drive
    
End Sub

Private Sub Drive2_Change()
    Dir2.Path = Drive2.Drive
End Sub

Private Sub File4_DblClick()

'If double clicked, send selected sound.
cmdSound_Click

End Sub

Public Sub Form_Load()
    
    'Turn off timer.
    Timer2.Enabled = False
    
    'Get the applications path.
    appPath = App.Path
    If Right$(appPath, 1) <> "\" Then appPath = appPath & "\"
    
    'Set up sound recording.
    MMControl1.DeviceType = "WaveAudio"
    MMControl2.DeviceType = "WaveAudio"
    MMControl1.FileName = appPath & "PS_SoundFile.wav"
    MMControl1.Command = "Open"
    MMControl2.Command = "open"
    
    'When file is downloaded blnWav is set to true if it was a sound to play.
    'Just in case... Works without it on win98.
    blnWav = False
    
    'Put local IP in text box
    txtLocalIP.Text = Winsock1(0).LocalIP
    
    'Load preferences.cfg if present.
    cmdLoadPreferences_Click
    
    'Load Favorite connections if present.
    loadFavorites
    
    'First time value.
    intStrikes = frmSecurity.VScroll1.Value
    
    'Load Security options.
    loadSecurity
    
    'Put loaded password in text box.
    frmSecurity.txtPassword.Text = strPassword
    'Put loaded trys at password in text box.
    frmSecurity.txtStrikes.Text = intStrikes
    frmSecurity.VScroll1.Value = intStrikes
    'Turn security on?
    If blnSecure Then
        frmSecurity.Picture2.BackColor = &HFF&
        frmSecurity.Picture1.BackColor = &HCCB7B9
        frmSecurity.lblSecure.Caption = "Secure"
        frmSecurity.lblSecure.ForeColor = &HFF&
    
        'Disable properties.
        frmSecurity.VScroll1.Enabled = False
        frmSecurity.txtStrikes.Enabled = False
        frmSecurity.txtPassword.Enabled = False
    End If
    
    'IF first time used, set port to default.
    If intPort = 0 Then
        intPort = 2001
    End If
    
    'Setup linening connection zero in winsock array.
    Winsock1(0).Close
    Winsock1(0).LocalPort = intPort
    Winsock1(0).Listen
    
    'turn off send button until conection established.
    cmdSend.Enabled = False
    
    'Call the statusbar update sub
    printConnections
    
    'Turn off timer until needed.
    Timer1.Enabled = False
    
    'Show log: off in status bar.
    StatusBar1.Panels(4).Text = "Log: Off"
    
    'Turn off save log button until log is started.
    cmdSaveLog.Enabled = False
    
    'Set the path for sounds folder to widows\media.
    File4.Path = "c:\windows\media"
    
End Sub

Public Sub cmdConnect_Click(connectionIP As String)

On Error GoTo errorhandler

    Dim intArrayNumber As Integer   'The array number to use.
    Dim i As Integer
    Dim blnConnected As Boolean
    
    'Check for errors in IP address
    If connectionIP = "" Then Exit Sub
    
    'For testing!
    alreadyConnected blnConnected, connectionIP
    
    If blnConnected = True Then
        MsgBox ("You are already connected to this node.")
        txtSend.SetFocus
    Else
    
        'Show connection text in output textbox.
        txtOutput.Text = txtOutput.Text + vbCrLf + "Connecting to IP " & connectionIP & "."
        txtOutput.SelStart = Len(txtOutput.Text)
        '-----------------------

        'Search if there's a used available Winsock control.
        For i = 0 To intNum_Connections
            'Is there an loaded unused index?
            If Winsock1(i).State = sckClosed Then
        
            intArrayNumber = i ' use a used closed spot.
                Exit For
            End If
        Next i 'Looking for used open number in array.
    
        'If none was found, create a new one.
        If intArrayNumber = 0 Then
    
            'Increment number of connections.
            intNum_Connections = intNum_Connections + 1
        
            'Load a new Winsock control for this connection. Only load new one after 2.
            Load Winsock1(intNum_Connections)
            
            'Load new listbox for new connections connections list.
            Load lstConnect(intNum_Connections)
            
            'Make new security spot in array.
            ReDim Preserve intAccess(intNum_Connections)
            intAccess(intNum_Connections) = 0
        
            'Make listbox array index to use for connection info.
            ReDim Preserve Connect(intNum_Connections)
        
            'Set the winsock index to new array number.
            intArrayNumber = intNum_Connections
            
        End If
  
        'connect if you can ---------------------
        Winsock1(intArrayNumber).Close
        Winsock1(intArrayNumber).LocalPort = 0
        Winsock1(intArrayNumber).Connect connectionIP, intPort
        'Increase number of current connections for statusbar.
        intNum_ConnectionsNow = intNum_ConnectionsNow + 1
    
        'Call the statusbar update sub
        printConnections
    
        'Turn on timer and set the connection integer to send myInfo to when .2 seconds have passed.
        intChannel = intArrayNumber
        Timer1.Enabled = True
           
        'Move to tab 1 to see connection.
        SSTab1.Tab = 0
    
    End If 'If blnConnected
    
Exit Sub

errorhandler:

    'Error connecting.
    txtOutput.Text = txtOutput.Text + vbCrLf + "Failed to connect."
    txtOutput.SelStart = Len(txtOutput.Text)
    Winsock1(1).Close

End Sub

Private Sub Form_Resize()

On Error GoTo noChange

    Me.Height = 6045
    Me.Width = 7470
    Exit Sub
    
noChange:

End Sub


Private Sub Form_Unload(Cancel As Integer)

Dim i As Integer

    'Close all open ports used.
    For i = 0 To intNum_Connections
        Winsock1(i).Close
    Next i
    
    'Save favorites list
    saveFavorites
    
    'Save and close wav file.
    MMControl1.Command = "Save"
    MMControl1.Command = "Close"
    
    'Save security options
    saveSecurity
    
    'Unload forms.
    Unload frmAbout
    Unload frmConnect
    Unload frmDefaultPort
    Unload frmSecurity
    Unload frmWelcome
    
    End

End Sub

Private Sub JoinNode_Click()

On Error GoTo alreadyConnected

Dim i As Integer
Dim Index As Integer
Dim strNodeName As String
Dim strNodeIP As String

    For i = 1 To intNum_Connections
        getSendersInfo i
        '*************************test************************
        
        '*****************************************************
        If tvwConnects.SelectedItem.Parent.key = strName & strIP Then
            Index = i
            Exit For
        End If
    Next i
   
    For i = 0 To lstConnect(Index).ListCount / 3 - 1 'ListCount is 1 based, so remove 1.
        'Get node name.
        lstConnect(Index).ListIndex = i * 3
        strNodeName = lstConnect(Index).Text
        
        If strNodeName = tvwConnects.SelectedItem Then
            'Get IP.
            lstConnect(Index).ListIndex = i * 3 + 1
            strNodeIP = lstConnect(Index).Text
            Exit For
        End If
    Next i
    
    If strNodeName <> "" Then
        If MsgBox("Connect to " & strNodeName & "?", vbYesNo) = vbYes Then
            'Try to connect to selected node.
            'frmConnect.txtConnection.Text = strNodeIP
            cmdConnect_Click strNodeIP
        
        End If ' yes/no
    End If ' node <> ""
    
    Exit Sub
    
alreadyConnected:
MsgBox ("You must choose a subnode to connect to.")

End Sub

Private Sub MMControl1_Done(NotifyCode As Integer)

    'When played. Go back to start.
    MMControl1.Command = "Prev"
    
End Sub


Private Sub MMControl1_RecordClick(Cancel As Integer)
    'IF play button is hit, set timer.
    'Timer2 is set for 3 sec.
    Timer2.Enabled = True
End Sub

Private Sub mnuAbout_Click()

    'Show about form.
    frmAbout.Show
    
End Sub

Private Sub mnuExit_Click()

    Unload Me
    
End Sub


Private Sub mnuPort_Click()
    
    frmDefaultPort.Show
    
End Sub


Private Sub mnuSearchNodes_Click()

    frmConnect.Show
    frmConnect.txtSearch.SetFocus
    
End Sub


Private Sub Timer1_Timer()

    'If State <> 7 then failed to connect.
    If Winsock1(intChannel).State <> 7 Then
        intSelText = Len(txtOutput.Text)
        txtOutput.Text = txtOutput.Text + vbCrLf + "Connection " & intChannel & " Failed."
        'Failed, so decrease # of connections.
        intNum_ConnectionsNow = intNum_ConnectionsNow - 1
        ReDim Preserve Connect(intNum_ConnectionsNow)
        'Call the statusbar update sub
        printConnections
        'Select new text for color change.
        txtOutput.SelStart = intSelText
        txtOutput.SelLength = Len(txtOutput.Text)
        txtOutput.SelColor = vbRed
        'Set select to end of text when done.
        txtOutput.SelStart = Len(txtOutput.Text)
        txtSend.SetFocus
        Winsock1(intChannel).Close
        
    'Connection was successfull
    Else
        If blnSecure Then
            'Send only server info tell password given.
            secureInfo intChannel
        Else
            sendMyInfo intChannel
        
            'Turn on Send button
            cmdSend.Enabled = True
            cmdDrop.Enabled = True
            txtName.Enabled = False
            cmdAddFavorites.Enabled = True
            cmdRelay.Enabled = True
            cmdPassive.Enabled = True
            
            'Turn on sound controle
            MMControl1.Enabled = True
                    
            'Safe to turn on Send button.
            txtSend.SetFocus
        End If
    End If
    
    'Turn timer off until needed again.
    Timer1.Enabled = False

End Sub

Private Sub Timer2_Timer()
    'When 3 seconds of recording is up. Save change.
    MMControl1.Command = "Save"
    MMControl1.Command = "Close"
    MMControl1.Command = "Open"
    Timer2.Enabled = False
    
End Sub

Private Sub Timer3_Timer()

'File has stopped sending, kick start it again...
sendToOne memoryIndex, "sendFile," & memoryChannel

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.key
        Case "connect"
            frmConnect.Show
            frmConnect.txtConnection.SetFocus
        
        Case "search"
            frmConnect.Show
            frmConnect.txtSearch.SetFocus
            
        Case "welcome"
            frmWelcome.Show
            
        Case "share"
            optShare(0) = True
                        
        Case "noShare"
            optShare(1) = True
                    
    End Select
End Sub


Private Sub Toolbar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Display the menu when right mouse button is pressed
    If Button = vbRightButton Then
        PopupMenu mnuNodes
    End If
    
End Sub


Private Sub tvwConnects_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Display the menu when right mouse button is pressed
    If Button = vbRightButton Then
        PopupMenu mnuNodes
    End If

End Sub

Private Sub txtName_Change()
    If InStr(1, txtName, ",") Then
        MsgBox ("Your name can not include a comma.")
        txtName = Mid(txtName, 1, InStr(1, txtName, ",") - 1)
        
    End If
        
End Sub





Private Sub txtSend_KeyPress(KeyAscii As Integer)
    
    'If Enter key pressed, send text, clear text box.
    If KeyAscii = 13 Then
        cmdSend_Click
    End If
    
End Sub


Private Sub txtSend_KeyUp(KeyCode As Integer, Shift As Integer)

'If return hit, clear box.
If KeyCode = 13 Then txtSend.Text = ""

End Sub

Private Sub Winsock1_Close(Index As Integer)

    'Get information for disconnected node starting at 0
    getSendersInfo Index
    
    'Connection was broke by other computer.
    txtOutput.Text = txtOutput.Text + vbCrLf + strName & " Disconnected."
    txtOutput.SelStart = Len(txtOutput.Text)
    txtSend.SetFocus
        
    'Close the connection just to be safe. I guess....
    Winsock1(Index).Close
    
    'Set security to zero.
    intAccess(Index) = 0
    
On Error GoTo portScan

    'Remove node from treeview.
    tvwConnects.Nodes.Remove strName & strIP

    'Connection has left, Open spot in array.
    Connect(Index).Name = ""
    
    'Clear their connections.
    lstConnect(Index).Clear
    
    'Update log, IPs.
    If ChkIPs And blnLog Then
        lstLogging.AddItem strName & " disconnected."
    End If
    'Update log, Date and time.
    If ChkTime And blnLog Then
        lstLogging.AddItem Time
    End If
    
    txtSend.SetFocus

    'Update current connections.
    intNum_ConnectionsNow = intNum_ConnectionsNow - 1
    printConnections
    
    'If no connections, disable buttons
    If intNum_ConnectionsNow = 0 Then
        cmdSend.Enabled = False
        cmdDrop.Enabled = False
        txtName.Enabled = True
    End If
    
    SSTab1.Tab = 0 'Put view tab to connections to see that someone left.
       
    sendConnectionsToAll Index 'Update node list of all connections.
    
    Exit Sub
    
portScan:
    
    
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
'A connection was requested from the server.

Dim i As Integer
Dim intArrayNumber As Integer

'only 0 in winsock array allowed for connecting.
If Index = 0 Then

'Search if there's a used available Winsock control.
For i = 0 To intNum_Connections
    '0 in winsock1 array is used for connecting, so it is open.
    If Winsock1(i).State = sckClosed Then
        intArrayNumber = i
        Exit For
    End If
Next i
    
    'If none was found, create a new one.
    If intArrayNumber = 0 Then
    
        'Increment number of connections.
        intNum_Connections = intNum_Connections + 1
        
        'Make new security spot in array.
        ReDim Preserve intAccess(intNum_Connections)
        intAccess(intNum_Connections) = 0
        
        'Load a new Winsock control for this connection. Only load new one after 2.
        Load Winsock1(intNum_Connections)
        
        'Set the winsock index to new array number.
        intArrayNumber = intNum_Connections
        
        'Increase collection array by one, and preserve contents.
        ReDim Preserve Connect(intNum_Connections)
        
        'Use listbox array for nodes connections information.
        
        Load lstConnect(intNum_Connections)
        
    End If
    
    'Let system assign an open port to array spot. 0 = pick random.
    Winsock1(intArrayNumber).LocalPort = 0
    
    'Then accept connection on that port.
    Winsock1(intArrayNumber).Accept requestID
    
    'Enable the Send button, so you can talk back.
    cmdSend.Enabled = True
    
    'Post connection in window and set focus to send textbox.
    txtOutput.Text = txtOutput.Text + vbCrLf + "Connection with " & Winsock1(intArrayNumber).RemoteHostIP & " made."
    txtOutput.SelStart = Len(txtOutput.Text)
    txtSend.SetFocus
    
    'Increase number of current connections for statusbar.
    intNum_ConnectionsNow = intNum_ConnectionsNow + 1
    
    'Call the statusbar update sub
    printConnections
    
    'Turn on timer and set the connection integer to send myInfo to when .2 seconds have passed.
    intChannel = intArrayNumber

    'Enable clock to send myInfo in .2 seconds.
    Timer1.Enabled = True
    
End If ' index = 0

End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    
    'Error code
    On Error GoTo booboo 'a no-no happend
    
    'String for receiving data
    Dim Incoming As String
    
    'String - what to do.
    Dim strCut As String
    
    'temp index
    Dim i As Integer
    
    'Recieve incoming string over net.
    Winsock1(Index).GetData Incoming, vbString, bytesTotal
    
    'Cut off the instruction part.
    cutString strCut, Incoming
    
    'Find out what they want.
    Select Case strCut
        Case "print"
            If blnSecure And intAccess(Index) < 1 Then
                If Incoming = strPassword Or Incoming = vbCrLf & strPassword Then
                    sendConnections Index 'Send my info &
                    intAccess(Index) = 1 'Accept connection.
                Else
                    intAccess(Index) = intAccess(Index) - 1
                    sendToOne Index, "print,Password invalid. Try again:"
                    If intAccess(Index) <= -intStrikes Then
                        passwordFailed Index
                    End If 'intAccess
                End If 'Incoming
            Else
                getSendersInfo Index
                If Not passive Then
                    txtOutput.Text = txtOutput.Text & vbCrLf & strName + ": " & Incoming
                    txtOutput.SelStart = Len(txtOutput.Text)
                    txtSend.SetFocus
                    'Show first tab if someone sends message.  Anouying but needed.
                    SSTab1.Tab = 0
                End If
                If Connect(Index).relay Then
                    sendToAllButOne Index, ">>" + strName + ": " + Incoming
                
                End If
                
                
            End If
        
        Case "myInfo"
            'Extract the name
            cutString strCut, Incoming
            
            'First part is name.
            strName = strCut
            
            'Next cut is the IP address.
            cutString strCut, Incoming
            strIP = Winsock1(Index).RemoteHostIP ' strCut
            
            'Share is the next part of string.
            cutString strCut, Incoming
            
            'Do they share files?
            If strCut = "yes" Then
                blnShare = True
            Else
                blnShare = False
            End If
            
            'Index is index now.
            strIndex = Index
            
            'Add senders info.
            Connect(Index).Name = strName
            Connect(Index).IP = strIP
            Connect(Index).Sharing = blnShare
            
            'Add who connected to tree view.
            Set nodTemp = tvwConnects.Nodes.Add(, , strName & strIP, strName, "fileClosed", "fileOpen")
            tvwConnects.Nodes.Item(strName & strIP).Selected = True
            tvwConnects.Nodes.Item(strName & strIP).Expanded = True
    
            
            'Is there other connections at other end?
            If Incoming <> "0" Then
                'Get the other connections of new connection.
                getConnections Index, Incoming
               
            End If
                  
            'Update log, IPs.
            If ChkIPs And blnLog Then
                lstLogging.AddItem strName & "(" & strIP & ") Connected."
            End If
            'Update log, Date and time.
            If ChkTime And blnLog Then
                lstLogging.AddItem Time
            End If

            'Print name of who connected in output textbox.
            txtOutput.Text = txtOutput.Text & vbCrLf & strName & " is Connected"
            txtOutput.SelStart = Len(txtOutput.Text)
            txtSend.SetFocus
            
            'Turn on drop connection button, since there's at least 1 connection.
            cmdDrop.Enabled = True
            
            'Request Welcome message.
            sendToOne Index, "welcome,null"

        Case "requestContacts"
            'Refresh a connection to see who they're connected to.
            sendConnections Index
            
        Case "listOfContacts"
            'Receiving contact list from a connection.
            clearNodes Index
            cutString strCut, Incoming 'Take out first string to see if theres info coming.
            Incoming = strCut & "," & Incoming 'Then put it back in to be compatable with sub call.
            If strCut > 0 Then
                getConnections Index, Incoming
            End If
            
        Case "welcome"
            If blnSecure Then
                sendToOne Index, "printWelcome,Secure server, enter password:"
            Else
                'Send welcome message to new connection.
                sendToOne Index, "printWelcome," & strWelcome
            End If
            
        Case "printWelcome"
            txtOutput.Text = txtOutput.Text + vbCrLf + Incoming
            txtOutput.SelStart = Len(txtOutput.Text)
            txtSend.SetFocus
            sendConnectionsToAll Index

        Case "requestDir"
            'send your file directory to who asked for it.
            sendDir Index
            
        Case "showFiles"
            'Fill box with files.
            fillDir Incoming
            
        Case "requestFile"
            'Someone requested a file.
            setupsendFile Index, Incoming
            
        Case "makeFile"
            cutString strCut, Incoming
            'Open a file for writing.
            makeFile Index, Val(strCut), Incoming
            
        Case "setupFile"
            'Store the other computers channel in your channel array.
            cutString strCut, Incoming
            intChannels(strCut) = Incoming
            sendFile Index, Val(strCut)
            
        Case "uploadingFile"
            'Someone wants to upload a file.
            If ChkUpload Then
                sendToOne Index, "requestFile," & Incoming
                getSendersInfo Index
                
            'Log that it's a upload to you, not one of your downloads.
            If ChkTransfers Then
                lstLogging.AddItem Incoming & " is being uploaded from " & strName
            End If
            
            'IF uploading not permitted. Tell them.
            Else
                sendToOne Index, "print,Uploading not permited."
            End If
            
        Case "sendFile"
            'Send file block.
            'cutString strCut, Incoming
            sendFile Index, Val(Incoming)
            
        Case "moreFile"
            cutString strCut, Incoming
            moreFile Index, Val(strCut), Incoming
            
        Case "fileDone"
            'cutString strCut, Incoming
            'Turn off backup timer for downloads that stop prematurely.
            Timer3.Enabled = False
            'End of file reached, close channel, set array spot to 0(open).
            Close #Incoming
            intChannels(Incoming) = 0
            
            getSendersInfo Index
            txtOutput.Text = txtOutput.Text + vbCrLf + "File successfully downloaded from " & strName & "."
            'If logging file downloads, log it.
            If ChkTransfers And blnLog Then
                lstLogging.AddItem "File successfully downloaded from " & strName & "."
            End If
            File1.Refresh
            File2.Refresh
            File3.Refresh
            
            'Was file a sound to play?
            If blnWav Then
                MMControl1.Command = "open"
                cmdSendSound.Enabled = True
                blnWav = False
                MMControl1.Command = "Play"
            End If
            
        Case "searchFor"
            cutString strCut, Incoming
            'Search for file locally.
            fileSearch Index, strCut, Incoming
            
        Case "fileFound"
            'frmConnect.lstSearch.AddItem "Working"
            fileFound Incoming
            
        Case "playWindowsSound"
            'Play a windows wav file from windows folder.
            playWindowsSound Incoming
            
    End Select
        
booboo:

    
End Sub

Private Sub sendToEveryone()

Dim i As Integer

For i = 1 To intNum_Connections
    If Winsock1(i).State <> sckClosed Then
        Winsock1(i).SendData ("print," & txtSend.Text)
    End If
Next i

       
End Sub

Private Sub sendToAllButOne(Index As Integer, output As String)

Dim i As Integer

For i = 1 To intNum_Connections
    If Winsock1(i).State <> sckClosed And i <> Index Then
        Winsock1(i).SendData ("print," & ">" & output)
    End If
Next i

       
End Sub

Private Sub sendToOne(Index As Integer, output As String)

On Error GoTo owchie
    'Use this for sending nonchat information
    Winsock1(Index).SendData output
    
Exit Sub

owchie:

End Sub

Private Sub printConnections()

    'Update current connections in statusbar
    StatusBar1.Panels(2).Text = "Connections open: " & intNum_ConnectionsNow
         
End Sub


Private Sub cutString(strCut As String, Incoming As String)
On Error GoTo cutError

    'Seporate into 2 seporate strings with comma.
    'First get everything before the comma and put it in strControl.
    strCut = Mid(Incoming, 1, InStr(1, Incoming, ",") - 1)

On Error GoTo cutError2
    'Second get everything behind comma and put it in strData.
    Incoming = Mid(Incoming, InStr(1, Incoming, ",") + 1, Len(Incoming))
    Exit Sub
    
cutError:
    MsgBox ("cutString error #1")
    MsgBox (Incoming)
    Exit Sub
    
cutError2:
    MsgBox ("cutString error #2")
    MsgBox (strCut)
    MsgBox (Incoming)
    
End Sub

Private Sub sendMyInfo(Index As Integer)
    
Dim i As Integer
Dim output As String
Dim strContact As String

'Build up My Information string. 3 parts.
output = "myInfo," & txtName.Text & ","
output = output & txtLocalIP.Text & ","
'Does this node share files?
If optShare(0) Then
    output = output & "yes"
Else
    output = output & "no"
End If

'Add information of other connections
If intNum_ConnectionsNow > 1 Then 'IF there's other connections beside the one just made.
    getSendersInfo Index
    strContact = strName & strIP 'Get name of this connection.
    output = output & "," & intNum_ConnectionsNow - 1 'Number of other connections beside this one.
    For i = 1 To intNum_ConnectionsNow
        getSendersInfo i
        If (strName & strIP) <> strContact Then 'Not this connection? Send info then.
            output = output & "," & strName
            output = output & "," & strIP
            If blnShare = True Then
                output = output & ",yes"
            Else
                output = output & ",no"
            End If
        End If
    Next i
    
    output = output & ",null"
Else
    output = output & ",0,null"
End If

'Send it using sendToOne subroutine
sendToOne Index, output

End Sub

Private Sub sendDir(Index As Integer)

Dim i As Integer
Dim intLength As Integer
Dim output As String

'Do I share files?
If optShare(1) Then
    'Do you share files?
    sendToOne Index, "print,Directory Search denied."

Else

    'Use second list box in case ftp tab is open when
    'request for file list is made. Otherwise, will
    'select each and slowly go down list as you watch...

    'Get number of files in filelistbox2.
    intLength = File2.ListCount

    'If nothing in directory, Nothing to send.
    If intLength = 0 Then
        Exit Sub
    End If

    output = "showFiles," & intLength

    For i = 0 To intLength - 1
        'Select filenames one at a time and append to string.
        File2.ListIndex = i
        output = output & "," & File2.FileName
    Next i

    'send the directory.
    sendToOne Index, output
 
End If
End Sub

Private Sub fillDir(Incoming As String)

Dim i As Integer
Dim strCut As String

    'Get the length of list in FTP Directory.
    cutString strCut, Incoming
    i = Val(strCut)
    
    'Loop if need to do more than once.
    If i > 1 Then
        For i = 1 To i - 1
             cutString strCut, Incoming
             lstFiles.AddItem strCut
        Next i
    End If
    
    'Add last one manually, nothing after last comma.
    'Would error if cutString used on last one.
    lstFiles.AddItem Incoming
    lstFiles.ListIndex = 0

End Sub

Private Sub setupsendFile(Index As Integer, Incoming As String)

Dim intLocalChannel As Integer
On Error GoTo FileNotFound

findChannel intLocalChannel
    
If Incoming = "PS_SoundFile.wav" Then 'Not a file, a sound.
    Open appPath & Incoming For Binary As #intLocalChannel

Else
    Open File3.Path & "\" & Incoming For Binary As #intLocalChannel

End If 'Sound file?

sendToOne Index, "makeFile," & intLocalChannel & "," & Incoming

Exit Sub

FileNotFound:

sendToOne Index, "print,File not found or currenty open. Refresh directory."

End Sub

Private Sub sendFile(Index As Integer, intLocalChannel As Integer)

    'All of file sent?
    If EOF(intLocalChannel) Then
        'Turn off timer that is a backup for file transfer.
        Timer3.Enabled = False
        
        sendToOne Index, "fileDone," & intChannels(intLocalChannel)
        
        
        '***Tesing perposes
        txtOutput.Text = txtOutput.Text + vbCrLf + "File completely Sent"
        
        
        Close intLocalChannel
        intChannels(intLocalChannel) = 0
        'Log if logging on.
        If ChkTransfers And blnLog Then
            'Dim dToday As Date
            lstLogging.AddItem "File sent."
        End If
        
        'If sound was send, reenable sound.
        MMControl1.Command = "Open"
        cmdSendSound.Enabled = True
        
    'Send some more data.
    Else
        strFileString = "moreFile," & intChannels(intLocalChannel) & ","
        strFileString = strFileString & Input(3000, #intLocalChannel)
                  
        sendToOne Index, strFileString
        Timer3.Enabled = True
    End If

End Sub

Private Sub makeFile(Index As Integer, intHostChannel As Integer, Incoming As String)

Dim intLocalChannel As Integer

On Error GoTo fileOpenError

    'Find a unused channel on this system.
    findChannel intLocalChannel

    'Save to channel used on other connection.
    'use it to other computer what file is being sent.
    intChannels(intLocalChannel) = intHostChannel
    
    'First check if sound to play.
    If Incoming = "PS_SoundFile.wav" Then 'Not a file, a sound.
        MMControl1.Command = "Close"
        cmdSendSound.Enabled = False
        'Set sound vaiable(not part of the control) so it will be played.
        blnWav = True
        'Save the sound file in app.path. Not a normal download.
        Open appPath & Incoming For Binary As #intLocalChannel
    Else
        'File transfer, so save in download directory.
        Open Dir1.Path & "\" & Incoming For Binary As #intLocalChannel
    End If
    
    'Tell host what channel was set up, and your local one.
    sendToOne Index, "setupFile," & intHostChannel & "," & intLocalChannel
    
    'Post that file is being downloaded.
    getSendersInfo Index
    txtOutput.Text = txtOutput.Text & vbCrLf & "Getting " & Incoming & " from " & strName & "."
    'If logging file downloads, log it.
    If ChkTransfers And blnLog Then
        lstLogging.AddItem "Getting " & Incoming & " from " & strName & "."
    End If
    
    Exit Sub
    
fileOpenError:
    MsgBox ("Error opening file!")

End Sub

Private Sub moreFile(Index As Integer, intLocalChannel As Integer, Incoming As String)

Put intLocalChannel, , Incoming

sendToOne Index, "sendFile," & intChannels(intLocalChannel)

'If the above doesn't get sent, then try again in 5 seconds using timer3.
memoryIndex = Index
memoryChannel = intChannels(intLocalChannel)
Timer3.Enabled = True

End Sub

Private Sub getSendersInfo(Index As Integer)

On Error GoTo woopsie

    'Fill the strings with senders info in listbox.
    strName = Connect(Index).Name
    strIP = Connect(Index).IP
    blnShare = Connect(Index).Sharing
Exit Sub

woopsie:
    'Missing information, bad connection.  Try reconnecting.
End Sub

Private Sub findChannel(intLocalChannel As Integer)
Dim i As Integer

'Find unused channel to Write with/Put into.
For i = 2 To 202
    If intChannels(i) = 0 Then
        intLocalChannel = i
        Exit For
    End If
Next i

End Sub

Private Sub getConnections(Index As Integer, Incoming As String)
Dim strCut As String
Dim i As Integer
Dim key As String
Dim share As String
Dim nodeName As String 'The nodes name

    'Clear the listbox
    lstConnect(Index).Clear
    
    getSendersInfo Index
    
    cutString strCut, Incoming
    For i = 0 To Val(strCut) - 1 'Compensate for 0 based listbox.
        cutString strCut, Incoming      'Add name
        lstConnect(Index).AddItem strCut
        nodeName = strCut
        cutString strCut, Incoming      'Add IP
        lstConnect(Index).AddItem strCut
        'Make a unique key with nodeName,nodeIP, and parents IP.
        key = nodeName & strCut & strIP
        cutString strCut, Incoming      'Add share?
        lstConnect(Index).AddItem strCut
        '**********************Check back here laterr*************
        
        '*******************************************************
        If strCut = "no" Then
            Set nodTemp = tvwConnects.Nodes.Add(strName & strIP, tvwChild, key, nodeName, "fileClosed", "fileClosed")
        Else
            Set nodTemp = tvwConnects.Nodes.Add(strName & strIP, tvwChild, key, nodeName, "fileOpen", "fileOpen")
        End If
        
    Next i
    
End Sub

Private Sub sendConnections(Index As Integer)

Dim strContact As String
Dim output As String
Dim i As Integer

If intNum_ConnectionsNow > 1 Then 'IF there's other connections beside the one just made.
    getSendersInfo Index
    strContact = strName & strIP 'Get name of this connection.
    output = "listOfContacts," & intNum_ConnectionsNow - 1 'Number of other connections beside this one.
    For i = 1 To intNum_ConnectionsNow
        getSendersInfo i
        If (strName & strIP) <> strContact Then 'Not this connection? Send info then.
            output = output & "," & strName
            output = output & "," & strIP
            If blnShare = False Then
                output = output & "," & "no"
            Else
                output = output & "," & "yes"
            End If
        End If
    Next i
    
    output = output & ",null" 'Tack on an extra string for cutString sub to work right.
    
Else

    output = "listOfContacts,0,null" 'No other contacts.
    
End If

sendToOne Index, output

End Sub

Private Sub clearNodes(Index As Integer)

Dim i As Integer
Dim strCut As String

    'Are there nodes in there already?
    'Take out nodes.
    If lstConnect(Index).ListCount > 0 Then
        For i = 0 To lstConnect(Index).ListCount / 3 - 1 'ListCount is 1 based, so remove 1.
            lstConnect(Index).ListIndex = i * 3
            strCut = lstConnect(Index).Text
            'Add IP to the Key.
            lstConnect(Index).ListIndex = i * 3 + 1
            strCut = strCut & lstConnect(Index).Text
            
            getSendersInfo Index
            strCut = strCut + strIP
            'Remove node.
            tvwConnects.Nodes.Remove strCut
        Next i
    End If
    
    'Clear the nodes listbox.
    lstConnect(Index).Clear

End Sub

Private Sub sendConnectionsToAll(Index As Integer)

Dim i As Integer

For i = 1 To intNum_Connections
    If Winsock1(i).State <> sckClosed Then
        If i <> Index Then
            sendConnections i
        End If
    End If
Next i

End Sub

Private Sub alreadyConnected(blnConnected As Boolean, strConnection As String)
Dim i As Integer
'Assume not connected.
blnConnected = False

'Trying to connect to yourself???
If strConnection = txtLocalIP.Text Then
    blnConnected = True
    Exit Sub
End If

'See if already connected to this IP before connecting.
If intNum_ConnectionsNow <> 0 Then
    For i = 1 To intNum_Connections
        If Winsock1(i).State <> sckClosed Then
            If strConnection = Winsock1(i).RemoteHostIP Then
                blnConnected = True
                Exit For
            End If
        End If
    Next i
End If
End Sub

Private Sub fileSearch(Index As Integer, strCut As String, Incoming As String)

    Dim i As Integer
    Dim strTemp As String
    Dim strNodeName As String
    Dim strNodeIP   As String
    Dim strFile As String
    Dim strFiles As String
    Dim strPassOn As String 'Build this string to pass on as you take incoming apart.
    Dim strFileFound As String 'Build this string and use if file found.
    
    'The text being looked for.
    strFile = strCut
    
    'Start building passOn string.
    strPassOn = "searchFor," & strFile & "," & txtName & "," & txtLocalIP
    strFileFound = txtName & "," & txtLocalIP
    
    'get previous sender in list.
    cutString strCut, Incoming
    strNodeName = strCut
    cutString strCut, Incoming
    
    'Internet IP fix. Done this way to be compatable with previous versions.
    strNodeIP = Winsock1(Index).RemoteHostIP ' This is used to get internet IPs.
    strPassOn = strPassOn & "," & strNodeName & "," & strNodeIP
            strFileFound = strFileFound & "," & strNodeName & "," & strNodeIP
    'End of internet fix.............
    
    If Incoming <> "null" Then
        'Look for end of incoming string.
        Do Until Incoming = "null"
    
            'get next sender in list.
            cutString strCut, Incoming
            strNodeName = strCut
            cutString strCut, Incoming
            strNodeIP = strCut
    
            If strNodeName & strNodeIP <> txtName & txtLocalIP Then
                'Sender not in list so far, keep going.
                strPassOn = strPassOn & "," & strNodeName & "," & strNodeIP
                strFileFound = strFileFound & "," & strNodeName & "," & strNodeIP
            Else
                Exit Sub 'Found self in string.  Kill search.
            End If
        
        Loop 'Keep adding everyone, one by one to list.
    End If ' Incoming = "null"
    
    strPassOn = strPassOn & ",null" 'null is end of string.
    strFileFound = strFileFound & ",null" 'Same here.
    
    'Send the search to all connections except where it came from.
    For i = 1 To intNum_Connections
        If Winsock1(i).State <> sckClosed Then
            If i <> Index Then
                sendToOne i, strPassOn
            End If
        End If
    Next i
    
    
    'Now look for file.
    strFiles = "Update your Privashare"
    For i = 0 To File2.ListCount - 1
        File2.ListIndex = i
        strTemp = InStr(1, File2.FileName, strFile, 1)
        If strTemp <> "0" Then
            'File found locally, Add to strFile.
            strFiles = strFiles & "*" & File2.FileName
        End If
        
    Next i
    
    'If files found, send message back to sender, thourgh path of connections.
    If strFiles <> "Update your Privashare" Then
        sendToOne Index, "fileFound," & strFiles & "," & strFileFound
    End If
    
End Sub

Private Sub saveFavorites()

Dim i As Integer
Dim strTemp As String

On Error GoTo saveFav

If frmConnect.lstFavName.ListCount <> 0 Then
 
    Open appPath & "favorites.cfg" For Output As #1

        For i = 0 To frmConnect.lstFavName.ListCount - 1
    
            frmConnect.lstFavName.ListIndex = i
            frmConnect.lstFavIP.ListIndex = i
    
            Write #1, frmConnect.lstFavName.Text
            Write #1, frmConnect.lstFavIP.Text
        
        Next i
        
        Close #1
        
End If

Exit Sub
    
saveFav:
    MsgBox ("saving favorites didn't work")


End Sub

Private Sub fileFound(Incoming As String)

Dim i As Integer
Dim strTemp As String
Dim strFileName As String
Dim strNodeName As String
Dim strNodeIP As String
Dim strFoundName As String
Dim strFoundIP  As String
Dim blnMoreFiles As Boolean

'Get the filename there.
cutString strTemp, Incoming
strFileName = strTemp

'Get the name and ip of computer with file.
cutString strTemp, Incoming
strFoundName = strTemp
cutString strTemp, Incoming
strFoundIP = strTemp

'Remove this computers info from string.
cutString strTemp, Incoming
cutString strTemp, Incoming



If Incoming = "null" Then ' You are the one who is looking for file.
    
    'Loop until there is no more file matches at host IP.
    blnMoreFiles = False
    Do Until blnMoreFiles = True
        
        If InStr(1, strFileName, "*", 1) <> "0" Then
            cutFiles strTemp, strFileName
            If strTemp <> "Update your Privashare" Then
                frmConnect.lstSearch.AddItem strTemp & vbTab & strFoundName & vbTab & strFoundIP
                frmConnect.lstSearchIP.AddItem strFoundIP
            End If
        Else
            frmConnect.lstSearch.AddItem strFileName & vbTab & strFoundName & vbTab & strFoundIP
            frmConnect.lstSearchIP.AddItem strFoundIP
            blnMoreFiles = True
        End If
    
    Loop
    
    frmConnect.lstSearch.ListIndex = frmConnect.lstSearch.ListCount - 1
    frmConnect.lstSearchIP.ListIndex = frmConnect.lstSearch.ListCount - 1
Else

    'Get computers info, next in list, to send "fileFound" back to.
    cutString strTemp, Incoming
    strNodeName = strTemp
    cutString strTemp, Incoming
    strNodeIP = strTemp

    'Find next connection to send back "fileFound" to.
    For i = 0 To intNum_Connections - 1
        getSendersInfo i
        If (strNodeName & strNodeIP) = (strName & strIP) Then
            'Set up string to pass back in search.
            strTemp = "fileFound," & strFileName & "," & strFoundName & "," & strFoundIP & "," & strNodeName & "," & strNodeIP & "," & Incoming
            'Send string back one more computer.
            sendToOne i, strTemp
        End If
    Next i

End If ' Incoming.

End Sub


Private Sub cutFiles(strCut As String, Incoming As String)
'On Error GoTo cutError

    'Seporate into 2 seporate strings with *.
    'First get everything before the comma and put it in strControl.
    strCut = Mid(Incoming, 1, InStr(1, Incoming, "*") - 1)

    'Second get everything behind comma and put it in strData.
    Incoming = Mid(Incoming, InStr(1, Incoming, "*") + 1, Len(Incoming))
    Exit Sub
    
End Sub

Private Sub passwordFailed(Index As Integer)
 
    getSendersInfo Index
    
    tvwConnects.Nodes.Remove strName & strIP
    Winsock1(Index).Close
    
    'Set Name to nothing so you know later it's not used.
    Connect(Index).Name = ""
    
    'Set security to zero.
    intAccess(Index) = 0
                
    intNum_ConnectionsNow = intNum_ConnectionsNow - 1 'Decrease number of connections.
    'Connection was dropped.  That wasn't nice...
                
    'If no connections, disable buttons
    If intNum_ConnectionsNow = 0 Then
        cmdSend.Enabled = False
        cmdDrop.Enabled = False
        txtName.Enabled = True
    End If
                
    txtOutput.Text = txtOutput.Text + vbCrLf + strName & " failed to enter correct password."
    txtOutput.SelStart = Len(txtOutput.Text)
    txtSend.SetFocus
                
    'Update log, IPs.
    If ChkIPs And blnLog Then
        If ChkTime Then
            lstLogging.AddItem strName & " failed to enter correct password at " & Time
        Else
            lstLogging.AddItem strName & " failed to enter correct password."
        End If
    End If

    'Clear their connections.
    lstConnect(Index).Clear
    
    sendConnectionsToAll Index 'Update node list of all connections.
    
    Exit Sub
    
End Sub

Private Sub secureInfo(Index As Integer)

Dim output As String

'Build up My Information string. 3 parts.
output = "myInfo," & txtName.Text & ","
output = output & txtLocalIP.Text & ","
output = output & "no"
output = output & ",0,null"

'Send it using sendToOne subroutine
sendToOne Index, output

End Sub

Private Sub playWindowsSound(Incoming As String)

    MMControl2.Command = "close"
    MMControl2.FileName = "c:\windows\media\" & Incoming
    MMControl2.Command = "Open"
    MMControl2.Command = "Play"
    
End Sub
