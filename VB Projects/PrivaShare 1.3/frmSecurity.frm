VERSION 5.00
Begin VB.Form frmSecurity 
   BackColor       =   &H00CCB7B9&
   Caption         =   "Security options"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2685
   Icon            =   "frmSecurity.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   2685
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   1720
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CCB7B9&
      Caption         =   "Allowed tries at password"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   2415
      Begin VB.VScrollBar VScroll1 
         Height          =   495
         Left            =   2040
         Max             =   3
         Min             =   1
         TabIndex        =   6
         Top             =   240
         Value           =   1
         Width           =   255
      End
      Begin VB.TextBox txtStrikes 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "2"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00CCB7B9&
         Caption         =   "Chances before connection booted"
         Height          =   450
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BackColor       =   &H00CCB7B9&
      Height          =   540
      Left            =   1320
      Picture         =   "frmSecurity.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   480
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Height          =   540
      Left            =   600
      Picture         =   "frmSecurity.frx":0884
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   480
      Width           =   540
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00CCB7B9&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00CCB7B9&
      BackStyle       =   0  'Transparent
      Caption         =   "System password:"
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
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label lblSecure 
      Alignment       =   2  'Center
      BackColor       =   &H00CCB7B9&
      Caption         =   "Unlocked"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Click on locks to turn on/off security"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   50
      Width           =   1815
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    'Make changes on close.
    strPassword = txtPassword.Text
    intStrikes = txtStrikes.Text
    frmSecurity.Hide
    
End Sub


Private Sub Picture1_Click()
    blnSecure = False
    Picture2.BackColor = &HCCB7B9
    Picture1.BackColor = &HFF&
    lblSecure.Caption = "Unlocked"
    lblSecure.ForeColor = &H80000012
    
    'Enable properties.
    VScroll1.Enabled = True
    txtStrikes.Enabled = True
    txtPassword.Enabled = True
    
End Sub

Private Sub Picture2_Click()

    Picture2.BackColor = &HFF&
    Picture1.BackColor = &HCCB7B9
    lblSecure.Caption = "Secure"
    lblSecure.ForeColor = &HFF&
    
    'Disable properties.
    VScroll1.Enabled = False
    txtStrikes.Enabled = False
    txtPassword.Enabled = False
    
    blnSecure = True
    
    'Set up password variables.
    strPassword = txtPassword.Text
    intStrikes = txtStrikes.Text
    
End Sub

Private Sub VScroll1_Change()
    
    'Choose number of chaces at password guess.
    txtStrikes = VScroll1.Value
    
End Sub

