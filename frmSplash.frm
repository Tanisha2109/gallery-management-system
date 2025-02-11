VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5730
   ClientLeft      =   2925
   ClientTop       =   2730
   ClientWidth     =   9150
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":1042
   ScaleHeight     =   5730
   ScaleWidth      =   9150
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   240
      Top             =   4920
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "SWETA KUMARI"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Top             =   1380
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "SHWETA KUMARI"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Top             =   13900
      Width           =   5415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "Developed By :-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   5415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "PRAGYA BHARTI"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   1050
      Width           =   5415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "(BCA-Final Year)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   5415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "MAGADH UNIVERSITY, BODH GAYA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Call GoToLogin
End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
 Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Call GoToLogin
    
End Sub

Private Sub GoToLogin()
    Load frmLogin
    frmLogin.Show
    Unload Me
    
End Sub
