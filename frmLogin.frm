VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login..."
   ClientHeight    =   2880
   ClientLeft      =   5535
   ClientTop       =   4590
   ClientWidth     =   4140
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4140
   Begin VB.Frame Frame1 
      Caption         =   "Login........."
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   495
         Left            =   2520
         TabIndex        =   6
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "&Login.."
         Height          =   495
         Left            =   1200
         TabIndex        =   5
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "$"
         TabIndex        =   4
         Text            =   "password"
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtUserName 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1320
         TabIndex        =   3
         Text            =   "enter user name"
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "User Name"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
 If MsgBox("Are you sure to exit?", vbYesNo) = vbYes Then
  End
  End If
End Sub

'-----------------------------------------------------------------------------
Private Function ValidateFormFields() As Boolean
'-----------------------------------------------------------------------------
    
    If Not ValidateRequiredField(txtUserName, "User Name") Then
        ValidateFormFields = False
        Exit Function
    End If
    
    If Not ValidateRequiredField(txtPassword, "Password") Then
        ValidateFormFields = False
        Exit Function
    End If
   
    ValidateFormFields = True
    
End Function

Private Sub cmdLogin_Click()
    If Not ValidateFormFields Then Exit Sub
    On Error GoTo ErrHandler
        If Not AuthenticateUser(txtUserName.Text, txtPassword.Text) Then
            MsgBox "invalid username & password, please try again", vbOKOnly
        Else
            Load MDIMain
            MDIMain.Caption = "Welecome " & txtUserName.Text & " " & MDIMain.Caption
            MDIMain.Show
            Unload Me
        End If
    Exit Sub
ErrHandler:
    MsgBox "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
   
End Sub

Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub


Private Sub txtPassword_Click()
txtPassword.Text = ""
End Sub

Private Sub txtUserName_Click()
txtUserName.Text = ""
txtPassword.Text = ""
End Sub

Private Sub txtUserName_GotFocus()
txtUserName.Text = ""
txtPassword.Text = ""
End Sub
