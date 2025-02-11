VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Test"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   3120
      TabIndex        =   0
      Top             =   1200
      Width           =   2775
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim FileName As String
FileName = App.Path & "\config.txt"
Open FileName For Input As #1
Contents = Input(LOF(1), #1)
Close #1
Text1.Text = Contents
End Sub

