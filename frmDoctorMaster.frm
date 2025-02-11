VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDoctorMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doctor Master"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9300
   Icon            =   "frmDoctorMaster.frx":0000
   LinkTopic       =   "Doctor Master"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   9300
   Begin VB.Frame frmView 
      Caption         =   "View"
      Height          =   3615
      Left            =   480
      TabIndex        =   5
      Top             =   1080
      Width           =   8535
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         DisabledPicture =   "frmDoctorMaster.frx":08CA
         Height          =   735
         Left            =   7440
         Picture         =   "frmDoctorMaster.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2760
         Width           =   915
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         DisabledPicture =   "frmDoctorMaster.frx":1A5E
         Height          =   735
         Left            =   7440
         Picture         =   "frmDoctorMaster.frx":2328
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1920
         Width           =   915
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         DisabledPicture =   "frmDoctorMaster.frx":2BF2
         Height          =   735
         Left            =   7440
         Picture         =   "frmDoctorMaster.frx":34BC
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1080
         Width           =   915
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         DisabledPicture =   "frmDoctorMaster.frx":3D86
         Height          =   735
         Left            =   7440
         Picture         =   "frmDoctorMaster.frx":4650
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   915
      End
      Begin MSComctlLib.ListView lvwBedMaster 
         Height          =   3255
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imlLVIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame frmEntry 
      Caption         =   "Entry"
      Height          =   3015
      Left            =   360
      TabIndex        =   0
      Top             =   4680
      Width           =   8535
      Begin VB.OptionButton rbNoSergen 
         Caption         =   "No"
         Height          =   375
         Left            =   6360
         TabIndex        =   27
         Top             =   2400
         Width           =   615
      End
      Begin VB.OptionButton rbIsSurgen 
         Caption         =   "Yes"
         Height          =   375
         Left            =   5640
         TabIndex        =   26
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox txtSpecialisolation 
         Height          =   405
         Left            =   5640
         TabIndex        =   25
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtPhoneno 
         Height          =   405
         Left            =   5640
         TabIndex        =   24
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtAddress 
         Height          =   405
         Left            =   5640
         TabIndex        =   23
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtLincese 
         Height          =   375
         Left            =   1560
         TabIndex        =   22
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtDegree 
         Height          =   375
         Left            =   1560
         TabIndex        =   21
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtDName 
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   360
         Width           =   5655
      End
      Begin VB.TextBox txtCode 
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         DisabledPicture =   "frmDoctorMaster.frx":4F1A
         Height          =   855
         Left            =   7440
         Picture         =   "frmDoctorMaster.frx":57E4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         DisabledPicture =   "frmDoctorMaster.frx":60AE
         Height          =   735
         Left            =   7440
         Picture         =   "frmDoctorMaster.frx":6978
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Surgeon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3840
         TabIndex        =   20
         Top             =   2280
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Special Isolation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3840
         TabIndex        =   19
         Top             =   1920
         Width           =   1755
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Phone No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3840
         TabIndex        =   18
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3840
         TabIndex        =   17
         Top             =   960
         Width           =   885
      End
      Begin VB.Label lblLincese 
         AutoSize        =   -1  'True
         Caption         =   "License"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   16
         Top             =   1920
         Width           =   825
      End
      Begin VB.Label lblDegree 
         AutoSize        =   -1  'True
         Caption         =   "Degree"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   15
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label lblcharge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   600
         TabIndex        =   4
         Top             =   840
         Width           =   435
      End
      Begin VB.Label lblward_name 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NAME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   600
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
   End
   Begin MSComctlLib.ImageList imlLVIcons 
      Left            =   8400
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDoctorMaster.frx":7242
            Key             =   "Custs"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   4560
      TabIndex        =   14
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label lblPhysMaint 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Master"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   480
      Left            =   2880
      TabIndex        =   11
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "frmDoctorMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Action As String


Private Sub cmdAdd_Click()
 EnableFrame True
 Action = "ADD"
End Sub


Private Sub EnableFrame(isEnableFrame As Boolean)

frmEntry.Enabled = isEnableFrame
cmdSave.Enabled = isEnableFrame
cmdCancel.Enabled = isEnableFrame


frmView.Enabled = Not isEnableFrame
cmdAdd.Enabled = Not isEnableFrame
cmdUpdate.Enabled = Not isEnableFrame
cmdDelete.Enabled = Not isEnableFrame
cmdClose.Enabled = Not isEnableFrame
End Sub

Private Sub cmdCancel_Click()
'isEnableFrame = False
 EnableFrame False
 
End Sub

Private Sub cmdClose_Click()
End
End Sub

Private Sub cmdDelete_Click()
EnableFrame True
Action = "DEL"
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrHandler
If ValidateFormFields = True Then
    
    ConnectToDB
    mobjCmd.CommandType = adCmdText
    Dim query As String
    
    Dim intIsSurgen As Integer
    If rbIsSurgen.Value = True Then
        intIsSurgen = 1
    Else
        intIsSurgen = 0
    End If
    
    
    If Action = "ADD" Then
        query = "insert into Doctor (DID, dname,DDegree,LICENSE,SURGEON,SPECIAL_ISOLATION ,ADDRESS, PHONENO) values ('" & txtCode.Text & "' , '" & txtDName.Text & "','" & txtDegree.Text & "','" & txtLincese.Text & "'," & intIsSurgen & ",'" & txtSpecialisolation.Text & "','" & txtAddress.Text & "','" & txtPhoneno.Text & "')"
        'query = "insert into Doctor (dname) values ( '" & txtDName.Text & "')"
    ElseIf Action = "UPD" Then
        'query = "update Doctor set  dname =  '" & txtDName.Text & "' where did = '" & txtCode.Text & "'"
        query = "update Doctor set  dname =  '" & txtDName.Text & "', DDegree = '" & txtDegree.Text & "', LICENSE = '" & txtLincese.Text & "', Surgeon = " & intIsSurgen & ", SPECIAL_ISOLATION= '" & txtSpecialisolation.Text & "', Address = '" & txtAddress.Text & "', Phoneno = '" & txtPhoneno.Text & "' where did = '" & txtCode.Text & "'"
    ElseIf Action = "DEL" Then
        query = "delete from  Doctor  where did = '" & txtCode.Text & "'"
    End If
    mobjCmd.CommandText = query
    mobjCmd.Execute
    DisconnectFromDB


'
'    If Action = "ADD" Then
'        If SetBedMaster(txtBedid.Text, txtWardname.Text, cmbwardtype.Text, txtHOD.Text, CDbl(txtchrge.Text), cmbstatus.Text) = True Then
'    ElseIf Action = "UPD" Then
    
    MsgBox "Record Successfully Save", vbOKOnly
    EnableFrame False
    LoadBedMasterList
    
    BlankControl
End If
Exit Sub
ErrHandler:
MsgBox "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"

End Sub

Private Sub cmdUpdate_Click()
EnableFrame True
Action = "UPD"
End Sub

Private Sub Form_Load()
EnableFrame False
CenterForm Me
SetupCustLVCols
LoadBedMasterList
End Sub

Private Function ValidateFormFields()

If Not ValidateRequiredField(txtDName, "Name") Then
        ValidateFormFields = False
        Exit Function
    End If
    
    If Not ValidateRequiredField(txtCode, "Code") Then
        ValidateFormFields = False
        Exit Function
    End If
    
        
    ValidateFormFields = True

End Function

'-----------------------------------------------------------------------------
Private Sub SetupCustLVCols()
'-----------------------------------------------------------------------------
                                 
    With lvwBedMaster
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Code", .Width * 0.15
        .ColumnHeaders.Add , , "Name", .Width * 0.2
        
        .ColumnHeaders.Add , , "DEGREE", .Width * 0.15
        .ColumnHeaders.Add , , "LICENSE", .Width * 0.2
        .ColumnHeaders.Add , , "SURGEON", .Width * 0.15
        .ColumnHeaders.Add , , "S_ISOLATION", .Width * 0.2
        .ColumnHeaders.Add , , "ADDRESS", .Width * 0.3
        .ColumnHeaders.Add , , "PHONENO", .Width * 0.15
    End With

End Sub

'-----------------------------------------------------------------------------
Private Sub LoadBedMasterList()
'-----------------------------------------------------------------------------
                                 
    Dim objCurrLI   As ListItem
    lvwBedMaster.ListItems.Clear
    
     ConnectToDB
    ClearCommandParameters
    mobjCmd.CommandType = adCmdText
    Dim query As String
    query = "select * from Doctor order by DName"
    mobjCmd.CommandText = query
   Set mobjRst = mobjCmd.Execute
  
    
    
    With mobjRst
        Do Until .EOF
            
            Set objCurrLI = lvwBedMaster.ListItems.Add(, , !DID & "", , "Custs")
            objCurrLI.SubItems(1) = !DName & ""
             objCurrLI.SubItems(2) = !DDEGREE & ""
              objCurrLI.SubItems(3) = !LICENSE & ""
               objCurrLI.SubItems(4) = !SURGEON & ""
                objCurrLI.SubItems(5) = !SPECIAL_ISOLATION & ""
                 objCurrLI.SubItems(6) = !ADDRESS & ""
                  objCurrLI.SubItems(7) = !PHONENO & ""
          
            
            .MoveNext
        Loop
    End With
    
    With lvwBedMaster
        If .ListItems.Count > 0 Then
            Set .SelectedItem = .ListItems(1)
            lvwBedMaster_ItemClick .SelectedItem
        End If
    End With
    
    Set objCurrLI = Nothing
    Set mobjRst = Nothing
    DisconnectFromDB
End Sub


Private Sub lvwBedMaster_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
 With Item
        txtCode.Text = .Text
        txtDName.Text = .SubItems(1)
        txtDegree.Text = .SubItems(2)
        txtLincese.Text = .SubItems(3)
       If IsNull(.SubItems(4)) = True Then
        rbNoSergen.Value = True
       ElseIf .SubItems(4) = 1 Then
        rbIsSurgen.Value = True
       Else
        rbNoSergen.Value = True
       End If
       
       txtSpecialisolation.Text = .SubItems(5)
       txtAddress.Text = .SubItems(6)
       txtPhoneno.Text = .SubItems(7)
      
       
    End With
End Sub

Private Sub BlankControl()
txtCode.Text = ""
txtDName.Text = ""
txtDegree.Text = ""
txtLincese.Text = ""
txtAddress.Text = ""
txtPhoneno.Text = ""
txtSpecialisolation.Text = ""


End Sub




