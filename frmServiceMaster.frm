VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServiceMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Service Master"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8775
   Icon            =   "frmServiceMaster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   8775
   Begin VB.Frame frmEntry 
      Caption         =   "Entry"
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   8535
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         DisabledPicture =   "frmServiceMaster.frx":08CA
         Height          =   735
         Left            =   6480
         Picture         =   "frmServiceMaster.frx":0F64
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtSID 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7080
         TabIndex        =   14
         Top             =   360
         Width           =   180
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         DisabledPicture =   "frmServiceMaster.frx":15BC
         Height          =   735
         Left            =   7440
         Picture         =   "frmServiceMaster.frx":1E86
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtServiceName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   9
         Top             =   360
         Width           =   5415
      End
      Begin VB.TextBox txtTariff 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         DisabledPicture =   "frmServiceMaster.frx":2750
         Height          =   855
         Left            =   7440
         Picture         =   "frmServiceMaster.frx":301A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblward_name 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SERVICE  NAME"
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
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label lblcharge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TARIFF"
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
         Left            =   840
         TabIndex        =   11
         Top             =   840
         Width           =   570
      End
   End
   Begin VB.Frame frmView 
      Caption         =   "View"
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   8535
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         DisabledPicture =   "frmServiceMaster.frx":38E4
         Height          =   735
         Left            =   7440
         Picture         =   "frmServiceMaster.frx":41AE
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         DisabledPicture =   "frmServiceMaster.frx":4A78
         Height          =   735
         Left            =   7440
         Picture         =   "frmServiceMaster.frx":5342
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1080
         Width           =   915
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         DisabledPicture =   "frmServiceMaster.frx":5C0C
         Height          =   735
         Left            =   7440
         Picture         =   "frmServiceMaster.frx":64D6
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1920
         Width           =   915
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         DisabledPicture =   "frmServiceMaster.frx":6DA0
         Height          =   735
         Left            =   7440
         Picture         =   "frmServiceMaster.frx":766A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2760
         Width           =   915
      End
      Begin MSComctlLib.ListView lvwBedMaster 
         Height          =   3255
         Left            =   120
         TabIndex        =   5
         Top             =   240
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
   Begin MSComctlLib.ImageList imlLVIcons 
      Left            =   7800
      Top             =   0
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
            Picture         =   "frmServiceMaster.frx":7F34
            Key             =   "Custs"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblPhysMaint 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Master"
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
      Left            =   2460
      TabIndex        =   13
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "frmServiceMaster"
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
    Dim lngSID As Integer
    query = "select max(SID) SID from servicemaster"
     mobjCmd.CommandText = query
    Set mobjRst = mobjCmd.Execute
    If Not mobjRst.EOF Then
        If IsNull(mobjRst!SID) = True Then
            lngSID = 0
        Else
            lngSID = CInt(mobjRst!SID)
        End If
    Else
        lngSID = 0
    End If
    lngSID = lngSID + 1
    
    If Action = "ADD" Then
        query = "insert into servicemaster (sid, sname, tariff) values (" & lngSID & " , '" & txtServiceName.Text & "' , " & CDbl(txtTariff.Text) & ")"
    ElseIf Action = "UPD" Then
        query = "update servicemaster set  sname =  '" & txtServiceName.Text & "', tariff =  " & CDbl(txtTariff.Text) & " where sid = " & txtSID.Text
    ElseIf Action = "DEL" Then
        query = "delete from  SERVICEMASTER  where sid = " & txtSID.Text
    End If
    mobjCmd.CommandText = query
    mobjCmd.Execute
    DisconnectFromDB
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

If Not ValidateRequiredField(txtServiceName, "Service Name") Then
        ValidateFormFields = False
        Exit Function
    End If
    
    If Not ValidateRequiredField(txtTariff, "Tariff") Then
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
        .ColumnHeaders.Add , , "ServiceID", .Width * 0.15
        .ColumnHeaders.Add , , "Name", .Width * 0.2
        .ColumnHeaders.Add , , "Tariff", .Width * 0.2
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
    query = "select * from servicemaster"
    mobjCmd.CommandText = query
   Set mobjRst = mobjCmd.Execute
  
    
    
    With mobjRst
        Do Until .EOF
            
            Set objCurrLI = lvwBedMaster.ListItems.Add(, , !SID & "", , "Custs")
            objCurrLI.SubItems(1) = !SNAME & ""
            objCurrLI.SubItems(2) = !TARIFF & ""
            
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
 With Item
        txtSID.Text = .Text
        txtServiceName.Text = .SubItems(1)
        txtTariff.Text = .SubItems(2)
        
    End With
End Sub

Private Sub BlankControl()
txtSID.Text = ""
txtServiceName.Text = ""
txtTariff.Text = ""

End Sub

