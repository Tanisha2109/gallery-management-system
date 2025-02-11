VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBedmaster 
   Caption         =   "Bed Master Entry"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmView 
      Caption         =   "View"
      Height          =   3615
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   8535
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         DisabledPicture =   "frmBedmaster.frx":0000
         Height          =   735
         Left            =   7440
         Picture         =   "frmBedmaster.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2760
         Width           =   915
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         DisabledPicture =   "frmBedmaster.frx":1194
         Height          =   735
         Left            =   7440
         Picture         =   "frmBedmaster.frx":1A5E
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1920
         Width           =   915
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         DisabledPicture =   "frmBedmaster.frx":2328
         Height          =   735
         Left            =   7440
         Picture         =   "frmBedmaster.frx":2BF2
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1080
         Width           =   915
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         DisabledPicture =   "frmBedmaster.frx":34BC
         Height          =   735
         Left            =   7440
         Picture         =   "frmBedmaster.frx":3D86
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   915
      End
      Begin MSComctlLib.ListView lvwBedMaster 
         Height          =   3255
         Left            =   120
         TabIndex        =   17
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
   Begin VB.Frame frmEntry 
      Caption         =   "Entry"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   8535
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         DisabledPicture =   "frmBedmaster.frx":4650
         Height          =   855
         Left            =   7440
         Picture         =   "frmBedmaster.frx":4F1A
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1080
         Width           =   975
      End
      Begin VB.ComboBox cmbstatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmBedmaster.frx":57E4
         Left            =   4560
         List            =   "frmBedmaster.frx":57EE
         TabIndex        =   10
         Text            =   "Select Here"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtchrge 
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
         Left            =   1440
         TabIndex        =   9
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtHOD 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   4560
         TabIndex        =   7
         Top             =   720
         Width           =   2655
      End
      Begin VB.ComboBox cmbwardtype 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmBedmaster.frx":5802
         Left            =   1440
         List            =   "frmBedmaster.frx":5812
         TabIndex        =   5
         Text            =   "Select Here"
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtWardname 
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
         Left            =   4560
         TabIndex        =   3
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtBedid 
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
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         DisabledPicture =   "frmBedmaster.frx":5834
         Height          =   735
         Left            =   7440
         Picture         =   "frmBedmaster.frx":60FE
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblcharge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CHARGE"
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
         Top             =   1320
         Width           =   705
      End
      Begin VB.Label lblstatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STATUS"
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
         Left            =   3360
         TabIndex        =   11
         Top             =   1200
         Width           =   690
      End
      Begin VB.Label lblhod 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HOD"
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
         Left            =   3360
         TabIndex        =   8
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lblward_type 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WARD TYPE"
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
         TabIndex        =   6
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label lblward_name 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WARD NAME"
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
         Left            =   3360
         TabIndex        =   4
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label lblbed_id 
         BackStyle       =   0  'Transparent
         Caption         =   "BED ID"
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
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
   Begin MSComctlLib.ImageList imlLVIcons 
      Left            =   7800
      Top             =   -120
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
            Picture         =   "frmBedmaster.frx":69C8
            Key             =   "Custs"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblPhysMaint 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bed Master"
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
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   7335
   End
End
Attribute VB_Name = "frmBedmaster"
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
    If Action = "ADD" Then
        query = "insert into bed_record (BED_ID, WARD_NAME, WARD_TYPE,HOD,CHARGE,STATUS) values ('" & txtBedid.Text & "' , '" & txtWardname.Text & "' , '" & cmbwardtype.Text & "' , '" & txtHOD.Text & "' , " & CDbl(txtchrge.Text) & " , '" & cmbstatus.Text & "')"
    ElseIf Action = "UPD" Then
        query = "update bed_record set  WARD_NAME =  '" & txtWardname.Text & "', WARD_TYPE =  '" & cmbwardtype.Text & "' , HOD =  '" & txtHOD.Text & "' ,CHARGE = " & CDbl(txtchrge.Text) & " ,STATUS = '" & cmbstatus.Text & "' where BED_ID = '" & txtBedid.Text & "'"
    ElseIf Action = "DEL" Then
        query = "delete from  bed_record  where BED_ID = '" & txtBedid.Text & "'"
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

If Not ValidateRequiredField(txtBedid, "Bed ID") Then
        ValidateFormFields = False
        Exit Function
    End If
    
    If Not ValidateRequiredField(txtchrge, "Charge") Then
        ValidateFormFields = False
        Exit Function
    End If
    If Not ValidateRequiredField(txtHOD, "HOD") Then
        ValidateFormFields = False
        Exit Function
    End If
    If Not ValidateRequiredField(txtWardname, "Ward Name") Then
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
        .ColumnHeaders.Add , , "Bed ID", .Width * 0.15
        .ColumnHeaders.Add , , "Ward Name", .Width * 0.2
        .ColumnHeaders.Add , , "Ward Type", .Width * 0.2
        .ColumnHeaders.Add , , "HOD", .Width * 0.2
        .ColumnHeaders.Add , , "Charge", .Width * 0.15
        .ColumnHeaders.Add , , "Status", .Width * 0.15
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
    query = "select * from bed_record"
    mobjCmd.CommandText = query
   Set mobjRst = mobjCmd.Execute
  
    
    
    With mobjRst
        Do Until .EOF
            
            Set objCurrLI = lvwBedMaster.ListItems.Add(, , !BED_ID & "", , "Custs")
            objCurrLI.SubItems(1) = !WARD_NAME & ""
            objCurrLI.SubItems(2) = !WARD_TYPE & ""
            objCurrLI.SubItems(3) = !HOD & ""
            objCurrLI.SubItems(4) = CStr(!Charge) & ""
            objCurrLI.SubItems(5) = !Status
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
        txtBedid.Text = .Text
        txtchrge.Text = .SubItems(4)
        txtHOD.Text = .SubItems(3)
        txtWardname.Text = .SubItems(1)
        cmbwardtype.Text = .SubItems(2)
        cmbstatus.Text = .SubItems(5)
    End With
End Sub

Private Sub BlankControl()
txtBedid.Text = ""
txtchrge.Text = ""
txtHOD.Text = ""
txtWardname.Text = ""

End Sub
