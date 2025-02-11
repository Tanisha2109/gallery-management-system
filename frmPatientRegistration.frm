VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatientRegistration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patient Registration"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8805
   Icon            =   "frmPatientRegistration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   8805
   Begin VB.Frame frmEntry 
      Caption         =   "Entry"
      Height          =   3015
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   8535
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         DisabledPicture =   "frmPatientRegistration.frx":08CA
         Height          =   855
         Left            =   7320
         Picture         =   "frmPatientRegistration.frx":0F64
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtPhone 
         Height          =   360
         Left            =   3960
         TabIndex        =   15
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtPinCode 
         Height          =   360
         Left            =   1440
         TabIndex        =   14
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtState 
         Height          =   360
         Left            =   3960
         TabIndex        =   13
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtcity 
         Height          =   360
         Left            =   1440
         TabIndex        =   12
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         DisabledPicture =   "frmPatientRegistration.frx":15BC
         Height          =   735
         Left            =   7320
         Picture         =   "frmPatientRegistration.frx":1E86
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtRegNo 
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
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtName 
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
         Left            =   3960
         TabIndex        =   8
         Top             =   240
         Width           =   3255
      End
      Begin VB.ComboBox cmbSex 
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
         ItemData        =   "frmPatientRegistration.frx":2750
         Left            =   1440
         List            =   "frmPatientRegistration.frx":275D
         TabIndex        =   9
         Text            =   "Select Here"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtAge 
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
         Left            =   3960
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtAddress 
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
         TabIndex        =   11
         Top             =   960
         Width           =   5775
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         DisabledPicture =   "frmPatientRegistration.frx":2776
         Height          =   855
         Left            =   7320
         Picture         =   "frmPatientRegistration.frx":3040
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1080
         Width           =   975
      End
      Begin HMS.Duncan_DatePicker dpRegDate 
         Height          =   300
         Left            =   1440
         TabIndex        =   30
         Top             =   2160
         Width           =   1815
         _extentx        =   3201
         _extenty        =   529
         firstdayofweek  =   1
         descriptionformat=   "d mmm yyyy"
         usehandcursor   =   0   'False
         shownonmonthdays=   -1  'True
         font            =   "frmPatientRegistration.frx":390A
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "REG. DATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   28
         Top             =   2160
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "PHONE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3240
         TabIndex        =   27
         Top             =   1800
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "PIN CODE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   26
         Top             =   1680
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "STATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3240
         TabIndex        =   25
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label lblCity 
         AutoSize        =   -1  'True
         Caption         =   "CITY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblbed_id 
         BackStyle       =   0  'Transparent
         Caption         =   "Reg. No."
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
         TabIndex        =   22
         Top             =   240
         Width           =   735
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
         Left            =   3360
         TabIndex        =   21
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblward_type 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SEX"
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
         TabIndex        =   20
         Top             =   600
         Width           =   345
      End
      Begin VB.Label lblhod 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AGE"
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
         TabIndex        =   19
         Top             =   600
         Width           =   345
      End
      Begin VB.Label lblcharge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS"
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
         TabIndex        =   18
         Top             =   960
         Width           =   825
      End
   End
   Begin VB.Frame frmView 
      Caption         =   "View"
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8535
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         DisabledPicture =   "frmPatientRegistration.frx":3936
         Height          =   735
         Left            =   7440
         Picture         =   "frmPatientRegistration.frx":4200
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         DisabledPicture =   "frmPatientRegistration.frx":4ACA
         Height          =   735
         Left            =   7440
         Picture         =   "frmPatientRegistration.frx":5394
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1080
         Width           =   915
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         DisabledPicture =   "frmPatientRegistration.frx":5C5E
         Height          =   735
         Left            =   7440
         Picture         =   "frmPatientRegistration.frx":6528
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1920
         Width           =   915
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         DisabledPicture =   "frmPatientRegistration.frx":6DF2
         Height          =   735
         Left            =   7440
         Picture         =   "frmPatientRegistration.frx":76BC
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2760
         Width           =   915
      End
      Begin MSComctlLib.ListView lvwBedMaster 
         Height          =   3255
         Left            =   120
         TabIndex        =   1
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
      Left            =   7680
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
            Picture         =   "frmPatientRegistration.frx":7F86
            Key             =   "Custs"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblPhysMaint 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Registration Master"
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
      Left            =   1005
      TabIndex        =   23
      Top             =   120
      Width           =   5325
   End
End
Attribute VB_Name = "frmPatientRegistration"
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
cmdPrint.Enabled = isEnableFrame

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

Private Sub cmdPrint_Click()
    On Error GoTo ErrHandler
    ConnectToDB
    mobjCmd.CommandType = adCmdText
    Dim query As String
    Dim strRegDat As String
    Dim strAge As String
    
    query = "select * from  patient_registration  where REGNO = '" & txtRegNo.Text & "'"
    mobjCmd.CommandText = query
    Set mobjRst = mobjCmd.Execute
    Set DRPatientInformation.DataSource = mobjRst
    DRPatientInformation.Show
    EnableFrame False
    BlankControl
Exit Sub
ErrHandler:
MsgBox "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"

End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrHandler
If ValidateFormFields = True Then

    ConnectToDB
    mobjCmd.CommandType = adCmdText
    Dim query As String
    Dim strRegDat As String
    Dim strAge As String
    strAge = Format$(DateAdd("yyyy", -(txtAge.Text), Now), "dd MMM yyyy")
    strRegDat = Format$(dpRegDate.DateSelected, "dd MMM yyyy")
    If Action = "ADD" Then
        query = "insert into patient_registration (REGNO, NAME, AGE,SEX,STREET,CITY, STATE, PIN, PHONE,REG_DATE)" _
        & "values ('" & txtRegNo.Text & "' , '" & txtName.Text & "' , '" & strAge & "' , '" & cmbSex.Text & "' , '" & txtAddress.Text & "', '" & txtcity.Text & "', '" & txtState.Text & "', '" & txtPinCode.Text & "', '" & txtPhone.Text & "', '" & strRegDat & "')"
    ElseIf Action = "UPD" Then
        query = "update patient_registration set  NAME =  '" & txtName.Text & "', AGE =  '" & strAge & "' , SEX =  '" & cmbSex.Text & "', STREET =  '" & txtAddress.Text & "' , CITY =  '" & txtcity.Text & "', STATE =  '" & txtState.Text & "', PIN =  '" & txtPinCode.Text & "', PHONE =  '" & txtPhone.Text & "', REG_DATE =  '" & strRegDat & "' where REGNO = '" & txtRegNo & "'"
    ElseIf Action = "DEL" Then
        If MsgBox("Are You Sure To Delete This Record", vbYesNo) = vbYes Then
            query = "delete from  patient_registration  where REGNO = '" & txtRegNo.Text & "'"
        Else
             query = "select REGNO from  patient_registration  where REGNO = '" & txtRegNo.Text & "'"
        End If
    End If
    mobjCmd.CommandText = query
    mobjCmd.Execute
    DisconnectFromDB
    MsgBox "Record Successfully Save", vbOKOnly
    EnableFrame False
    LoadPatientMasterList
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
LoadPatientMasterList
End Sub

Private Function ValidateFormFields()

If Not ValidateRequiredField(txtRegNo, "Reg No") Then
        ValidateFormFields = False
        Exit Function
    End If
    
    If Not ValidateRequiredField(txtName, "Name") Then
        ValidateFormFields = False
        Exit Function
    End If
    If Not ValidateRequiredField(txtAge, "Age") Then
        ValidateFormFields = False
        Exit Function
    End If
    If Not ValidateRequiredField(txtAddress, "Address ") Then
        ValidateFormFields = False
        Exit Function
    End If
   
    If Not ValidateRequiredField(txtState, "State ") Then
        ValidateFormFields = False
        Exit Function
    End If
        If Not ValidateRequiredField(txtcity, "City ") Then
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
        .ColumnHeaders.Add , , "Reg. No", .Width * 0.15
        .ColumnHeaders.Add , , "Name", .Width * 0.2
        .ColumnHeaders.Add , , "Age", .Width * 0.2
        .ColumnHeaders.Add , , "Sex", .Width * 0.2
        .ColumnHeaders.Add , , "Add", .Width * 0.2
        .ColumnHeaders.Add , , "City", .Width * 0.15
        .ColumnHeaders.Add , , "State", .Width * 0.15
        .ColumnHeaders.Add , , "Pin", .Width * 0.15
        .ColumnHeaders.Add , , "Phone", .Width * 0.15
        .ColumnHeaders.Add , , "Reg Date", .Width * 0.15
    End With

End Sub

'-----------------------------------------------------------------------------
Private Sub LoadPatientMasterList()
'-----------------------------------------------------------------------------
    On Error GoTo ErrHandler
    Dim objCurrLI   As ListItem
    lvwBedMaster.ListItems.Clear
    
    ConnectToDB
    ClearCommandParameters
    mobjCmd.CommandType = adCmdText
    Dim query As String
    query = "select REGNO, NAME, AGE,SEX,STREET,CITY, STATE, PIN, PHONE,REG_DATE from patient_registration order by REG_DATE desc, REGNO desc"
    mobjCmd.CommandText = query
    Set mobjRst = mobjCmd.Execute
    With mobjRst
        Do Until .EOF
            
            Set objCurrLI = lvwBedMaster.ListItems.Add(, , !REGNO & "", , "Custs")
            objCurrLI.SubItems(1) = !Name & ""
            objCurrLI.SubItems(2) = DateDiff("yyyy", !AGE, Now) & ""
            objCurrLI.SubItems(3) = !SEX & ""
            objCurrLI.SubItems(4) = !STREET & ""
            objCurrLI.SubItems(5) = !CITY & ""
            objCurrLI.SubItems(6) = !State & ""
            objCurrLI.SubItems(7) = !PIN & ""
            objCurrLI.SubItems(8) = !PHONE & ""
            objCurrLI.SubItems(9) = !REG_DATE & ""
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
    Exit Sub
ErrHandler:
MsgBox "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"

End Sub


Private Sub lvwBedMaster_ItemClick(ByVal Item As MSComctlLib.ListItem)
 With Item
        txtRegNo.Text = .Text
        txtName.Text = .SubItems(1)
        txtAge.Text = .SubItems(2)
        cmbSex.Text = .SubItems(3)
        
        txtAddress.Text = .SubItems(4)
        txtcity.Text = .SubItems(5)
        txtState.Text = .SubItems(6)
        txtPinCode.Text = .SubItems(7)
        txtPhone.Text = .SubItems(8)
        dpRegDate.DateSelected = .SubItems(9)
    End With
End Sub

Private Sub BlankControl()
        txtRegNo.Text = ""
        txtName.Text = ""
        txtAge.Text = ""
        txtAddress.Text = ""
        txtcity.Text = ""
        txtState.Text = ""
        txtPinCode.Text = ""
        txtPhone.Text = ""
        

End Sub


