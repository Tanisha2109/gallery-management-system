VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmvieIPDRegistration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List Of IPD Registration"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10005
   Icon            =   "frmvieIPDRegistration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   10005
   Begin VB.Frame frmView 
      Caption         =   "View"
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin VB.OptionButton rbByDischage 
         Caption         =   "By Discharge Date"
         Height          =   195
         Left            =   3840
         TabIndex        =   11
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton rbAdmission 
         Caption         =   "By Admission Date"
         Height          =   195
         Left            =   3840
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin HMS.Duncan_DatePicker dpOnDate 
         Height          =   300
         Left            =   5640
         TabIndex        =   9
         Top             =   480
         Width           =   1455
         _ExtentX        =   4048
         _ExtentY        =   529
         FirstDayOfWeek  =   1
         DescriptionFormat=   "d mmm yyyy"
         UseHandCursor   =   0   'False
         DateSelected    =   41711
         ShowNonMonthDays=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdRunningPatientPrint 
         Caption         =   "Print Runnig Patient"
         DisabledPicture =   "frmvieIPDRegistration.frx":08CA
         Height          =   735
         Left            =   1800
         Picture         =   "frmvieIPDRegistration.frx":0F64
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdRunningPatient 
         Caption         =   "Running Patients"
         DisabledPicture =   "frmvieIPDRegistration.frx":15BC
         Height          =   735
         Left            =   240
         Picture         =   "frmvieIPDRegistration.frx":1E86
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdSearch 
         DisabledPicture =   "frmvieIPDRegistration.frx":2379
         Height          =   735
         Left            =   7320
         Picture         =   "frmvieIPDRegistration.frx":2C43
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtIPDNO 
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
         Left            =   120
         TabIndex        =   3
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton cmdDischarge 
         Caption         =   "Discharge"
         DisabledPicture =   "frmvieIPDRegistration.frx":3136
         Height          =   855
         Left            =   3000
         Picture         =   "frmvieIPDRegistration.frx":3A00
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         DisabledPicture =   "frmvieIPDRegistration.frx":42CA
         Height          =   735
         Left            =   8160
         Picture         =   "frmvieIPDRegistration.frx":4964
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin MSComctlLib.ListView lvwBedMaster 
         Height          =   3375
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   5953
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
      Begin MSComctlLib.ImageList imlLVIcons 
         Left            =   8640
         Top             =   4680
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
               Picture         =   "frmvieIPDRegistration.frx":4FBC
               Key             =   "Custs"
            EndProperty
         EndProperty
      End
      Begin HMS.Duncan_DatePicker dpDischargeOn 
         Height          =   300
         Left            =   1320
         TabIndex        =   12
         Top             =   5280
         Width           =   1455
         _ExtentX        =   4048
         _ExtentY        =   529
         FirstDayOfWeek  =   1
         DescriptionFormat=   "d mmm yyyy"
         UseHandCursor   =   0   'False
         DateSelected    =   41711
         ShowNonMonthDays=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Discharge On"
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
         Left            =   1320
         TabIndex        =   13
         Top             =   4920
         Width           =   1455
      End
      Begin VB.Line Line1 
         X1              =   3720
         X2              =   3720
         Y1              =   120
         Y2              =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "IPDNo"
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
         TabIndex        =   6
         Top             =   4920
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmvieIPDRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intSelection As Integer
Private Sub cmdCancel_Click()
If Not MsgBox("Are You Sure To Selected Patient Has Discharge?", vbYesNo) = vbYes Then
    Exit Sub
    End If
        'On Error GoTo ErrHandler
        On Error Resume Next
        ConnectToDB
        ClearCommandParameters
        mobjCmd.CommandType = adCmdText
        Dim query As String
        'query = "UPDATE ipd_registration SET ISDISCHARGE = 1, DDATE = '" & GetFormatedDate(dpDischargeOn.DateSelected) & "'  where IPDNO = " & txtIPDNO.Text & ""
        query = "UPDATE ABCV SET ISDISCHARGE = 1, DDATE = '" & GetFormatedDate(dpDischargeOn.DateSelected) & "'  where IPDNO = 1"
        mobjCmd.CommandText = query
        Set mobjRst = mobjCmd.Execute
        DisconnectFromDB
        MsgBox "Data Successfully Saved", vbOKOnly
        Call LoadIPDLIST
End Sub

Private Sub cmdDischarge_Click()
        If Not MsgBox("Are You Sure To Selected Patient Has Discharge?", vbYesNo) = vbYes Then
            Exit Sub
        End If
        On Error GoTo ErrHandler
        'On Error Resume Next
        ConnectToDB
        ClearCommandParameters
        mobjCmd.CommandType = adCmdText
        Dim query As String
        query = "UPDATE ipd_registration SET ISDISCHARGE = 1, DISCHARGEDATE = '" & GetFormatedDate(dpDischargeOn.DateSelected) & "'  where IPD_NO = " & txtIPDNO.Text & ""
        'query = "UPDATE ABCV SET ISDISCHARGE = 1, DDATE = '" & GetFormatedDate(dpDischargeOn.DateSelected) & "'  where IPDNO = 1"
        mobjCmd.CommandText = query
        Set mobjRst = mobjCmd.Execute
        DisconnectFromDB
        MsgBox "Data Successfully Saved", vbOKOnly
        Call LoadIPDLIST
        Exit Sub
ErrHandler:
MsgBox "Error # :" & Err.Description

        
End Sub

Private Sub cmdPrint_Click()
 On Error GoTo ErrHandler
    ConnectToDB
    ClearCommandParameters
    mobjCmd.CommandType = adCmdText
    Dim query As String
    If intSelection = 1 Then
     query = "select * from view_ipd_registration where ONDATE = '" & GetFormatedDate(dpOnDate.DateSelected) & "' order by Name Asc"
     DRIPDList.Sections("section4").Controls("lblHeader").Caption = "New Admited Patient List In IPD As On " & GetFormatedDate(dpOnDate.DateSelected)
    ElseIf intSelection = 2 Then
        query = "select * from view_ipd_registration where DDATE = '" & GetFormatedDate(dpOnDate.DateSelected) & "' order by Name Asc"
         DRIPDList.Sections("section4").Controls("lblHeader").Caption = "Discharge Patient List From IPD As On " & GetFormatedDate(dpOnDate.DateSelected)
    End If
    mobjCmd.CommandText = query
    Set mobjRst = mobjCmd.Execute
    Set DRIPDList.DataSource = mobjRst
    DRIPDList.Show
    Exit Sub
ErrHandler:
    MsgBox "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
   
End Sub

Private Sub cmdRunningPatient_Click()
intSelection = 0
cmdDischarge.Enabled = True
Call LoadIPDLIST
End Sub

Private Sub cmdRunningPatientPrint_Click()

   On Error GoTo ErrHandler
    ConnectToDB
    ClearCommandParameters
    mobjCmd.CommandType = adCmdText
    Dim query As String
    query = "select * from view_IPD_registration where ISDISCHARGE = 0  order by Name"
    mobjCmd.CommandText = query
    Set mobjRst = mobjCmd.Execute
    Set DRIPDList.DataSource = mobjRst
    DRIPDList.Sections("section4").Controls("lblHeader").Caption = "Running Patient In IPD"
    DRIPDList.Show
    Exit Sub
ErrHandler:
    MsgBox "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"

End Sub

Private Sub cmdSearch_Click()
Call LoadIPDLIST
End Sub

Private Sub cmdUpdate_Click()

End Sub

Private Sub Form_Load()
    Call SetupGrid
End Sub

'-----------------------------------------------------------------------------
Private Sub SetupGrid()
'-----------------------------------------------------------------------------
                                 
    With lvwBedMaster
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "IPDNO", .Width * 0.15
        .ColumnHeaders.Add , , "OnDate", .Width * 0.15
        .ColumnHeaders.Add , , "BEDNO", .Width * 0.2
        .ColumnHeaders.Add , , "RegNo", .Width * 0.15
        .ColumnHeaders.Add , , "Name", .Width * 0.2
        .ColumnHeaders.Add , , "Sex", .Width * 0.2
        .ColumnHeaders.Add , , "Address", .Width * 0.2
        .ColumnHeaders.Add , , "Phone", .Width * 0.2
        .ColumnHeaders.Add , , "ISDISCHARGE", .Width * 0.2
        
    End With

End Sub

'-----------------------------------------------------------------------------
Private Sub LoadIPDLIST()
'-----------------------------------------------------------------------------
    On Error GoTo ErrHandler
    Dim objCurrLI   As ListItem
    lvwBedMaster.ListItems.Clear
    
    ConnectToDB
    ClearCommandParameters
    mobjCmd.CommandType = adCmdText
    Dim query As String
    If intSelection = 0 Then
    query = "select * from view_ipd_registration where ISDISCHARGE = 0 order by ONDATE desc "
    ElseIf intSelection = 1 Then
     query = "select * from view_ipd_registration where ONDATE = '" & GetFormatedDate(dpOnDate.DateSelected) & "' order by Name Asc"
    ElseIf intSelection = 2 Then
    query = "select * from view_ipd_registration where DDATE = '" & GetFormatedDate(dpOnDate.DateSelected) & "' order by Name Asc"
    End If
    mobjCmd.CommandText = query
    Set mobjRst = mobjCmd.Execute
    With mobjRst
        Do Until .EOF
            
            Set objCurrLI = lvwBedMaster.ListItems.Add(, , !IPDNo & "", , "Custs")
            objCurrLI.SubItems(1) = !ONDATE & ""
            objCurrLI.SubItems(2) = !BEDNO & ""
            objCurrLI.SubItems(3) = !REGNO & ""
            objCurrLI.SubItems(4) = !Name & ""
            objCurrLI.SubItems(5) = !SEX & ""
            objCurrLI.SubItems(6) = !STREET & ""
            objCurrLI.SubItems(7) = !PHONE & ""
            objCurrLI.SubItems(8) = !ISDISCHARGE & ""
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
        txtIPDNO.Text = .Text
       
       
    End With
End Sub

Private Sub rbAdmission_Click()
intSelection = 1
cmdDischarge.Enabled = False
End Sub

Private Sub rbByDischage_Click()
intSelection = 2
cmdDischarge.Enabled = False
End Sub
