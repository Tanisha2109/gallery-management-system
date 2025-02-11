VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIPDServiceUsedByPatient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Service Used By Patient"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9870
   Icon            =   "frmIPDServiceUsedByPatient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   9870
   Begin VB.TextBox txtQty 
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
      Left            =   7800
      TabIndex        =   22
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   6015
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   9375
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         DisabledPicture =   "frmIPDServiceUsedByPatient.frx":08CA
         Height          =   855
         Left            =   1680
         Picture         =   "frmIPDServiceUsedByPatient.frx":0F64
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         DisabledPicture =   "frmIPDServiceUsedByPatient.frx":15BC
         Height          =   735
         Left            =   8040
         Picture         =   "frmIPDServiceUsedByPatient.frx":1E86
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   5160
         Width           =   975
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
         Left            =   7560
         TabIndex        =   18
         Top             =   4200
         Width           =   1695
      End
      Begin VB.TextBox txtServiceName 
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
         Left            =   7560
         TabIndex        =   16
         Top             =   2880
         Width           =   1695
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
         Left            =   7560
         TabIndex        =   14
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton cmdLoadServiceUsed 
         Caption         =   "Service Used"
         DisabledPicture =   "frmIPDServiceUsedByPatient.frx":2750
         Height          =   855
         Left            =   120
         Picture         =   "frmIPDServiceUsedByPatient.frx":301A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtIPDNO 
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
         Left            =   3120
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdSearch 
         DisabledPicture =   "frmIPDServiceUsedByPatient.frx":350D
         Height          =   615
         Left            =   2400
         Picture         =   "frmIPDServiceUsedByPatient.frx":3DD7
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   735
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
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin MSComctlLib.ListView lvwBedMaster 
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   1720
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
      Begin MSComctlLib.ListView lvServiceUsede 
         Height          =   2895
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   5106
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
      Begin MSComctlLib.ListView lvServiceName 
         Height          =   3735
         Left            =   3840
         TabIndex        =   9
         Top             =   2040
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   6588
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
      Begin HMS.Duncan_DatePicker dpOnDate 
         Height          =   300
         Left            =   7560
         TabIndex        =   12
         Top             =   4800
         Width           =   1695
         _extentx        =   2990
         _extenty        =   529
         firstdayofweek  =   1
         descriptionformat=   "d mmm yyyy"
         usehandcursor   =   0   'False
         dateselected    =   41715
         shownonmonthdays=   -1  'True
         font            =   "frmIPDServiceUsedByPatient.frx":42CA
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tariff"
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
         Left            =   7560
         TabIndex        =   19
         Top             =   3960
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Service Name"
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
         Left            =   7560
         TabIndex        =   17
         Top             =   2640
         Width           =   1185
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SID"
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
         Left            =   7560
         TabIndex        =   15
         Top             =   2040
         Width           =   285
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "On Date"
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
         Left            =   7560
         TabIndex        =   13
         Top             =   4560
         Width           =   675
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Add New Service"
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
         Left            =   4560
         TabIndex        =   10
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
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
         Left            =   7560
         TabIndex        =   8
         Top             =   3240
         Width           =   1335
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
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
   End
   Begin MSComctlLib.ImageList imlLVIcons 
      Left            =   6000
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
            Picture         =   "frmIPDServiceUsedByPatient.frx":42F6
            Key             =   "Custs"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblPhysMaint 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Used In IPD"
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
      Left            =   1950
      TabIndex        =   0
      Top             =   0
      Width           =   3915
   End
End
Attribute VB_Name = "frmIPDServiceUsedByPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim listSelectedPatient As ListItem

Private Sub cmdLoadServiceUsed_Click()
Call LoadServiceUsed
End Sub

Private Sub cmdPrint_Click()
If txtIPDNO.Text = "" Then
    MsgBox "Please select patient ", vbOKOnly
    Exit Sub
End If
    On Error GoTo ErrHandler
    
    
    ConnectToDB
    ClearCommandParameters
    mobjCmd.CommandType = adCmdText
    Dim query As String
   
    query = "select * from VIEW_IPD_SERVICEUSED where IPD_NO = " & txtIPDNO.Text
    mobjCmd.CommandText = query
    Set mobjRst = mobjCmd.Execute
    Set DRIPDSERVICEUSED.DataSource = mobjRst
    DRIPDSERVICEUSED.Sections("section4").Controls("lblHeader").Caption = "Service Used By " & listSelectedPatient.SubItems(4) & ", BED NO : " & listSelectedPatient.SubItems(2)
    
    DRIPDSERVICEUSED.Show
    Exit Sub
Exit Sub
ErrHandler:
    MsgBox "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
  
End Sub

Private Sub cmdSave_Click()
    If txtIPDNO.Text = "" Or txtSID.Text = "" Or txtQty.Text = "" Then
        MsgBox "Please select properly patient  Service and Qty", vbOKOnly
        Exit Sub
    End If
    Dim objCurrLI As ListItem
    
    On Error GoTo ErrHandler
    ConnectToDB
    ClearCommandParameters
    mobjCmd.CommandType = adCmdText
    Dim query As String
    
    
    query = "select max(SUID) SUID from IPD_SERVICEUSED"
     mobjCmd.CommandText = query
    Set mobjRst = mobjCmd.Execute
    Dim intSUID  As Integer
    If Not mobjRst.EOF Then
        If IsNull(mobjRst!SUID) = True Then
            intSUID = 0
        Else
            intSUID = CInt(mobjRst!SUID)
        End If
    Else
        intSUID = 0
    
    End If
    intSUID = intSUID + 1
    
    
    query = "insert into IPD_SERVICEUSED ( SUID, IPD_NO, SID, TARIFF, ONDATE, QTY) values(" & intSUID & ", " & txtIPDNO.Text & "," & txtSID.Text & ", " & CDbl(txtTariff.Text) & ", '" & GetFormatedDate(dpOnDate.DateSelected) & "'," & CDbl(txtQty.Text) & ")"
    
    mobjCmd.CommandText = query
    Set mobjRst = mobjCmd.Execute
    DisconnectFromDB
    Call LoadServiceUsed
Exit Sub
ErrHandler:
    MsgBox "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
   
End Sub

Private Sub cmdSearch_Click()
Call LoadIPDLIST
End Sub

Private Sub Form_Load()
Call SetupGrid
Call SetupServiceUserGrid
Call SetupServiceMasterGrid
Call LoadServiceMaster
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
    txtIPDNO.Text = ""
    Dim objCurrLI   As ListItem
    lvwBedMaster.ListItems.Clear
    
    ConnectToDB
    ClearCommandParameters
    mobjCmd.CommandType = adCmdText
    Dim query As String
   
    query = "select * from view_ipd_registration where REGNO = '" & txtRegNo.Text & "' and  ISDISCHARGE = 0 "
    
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
    Set listSelectedPatient = Item
    With Item
        txtIPDNO.Text = .Text
    End With
End Sub
'-----------------------------------------------------------------------------
Private Sub SetupServiceUserGrid()
'-----------------------------------------------------------------------------
                                 
    With lvServiceUsede
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "SUID", .Width * 0.15
        .ColumnHeaders.Add , , "SNAME", .Width * 0.2
        .ColumnHeaders.Add , , "IPDNO", .Width * 0.15
        .ColumnHeaders.Add , , "TARIFF", .Width * 0.2
        .ColumnHeaders.Add , , "ONDATE", .Width * 0.15
        .ColumnHeaders.Add , , "QTY", .Width * 0.15
    End With

End Sub
Private Sub txtIPDNO_Change()
'Call LoadServiceUsed
End Sub

Private Sub LoadServiceUsed()
On Error GoTo ErrHandler
    Dim objCurrLI   As ListItem
    lvServiceUsede.ListItems.Clear
    
    ConnectToDB
    ClearCommandParameters
    mobjCmd.CommandType = adCmdText
    Dim query As String
   
    query = "select * from VIEW_IPD_SERVICEUSED where IPD_NO = " & txtIPDNO.Text
    mobjCmd.CommandText = query
    Set mobjRst = mobjCmd.Execute
    With mobjRst
        Do Until .EOF
            
            Set objCurrLI = lvServiceUsede.ListItems.Add(, , !SUID & "", , "Custs")
            objCurrLI.SubItems(1) = !SNAME & ""
            objCurrLI.SubItems(2) = !IPD_NO & ""
            objCurrLI.SubItems(3) = !TARIFF & ""
            objCurrLI.SubItems(4) = !ONDATE & ""
            objCurrLI.SubItems(5) = !QTY & ""
            .MoveNext
        Loop
    End With
    Set objCurrLI = Nothing
    Set mobjRst = Nothing
    DisconnectFromDB
Exit Sub
ErrHandler:
    MsgBox "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
  
End Sub


'-----------------------------------------------------------------------------
Private Sub SetupServiceMasterGrid()
'-----------------------------------------------------------------------------
                                 
    With lvServiceName
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "SID", .Width * 0.15
        .ColumnHeaders.Add , , "SNAME", .Width * 0.2
        .ColumnHeaders.Add , , "TARIFF", .Width * 0.15
        
    End With

End Sub

Private Sub LoadServiceMaster()
On Error GoTo ErrHandler
    Dim objCurrLI   As ListItem
    lvServiceName.ListItems.Clear
    
    ConnectToDB
    ClearCommandParameters
    mobjCmd.CommandType = adCmdText
    Dim query As String
   
    query = "select * from SERVICEMASTER order by sName"
    
    mobjCmd.CommandText = query
    Set mobjRst = mobjCmd.Execute

    
    
    With mobjRst
        Do Until .EOF
            
            Set objCurrLI = lvServiceName.ListItems.Add(, , !SID & "", , "Custs")
            objCurrLI.SubItems(1) = !SNAME & ""
            objCurrLI.SubItems(2) = !TARIFF & ""
            .MoveNext
        Loop
    End With
    
    With lvServiceName
        If .ListItems.Count > 0 Then
            Set .SelectedItem = .ListItems(1)
            lvServiceName_ItemClick .SelectedItem
        End If
    End With
    
    Set objCurrLI = Nothing
    Set mobjRst = Nothing
    DisconnectFromDB
Exit Sub
ErrHandler:
    MsgBox "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
   
End Sub


Private Sub lvServiceName_ItemClick(ByVal Item As MSComctlLib.ListItem)
With Item
        txtSID.Text = .Text
        txtServiceName.Text = .SubItems(1)
        txtTariff.Text = .SubItems(2)
        End With
End Sub

