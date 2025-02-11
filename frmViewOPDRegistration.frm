VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewOPDRegistration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List OPD  Registration"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10020
   Icon            =   "frmViewOPDRegistration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6495
   ScaleWidth      =   10020
   Begin VB.Frame frmView 
      Caption         =   "View"
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
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
         Left            =   5235
         TabIndex        =   13
         Top             =   5400
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.TextBox txtFee 
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
         TabIndex        =   11
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton cmdReceivedPayment 
         Caption         =   "Received Payment"
         DisabledPicture =   "frmViewOPDRegistration.frx":08CA
         Height          =   855
         Left            =   5160
         Picture         =   "frmViewOPDRegistration.frx":0A8C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4920
         Width           =   1600
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         DisabledPicture =   "frmViewOPDRegistration.frx":0D87
         Height          =   735
         Left            =   3960
         Picture         =   "frmViewOPDRegistration.frx":1421
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Completed"
         DisabledPicture =   "frmViewOPDRegistration.frx":1A79
         Height          =   855
         Left            =   7080
         Picture         =   "frmViewOPDRegistration.frx":2343
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4920
         Width           =   1600
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         DisabledPicture =   "frmViewOPDRegistration.frx":2C0D
         Height          =   855
         Left            =   1800
         Picture         =   "frmViewOPDRegistration.frx":34D7
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4920
         Width           =   1600
      End
      Begin VB.TextBox txtOPDID 
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
         Left            =   480
         TabIndex        =   4
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton cmdSearch 
         DisabledPicture =   "frmViewOPDRegistration.frx":3DA1
         Height          =   735
         Left            =   3120
         Picture         =   "frmViewOPDRegistration.frx":466B
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin MSComctlLib.ListView lvwBedMaster 
         Height          =   3615
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   6376
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
         Left            =   8280
         Top             =   240
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
               Picture         =   "frmViewOPDRegistration.frx":4B5E
               Key             =   "Custs"
            EndProperty
         EndProperty
      End
      Begin HMS.Duncan_DatePicker dpOnDate 
         Height          =   300
         Left            =   1440
         TabIndex        =   10
         Top             =   360
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
         Caption         =   "Fee"
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
         Left            =   4080
         TabIndex        =   12
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "OPDID"
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
         Left            =   600
         TabIndex        =   5
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label lblbed_id 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmViewOPDRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
If Not MsgBox("Are You Sure To Cancel This?", vbYesNo) = vbYes Then
    Exit Sub
    End If
        On Error GoTo ErrHandler
        'On Error Resume Next
        ConnectToDB
        mobjCmd.CommandType = adCmdText
        Dim query As String
        query = "UPDATE OPD_registration SET ISCANCEL = 1  where OPDID = " & txtOPDID.Text & ""
        mobjCmd.CommandText = query
        mobjCmd.Execute
        DisconnectFromDB
        Call LoadOPDLIST
        Exit Sub
ErrHandler:
    MsgBox "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"

End Sub

Private Sub cmdPrint_Click()
On Error GoTo ErrHandler
    ConnectToDB
    ClearCommandParameters
    mobjCmd.CommandType = adCmdText
    Dim query As String
    query = "select * from view_opd_registration where ondate = '" & GetFormatedDate(dpOnDate.DateSelected) & "'  order by Name"
    mobjCmd.CommandText = query
    Set mobjRst = mobjCmd.Execute
    Set DROPDList.DataSource = mobjRst
    DROPDList.Show
    Exit Sub
ErrHandler:
    MsgBox "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"

End Sub

Private Sub cmdReceivedPayment_Click()
        If Not MsgBox("Are You Sure Payment Received?", vbYesNo) = vbYes Then
            Exit Sub
        End If
        On Error GoTo ErrHandler
        'On Error Resume Next
        ConnectToDB
        mobjCmd.CommandType = adCmdText
        Dim query As String
        query = "UPDATE OPD_registration SET ISPAIDFEE = 1  where OPDID = " & txtOPDID.Text & ""
        mobjCmd.CommandText = query
        mobjCmd.Execute
        
        'Patient_Account insetion
        query = "select max(PAID) PAID from Patient_Account"
        mobjCmd.CommandText = query
        Set mobjRst = mobjCmd.Execute
        Dim PAID As Integer
        If Not mobjRst.EOF Then
            If IsNull(mobjRst!PAID) = True Then
                PAID = 0
            Else
                PAID = CInt(mobjRst!PAID)
            End If
        Else
            PAID = 0
    
    End If
    PAID = PAID + 1
        
        query = "insert into Patient_Account(PAID, REGNO, OnDate , CREDIT , Narrations ) values (" & PAID & ",'" & txtRegNo.Text & "', '" & GetFormatedDate(Date) & "', " & txtFee.Text & ", 'Fee Paid For OPDID : " & txtOPDID.Text & "')"
        mobjCmd.CommandText = query
        mobjCmd.Execute
        
        DisconnectFromDB
        Call LoadOPDLIST
        Exit Sub
ErrHandler:
MsgBox "Error # :" & Err.Description
End Sub


Private Sub cmdSearch_Click()
Call LoadOPDLIST
End Sub

Private Sub cmdUpdate_Click()
If Not MsgBox("Are You Sure To Complete This?", vbYesNo) = vbYes Then
    Exit Sub
    End If
        On Error GoTo ErrHandler
        'On Error Resume Next
        ConnectToDB
        mobjCmd.CommandType = adCmdText
        Dim query As String
        query = "UPDATE OPD_registration SET ISCOMPLETE = 1  where OPDID = " & txtOPDID.Text & ""
        mobjCmd.CommandText = query
        mobjCmd.Execute
        DisconnectFromDB
Call LoadOPDLIST
Exit Sub
ErrHandler:
    MsgBox "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"

End Sub

Private Sub Form_Load()
    Call SetupGrid
End Sub

'-----------------------------------------------------------------------------
Private Sub SetupGrid()
'-----------------------------------------------------------------------------
                                 
    With lvwBedMaster
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "OPDID", .Width * 0.15
        .ColumnHeaders.Add , , "OnDate", .Width * 0.15
        .ColumnHeaders.Add , , "Time", .Width * 0.2
        .ColumnHeaders.Add , , "DID", .Width * 0.15
        .ColumnHeaders.Add , , "RegNo", .Width * 0.15
        .ColumnHeaders.Add , , "Name", .Width * 0.2
        .ColumnHeaders.Add , , "Sex", .Width * 0.2
        .ColumnHeaders.Add , , "Address", .Width * 0.2
        .ColumnHeaders.Add , , "Phone", .Width * 0.2
        .ColumnHeaders.Add , , "FEE", .Width * 0.2
        
    End With

End Sub

'-----------------------------------------------------------------------------
Private Sub LoadOPDLIST()
'-----------------------------------------------------------------------------
     On Error GoTo ErrHandler
    Dim objCurrLI   As ListItem
    lvwBedMaster.ListItems.Clear
    
    ConnectToDB
    ClearCommandParameters
    mobjCmd.CommandType = adCmdText
    Dim query As String
    query = "select * from view_opd_registration where ondate = '" & GetFormatedDate(dpOnDate.DateSelected) & "' AND ISCANCEL=0 AND ISCOMPLETE=0 order by Name"
    mobjCmd.CommandText = query
    Set mobjRst = mobjCmd.Execute

    
    
    With mobjRst
        Do Until .EOF
            
            Set objCurrLI = lvwBedMaster.ListItems.Add(, , !OPDID & "", , "Custs")
            objCurrLI.SubItems(1) = !ONDATE & ""
            objCurrLI.SubItems(2) = !Time & ""
            objCurrLI.SubItems(3) = !DID & ""
            objCurrLI.SubItems(4) = !REGNO & ""
            objCurrLI.SubItems(5) = !Name & ""
            objCurrLI.SubItems(6) = !SEX & ""
            objCurrLI.SubItems(7) = !STREET & ""
            objCurrLI.SubItems(8) = !PHONE & ""
            objCurrLI.SubItems(9) = !FEE & ""
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
        txtOPDID.Text = .Text
       txtFee.Text = .SubItems(9)
       txtRegNo.Text = .SubItems(4)
       
    End With
End Sub
