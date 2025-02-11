VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIPDBillPayment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bill Payment"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9825
   Icon            =   "frmIPDBillPayment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   9825
   Begin VB.Frame Frame3 
      Caption         =   "Make Payment"
      Height          =   1335
      Left            =   240
      TabIndex        =   14
      Top             =   4440
      Width           =   9375
      Begin VB.CommandButton cmdPaymentHistoryPrint 
         Caption         =   "View / Print Payment History"
         DisabledPicture =   "frmIPDBillPayment.frx":08CA
         Height          =   735
         Left            =   6600
         Picture         =   "frmIPDBillPayment.frx":0F64
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         DisabledPicture =   "frmIPDBillPayment.frx":15BC
         Height          =   735
         Left            =   5400
         Picture         =   "frmIPDBillPayment.frx":1E86
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtToPaidAMount 
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
         Left            =   1080
         TabIndex        =   16
         Top             =   600
         Width           =   1575
      End
      Begin HMS.Duncan_DatePicker dpOnDate 
         Height          =   300
         Left            =   3720
         TabIndex        =   18
         Top             =   600
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
      Begin VB.Label Label6 
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
         Left            =   2880
         TabIndex        =   17
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         TabIndex        =   15
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Payment Details"
      Height          =   1335
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   9375
      Begin VB.CommandButton cmdPrint 
         Caption         =   "View / Print Service List"
         DisabledPicture =   "frmIPDBillPayment.frx":2750
         Height          =   735
         Left            =   7440
         Picture         =   "frmIPDBillPayment.frx":2DEA
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdUpdateDiscount 
         Caption         =   "Update Discount"
         DisabledPicture =   "frmIPDBillPayment.frx":3442
         Height          =   735
         Left            =   5760
         Picture         =   "frmIPDBillPayment.frx":3D0C
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtTotalDues 
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
         Left            =   4080
         TabIndex        =   20
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtDiscount 
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
         Left            =   4080
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtTotalPaid 
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
         Left            =   1080
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtTotalBill 
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
         Left            =   1080
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Dues"
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
         Left            =   2760
         TabIndex        =   19
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Discount"
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
         Left            =   2760
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Paid"
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
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Bill."
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
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pateint Details"
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   9375
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
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdSearch 
         DisabledPicture =   "frmIPDBillPayment.frx":45D6
         Height          =   615
         Left            =   2640
         Picture         =   "frmIPDBillPayment.frx":4EA0
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   735
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
         Left            =   3360
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSComctlLib.ListView lvwBedMaster 
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Top             =   840
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
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
   End
   Begin MSComctlLib.ImageList imlLVIcons 
      Left            =   6600
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
            Picture         =   "frmIPDBillPayment.frx":5393
            Key             =   "Custs"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblPhysMaint 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IPD Bill Payment"
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
      Left            =   2955
      TabIndex        =   1
      Top             =   240
      Width           =   3345
   End
End
Attribute VB_Name = "frmIPDBillPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim listSelectedPatient As ListItem

Private Sub cmdPaymentHistoryPrint_Click()
If txtIPDNO.Text = "" Then
    MsgBox "Please select patient ", vbOKOnly
    Exit Sub
End If
    On Error GoTo ErrHandler
    
    
    ConnectToDB
    ClearCommandParameters
    mobjCmd.CommandType = adCmdText
    Dim query As String
   
    query = "select * from IPD_PAYMENTHISTORY where IPD_NO = " & txtIPDNO.Text
    mobjCmd.CommandText = query
    Set mobjRst = mobjCmd.Execute
    Set DRIPDPAYMENTHISTORY.DataSource = mobjRst
    DRIPDPAYMENTHISTORY.Sections("section4").Controls("lblHeader").Caption = "Paid By " & listSelectedPatient.SubItems(4) & ", BED NO : " & listSelectedPatient.SubItems(2) & " For IPD No: " & txtIPDNO.Text
    
    DRIPDPAYMENTHISTORY.Show
    Exit Sub
Exit Sub
ErrHandler:
    MsgBox "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
  

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
If txtIPDNO.Text = "" Or txtToPaidAMount.Text = "" Then
    MsgBox "Please select patient first and enter to paid amount", vbOKOnly
    Exit Sub
End If
    On Error GoTo ErrHandler
   
    ConnectToDB
    ClearCommandParameters
    mobjCmd.CommandType = adCmdText
    Dim query As String
   
    query = "select max(IPDPHID) IPDPHID from IPD_PAYMENTHISTORY"
     mobjCmd.CommandText = query
    Set mobjRst = mobjCmd.Execute
    Dim intIPDPHID As Integer
    If Not mobjRst.EOF Then
        If IsNull(mobjRst!IPDPHID) = True Then
            intIPDPHID = 0
        Else
            intIPDPHID = CInt(mobjRst!IPDPHID)
        End If
    Else
        intIPDPHID = 0
    
    End If
    intIPDPHID = intIPDPHID + 1
    query = "insert into  IPD_PAYMENTHISTORY (IPDPHID, IPD_NO, ONDATE, PAID) values (" & intIPDPHID & ", " & txtIPDNO.Text & ", '" & GetFormatedDate(dpOnDate.DateSelected) & "', " & CDbl(txtToPaidAMount.Text) & ")"
    mobjCmd.CommandText = query
    mobjCmd.Execute
    
    'Account history
    
    query = "select max(PAID) PAID from PATIENT_ACCOUNT"
     mobjCmd.CommandText = query
    Set mobjRst = mobjCmd.Execute
    Dim intPAID As Integer
    If Not mobjRst.EOF Then
        If IsNull(mobjRst!PAID) = True Then
            intPAID = 0
        Else
            intPAID = CInt(mobjRst!PAID)
        End If
    Else
        intPAID = 0
    
    End If
    intPAID = intPAID + 1
    query = "insert into  PATIENT_ACCOUNT (PAID, REGNO, ONDATE, NARRATIONS, CREDIT) values (" & intPAID & ", '" & listSelectedPatient.SubItems(3) & "', '" & GetFormatedDate(dpOnDate.DateSelected) & "', 'Paid For IPD No : " & txtIPDNO.Text & "', " & CDbl(txtToPaidAMount.Text) & ")"
    mobjCmd.CommandText = query
    mobjCmd.Execute
    Set mobjRst = Nothing
    DisconnectFromDB
      Call LoadPaymentDetails
Exit Sub
ErrHandler:
    MsgBox "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"

End Sub

Private Sub cmdSearch_Click()
Call LoadIPDLIST
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
     Set objCurrLI = Nothing
    Set mobjRst = Nothing
    DisconnectFromDB
    With lvwBedMaster
        If .ListItems.Count > 0 Then
            Set .SelectedItem = .ListItems(1)
            lvwBedMaster_ItemClick .SelectedItem
        End If
    End With
    
   
Exit Sub
ErrHandler:
    MsgBox "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
   
End Sub

Private Sub cmdUpdateDiscount_Click()
    If txtIPDNO.Text = "" Then
    MsgBox "Please select Patient", vbOKOnly
Exit Sub
End If
On Error GoTo ErrHandler
   
    ConnectToDB
    ClearCommandParameters
    mobjCmd.CommandType = adCmdText
    Dim query As String
   
    query = "update ipd_registration set DISCOUNT = " & txtDiscount.Text & "   where IPD_NO = " & txtIPDNO.Text
    mobjCmd.CommandText = query
    mobjCmd.Execute
    Set mobjRst = Nothing
    DisconnectFromDB
      Call LoadPaymentDetails
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

Private Sub lvwBedMaster_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set listSelectedPatient = Item
    With Item
        txtIPDNO.Text = .Text
    End With
    Call LoadPaymentDetails
End Sub

Private Sub LoadPaymentDetails()
On Error GoTo ErrHandler
   
    ConnectToDB
    ClearCommandParameters
    mobjCmd.CommandType = adCmdText
    Dim query As String
   
    query = "update ipd_registration set TOTALBILL = (select sum(TARIFF * QTY ) from  IPD_SERVICEUSED where IPD_NO = " & txtIPDNO.Text & " )   where IPD_NO = " & txtIPDNO.Text
    mobjCmd.CommandText = query
    mobjCmd.Execute
   
   
    query = "select * from view_ipd_registration where REGNO = '" & txtRegNo.Text & "'"
    
    mobjCmd.CommandText = query
    Set mobjRst = mobjCmd.Execute

    
    
    With mobjRst
        Do Until .EOF
            
            txtTotalBill.Text = !TOTALBILL
            txtDiscount.Text = !DISCOUNT
            If IsNull(!TOTALPAID) = True Then
                txtTotalPaid.Text = "0"
            Else
                txtTotalPaid.Text = !TOTALPAID
            End If
            .MoveNext
        Loop
    End With
   
    
   
    Set mobjRst = Nothing
    DisconnectFromDB
    
    txtTotalDues.Text = CDbl(txtTotalBill.Text) - (CDbl(txtDiscount.Text) + CDbl(txtTotalPaid.Text))
    
Exit Sub
ErrHandler:
    MsgBox "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
   
End Sub
