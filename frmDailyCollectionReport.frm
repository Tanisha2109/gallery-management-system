VERSION 5.00
Begin VB.Form frmDailyCollectionReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Collection Report"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7560
   Icon            =   "frmDailyCollectionReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2745
   ScaleWidth      =   7560
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   6615
      Begin VB.CommandButton cmdPaymentHistoryPrint 
         Caption         =   "Print"
         DisabledPicture =   "frmDailyCollectionReport.frx":08CA
         Height          =   735
         Left            =   5280
         Picture         =   "frmDailyCollectionReport.frx":0F64
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin HMS.Duncan_DatePicker dpFromDate 
         Height          =   300
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   1455
         _extentx        =   4048
         _extenty        =   529
         firstdayofweek  =   1
         descriptionformat=   "d mmm yyyy"
         usehandcursor   =   0
         dateselected    =   41711
         shownonmonthdays=   -1
         font            =   "frmDailyCollectionReport.frx":15BC
      End
      Begin HMS.Duncan_DatePicker dpToDate 
         Height          =   300
         Left            =   3720
         TabIndex        =   4
         Top             =   360
         Width           =   1455
         _extentx        =   4048
         _extenty        =   529
         firstdayofweek  =   1
         descriptionformat=   "d mmm yyyy"
         usehandcursor   =   0
         dateselected    =   41711
         shownonmonthdays=   -1
         font            =   "frmDailyCollectionReport.frx":15E8
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
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
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
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
         Width           =   975
      End
   End
   Begin VB.Label lblPhysMaint 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Daily Collection Report"
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
      Left            =   870
      TabIndex        =   3
      Top             =   600
      Width           =   5565
   End
End
Attribute VB_Name = "frmDailyCollectionReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPaymentHistoryPrint_Click()

    On Error GoTo ErrHandler
    Dim dblTotalCR As Double
    Dim dblTotalDR As Double
    
    ConnectToDB
    ClearCommandParameters
    mobjCmd.CommandType = adCmdText
    Dim query As String
   
    query = "select * from VIEW_PATIENT_ACCOUNT where ONDATE >= '" & GetFormatedDate(dpFromDate.DateSelected) & "' and ONDATE <= '" & GetFormatedDate(dpToDate.DateSelected) & "'"
    mobjCmd.CommandText = query
    Set mobjRst = mobjCmd.Execute
    
    
    With mobjRst
        Do Until .EOF
            dblTotalCR = dblTotalCR + CDbl(!CREDIT)
            dblTotalDR = dblTotalDR + CDbl(!DEBIT)
            
            .MoveNext
        Loop
    End With
    
    
    Set DRDailyCollectionReport.DataSource = mobjRst
    DRDailyCollectionReport.Sections("section4").Controls("lblHeader").Caption = "Between " & GetFormatedDate(dpFromDate.DateSelected) & " And  " & GetFormatedDate(dpToDate.DateSelected)
    DRDailyCollectionReport.Sections("section5").Controls("lblTotalDR").Caption = dblTotalDR
    DRDailyCollectionReport.Sections("section5").Controls("lblToalCR").Caption = dblTotalCR
    DRDailyCollectionReport.Show
    Exit Sub
Exit Sub
ErrHandler:
    MsgBox "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
  


End Sub
