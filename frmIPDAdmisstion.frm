VERSION 5.00
Begin VB.Form frmIPDAdmisstion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IPD Admission"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9045
   Icon            =   "frmIPDAdmisstion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5445
   ScaleWidth      =   9045
   Begin VB.Frame frmEntrys 
      Height          =   4095
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   7815
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         DisabledPicture =   "frmIPDAdmisstion.frx":08CA
         Height          =   615
         Left            =   3960
         Picture         =   "frmIPDAdmisstion.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   120
         Width           =   675
      End
      Begin VB.TextBox txtRegDate 
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
         Left            =   5880
         TabIndex        =   25
         Top             =   840
         Width           =   1695
      End
      Begin VB.ComboBox cmbBed 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   3360
         Width           =   2295
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         DisabledPicture =   "frmIPDAdmisstion.frx":1A5E
         Height          =   735
         Left            =   5760
         Picture         =   "frmIPDAdmisstion.frx":2328
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3120
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
         Top             =   1680
         Width           =   2175
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
         Left            =   4560
         TabIndex        =   10
         Top             =   1320
         Width           =   975
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
         ItemData        =   "frmIPDAdmisstion.frx":2BF2
         Left            =   1440
         List            =   "frmIPDAdmisstion.frx":2BFF
         TabIndex        =   9
         Text            =   "Select Here"
         Top             =   1320
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
         Left            =   1440
         TabIndex        =   8
         Top             =   840
         Width           =   3015
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
      Begin VB.CommandButton cmdSearch 
         DisabledPicture =   "frmIPDAdmisstion.frx":2C18
         Height          =   615
         Left            =   3240
         Picture         =   "frmIPDAdmisstion.frx":34E2
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox txtcity 
         Height          =   360
         Left            =   4560
         TabIndex        =   5
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtState 
         Height          =   360
         Left            =   1440
         TabIndex        =   4
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtPinCode 
         Height          =   360
         Left            =   4560
         TabIndex        =   3
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtPhone 
         Height          =   360
         Left            =   1440
         TabIndex        =   2
         Top             =   2400
         Width           =   1575
      End
      Begin HMS.Duncan_DatePicker dpRegDate 
         Height          =   300
         Left            =   3720
         TabIndex        =   28
         Top             =   3360
         Width           =   1455
         _extentx        =   4048
         _extenty        =   529
         firstdayofweek  =   1
         descriptionformat=   "d mmm yyyy"
         usehandcursor   =   0   'False
         dateselected    =   41711
         shownonmonthdays=   -1  'True
         font            =   "frmIPDAdmisstion.frx":39D5
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reg Date"
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
         Left            =   4680
         TabIndex        =   26
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "BED"
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
         Left            =   240
         TabIndex        =   23
         Top             =   3000
         Width           =   345
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   7800
         Y1              =   2880
         Y2              =   2880
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
         TabIndex        =   21
         Top             =   1680
         Width           =   825
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
         TabIndex        =   20
         Top             =   1320
         Width           =   345
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
         TabIndex        =   19
         Top             =   1320
         Width           =   345
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
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   495
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
         TabIndex        =   17
         Top             =   240
         Width           =   735
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
         Left            =   3720
         TabIndex        =   16
         Top             =   1800
         Width           =   375
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
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   555
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
         Left            =   3720
         TabIndex        =   14
         Top             =   2160
         Width           =   810
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
         Left            =   120
         TabIndex        =   13
         Top             =   2520
         Width           =   600
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
         Left            =   2640
         TabIndex        =   12
         Top             =   3360
         Width           =   885
      End
   End
   Begin VB.Label lblPhysMaint 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IN DOOR Registration"
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
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   7815
   End
End
Attribute VB_Name = "frmIPDAdmisstion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdNew_Click()
frmPatientRegistration.Show
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrHandler
'On Error Resume Next
    ConnectToDB
    mobjCmd.CommandType = adCmdText
    Dim query As String
    Dim IPDNo As Integer
    Dim strBEDID As Variant
    Dim strBID As String
    
    strBEDID = Split(cmbBed.Text, "-")
    strBID = strBEDID(1)
    'strBEDID = strBEDID(1)
    query = "select *  from ipd_registration where REGNO = '" & txtRegNo.Text & "' and isdischarge = 0"
     mobjCmd.CommandText = query
    Set mobjRst = mobjCmd.Execute
    If Not mobjRst.EOF Then
        MsgBox "Patient Exist", vbOKOnly
        DisconnectFromDB
        Exit Sub
    End If
    query = "select max(ipd_no) IPDNO from ipd_registration"
     mobjCmd.CommandText = query
    Set mobjRst = mobjCmd.Execute
    If Not mobjRst.EOF Then
        If mobjRst!IPDNo = Null Then
            IPDNo = 0
        Else
            IPDNo = CInt(mobjRst!IPDNo)
        End If
    Else
        IPDNo = 0
    
    End If
    IPDNo = IPDNo + 1
    mobjCmd.CommandText = "insert into ipd_registration(ipd_no,regno,bed_no,admitdate, isdischarge) values(" & IPDNo & ",'" & txtRegNo.Text & "', '" & strBID & "', '" & GetFormatedDate(dpRegDate.DateSelected) & "', 0)"
    mobjCmd.Execute
    
    MsgBox "Saved", vbOKOnly
   
    
    DisconnectFromDB
Call LoadWard
Exit Sub
ErrHandler:
    MsgBox "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
End Sub

Private Sub cmdSearch_Click()
'On Error GoTo ErrHandler
On Error Resume Next
    ConnectToDB
    mobjCmd.CommandType = adCmdText
    Dim query As String
    Dim strRegDat As String
    Dim strAge As String
    
    query = "select * from  patient_registration  where REGNO = '" & txtRegNo.Text & "'"
    mobjCmd.CommandText = query
    Set mobjRst = mobjCmd.Execute
    If mobjRst.EOF Then
    
        MsgBox "Invalid Reg. No", vbOKOnly
        cmdSave.Enabled = False
    Else
        cmdSave.Enabled = True
        With mobjRst
        Do Until .EOF
            
            
            txtName.Text = !Name
            txtAge.Text = DateDiff("yyyy", !AGE, Now)
            cmbSex.Text = !SEX
           txtAddress.Text = !STREET
            txtcity.Text = !CITY
            txtState.Text = !State
            txtPinCode.Text = !PIN
            txtPhone.Text = !PHONE
            'claderRegDate.Value = !REG_DATE
            txtRegDate.Text = !REG_DATE
            .MoveNext
        Loop
    End With
    End If
    
    DisconnectFromDB
Exit Sub

ErrHandler:
    'MsgBox "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
    
End Sub

Private Sub dpDischargeOn_DateChanged(ByVal FromDate As Date, ByVal ToDate As Date)

End Sub

Private Sub Form_Load()
cmdSave.Enabled = False

Call LoadWard
End Sub


Private Sub LoadWard()
'On Error GoTo ErrHandler

    ConnectToDB
    mobjCmd.CommandType = adCmdText
    Dim query As String
    Dim strRegDat As String
    Dim strAge As String
    'query = "select ward_name || ' ' || ward_type || '-' || bed_id WName from  bed_record "
    query = "select ward_name || ' ' || ward_type || '-' || bed_id WName from  bed_record  where bed_id not in (select bed_no from ipd_registration where isdischarge =0)"
    'query = "select bed_no WName from ipd_registration where isdischarge ='0'"
    mobjCmd.CommandText = query
    Set mobjRst = mobjCmd.Execute
   cmbBed.Clear
        With mobjRst
        Do Until .EOF
            
            cmbBed.AddItem (!WName)
            
            .MoveNext
        Loop
    End With
    
    
    DisconnectFromDB

    Exit Sub
End Sub

