VERSION 5.00
Begin VB.Form frmOPDRegistration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OPD Registration"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9375
   Icon            =   "frmOPDRegistration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   9375
   Begin VB.Frame frmEntry 
      Height          =   4455
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   7815
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
         Left            =   1680
         TabIndex        =   30
         Top             =   3840
         Width           =   1575
      End
      Begin VB.TextBox txtTiming 
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
         Left            =   1680
         TabIndex        =   28
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox txtPhone 
         Height          =   360
         Left            =   1440
         TabIndex        =   13
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox txtPinCode 
         Height          =   360
         Left            =   4560
         TabIndex        =   12
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtState 
         Height          =   360
         Left            =   1440
         TabIndex        =   11
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtcity 
         Height          =   360
         Left            =   4560
         TabIndex        =   10
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton cmdSearch 
         DisabledPicture =   "frmOPDRegistration.frx":08CA
         Height          =   615
         Left            =   3240
         Picture         =   "frmOPDRegistration.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   1440
         TabIndex        =   8
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
         Left            =   1440
         TabIndex        =   7
         Top             =   840
         Width           =   3015
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
         ItemData        =   "frmOPDRegistration.frx":1687
         Left            =   1440
         List            =   "frmOPDRegistration.frx":1694
         TabIndex        =   6
         Text            =   "Select Here"
         Top             =   1320
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
         Left            =   4560
         TabIndex        =   5
         Top             =   1320
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
         TabIndex        =   4
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         DisabledPicture =   "frmOPDRegistration.frx":16AD
         Height          =   735
         Left            =   6720
         Picture         =   "frmOPDRegistration.frx":1F77
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3360
         Width           =   975
      End
      Begin VB.ComboBox cmbDoctor 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   3000
         Width           =   5055
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
         TabIndex        =   1
         Top             =   840
         Width           =   1695
      End
      Begin HMS.Duncan_DatePicker dpRegDate 
         Height          =   300
         Left            =   4680
         TabIndex        =   31
         Top             =   3600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         FirstDayOfWeek  =   1
         DescriptionFormat=   "d mmm yyyy"
         UseHandCursor   =   0   'False
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "FEE"
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
         Left            =   1200
         TabIndex        =   29
         Top             =   3960
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "TIMING"
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
         Left            =   960
         TabIndex        =   27
         Top             =   3480
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ON. DATE"
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
         TabIndex        =   25
         Top             =   3600
         Width           =   795
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
         TabIndex        =   24
         Top             =   2520
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
         Left            =   3720
         TabIndex        =   23
         Top             =   2160
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
         Left            =   120
         TabIndex        =   22
         Top             =   2040
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
         Left            =   3720
         TabIndex        =   21
         Top             =   1800
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
         TabIndex        =   20
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
         Left            =   120
         TabIndex        =   19
         Top             =   960
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
         TabIndex        =   18
         Top             =   1320
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
         TabIndex        =   17
         Top             =   1320
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
         TabIndex        =   16
         Top             =   1680
         Width           =   825
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   7800
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DOCTOR NAME"
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
         TabIndex        =   15
         Top             =   3000
         Width           =   1275
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
         TabIndex        =   14
         Top             =   840
         Width           =   765
      End
   End
   Begin VB.Label lblPhysMaint 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OPD Registration"
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
      Left            =   2715
      TabIndex        =   26
      Top             =   240
      Width           =   3435
   End
End
Attribute VB_Name = "frmOPDRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function ValidateFormFields()

If Not ValidateRequiredField(cmbDoctor, "Doctor") Then
        ValidateFormFields = False
        Exit Function
    End If
    
    If Not ValidateRequiredField(txtTiming, "Time") Then
        ValidateFormFields = False
        Exit Function
    End If
    If Not ValidateRequiredField(txtFee, "Fee") Then
        ValidateFormFields = False
        Exit Function
    End If
    
        
    ValidateFormFields = True

End Function
Private Sub cmdSave_Click()
    On Error GoTo ErrHandler
    'On Error Resume Next
    If ValidateFormFields = True Then
        ConnectToDB
        mobjCmd.CommandType = adCmdText
        Dim query As String
        Dim OPDNo As Integer
        Dim strdocID As Variant
        Dim strDID As String
        
        strdocID = Split(cmbDoctor.Text, "-")
        strDID = strdocID(1)
        query = "select *  from OPD_registration where REGNO = '" & txtRegNo.Text & "' and ONDATE = '" & GetFormatedDate(dpRegDate.DateSelected) & "'"
         mobjCmd.CommandText = query
        Set mobjRst = mobjCmd.Execute
        If Not mobjRst.EOF Then
            MsgBox "Patient Exist", vbOKOnly
            DisconnectFromDB
            Exit Sub
        End If
        query = "select max(OPDID) OPDNO from OPD_registration"
         mobjCmd.CommandText = query
        Set mobjRst = mobjCmd.Execute
        If Not mobjRst.EOF Then
            If IsNull(mobjRst!OPDNo) = True Then
                OPDNo = 0
            Else
                OPDNo = CInt(mobjRst!OPDNo)
            End If
        Else
            OPDNo = 0
        
        End If
        OPDNo = OPDNo + 1
        mobjCmd.CommandText = "insert into OPD_registration(OPDID,DID,ONDATE,TIME,FEE,REGNO,ISCOMPLETE,ISCANCEL) values(" & OPDNo & ",'" & strDID & "', '" & GetFormatedDate(dpRegDate.DateSelected) & "', '" & txtTiming.Text & "', " & CDbl(txtFee.Text) & ", '" & txtRegNo.Text & "', 0,0)"
        mobjCmd.Execute
        MsgBox "Saved", vbOKOnly
        DisconnectFromDB
        Call LoadWard
    End If
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


ErrHandler:
    'MsgBox "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
    Exit Sub
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
    query = "select DName  || '-' || DID DName from  Doctor "
    'query = "select bed_no WName from ipd_registration where isdischarge ='0'"
    mobjCmd.CommandText = query
    Set mobjRst = mobjCmd.Execute
   cmbDoctor.Clear
        With mobjRst
        Do Until .EOF
            
            cmbDoctor.AddItem (!DName)
            
            .MoveNext
        Loop
    End With
    
    
    DisconnectFromDB

    Exit Sub
End Sub

