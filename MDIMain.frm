VERSION 5.00
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "(Narayni Emergency Hospital)"
   ClientHeight    =   4740
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10815
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu masterentry 
      Caption         =   "&Master Entry"
      Index           =   0
      Begin VB.Menu bedentry 
         Caption         =   "&Bed Entry"
         Index           =   1
      End
      Begin VB.Menu doctorentry 
         Caption         =   "&Doctor Entry"
         Index           =   2
      End
      Begin VB.Menu serviceentry 
         Caption         =   "&Service Entry"
         Index           =   3
      End
   End
   Begin VB.Menu patientregistration 
      Caption         =   "Patient &Registration"
      Index           =   4
   End
   Begin VB.Menu opd 
      Caption         =   "&OPD"
      Index           =   5
      Begin VB.Menu opdadmission 
         Caption         =   "Admission"
         Index           =   7
      End
      Begin VB.Menu opdview 
         Caption         =   "&View"
         Index           =   8
      End
   End
   Begin VB.Menu ipd 
      Caption         =   "&IPD"
      Index           =   6
      Begin VB.Menu ipdadmission 
         Caption         =   "&Admission"
         Index           =   9
      End
      Begin VB.Menu ipdView 
         Caption         =   "&View"
         Index           =   10
      End
      Begin VB.Menu ipdserviceused 
         Caption         =   "&Service Used"
         Index           =   11
      End
   End
   Begin VB.Menu Account 
      Caption         =   "&Account"
      Index           =   12
      Begin VB.Menu ipdcollection 
         Caption         =   "IPD &Collection"
         Index           =   13
      End
   End
   Begin VB.Menu dcr 
      Caption         =   "Daily &Collection Report"
      Index           =   13
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bedentry_Click(Index As Integer)
frmBedmaster.Show

End Sub

Private Sub dcr_Click(Index As Integer)
frmDailyCollectionReport.Show
End Sub

Private Sub doctorentry_Click(Index As Integer)
frmDoctorMaster.Show
End Sub

Private Sub ipdadmission_Click(Index As Integer)
frmIPDAdmisstion.Show
End Sub

Private Sub ipdcollection_Click(Index As Integer)
frmIPDBillPayment.Show
End Sub

Private Sub ipdserviceused_Click(Index As Integer)
frmIPDServiceUsedByPatient.Show
End Sub

Private Sub ipdView_Click(Index As Integer)
frmvieIPDRegistration.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If MsgBox("Are you sure to exit?", vbYesNo) = vbYes Then
        End
    Else
      Cancel = True
    End If
End Sub

Private Sub opdadmission_Click(Index As Integer)
frmOPDRegistration.Show
End Sub

Private Sub opdview_Click(Index As Integer)
frmViewOPDRegistration.Show
End Sub

Private Sub patientregistration_Click(Index As Integer)
frmPatientRegistration.Show
End Sub

Private Sub serviceentry_Click(Index As Integer)
frmServiceMaster.Show
End Sub
