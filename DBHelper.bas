Attribute VB_Name = "DBHelper"
Option Explicit

Public mobjConn                As ADODB.Connection
Public mobjCmd                 As ADODB.Command
Public mobjRst                 As ADODB.Recordset
Public strConnectionString As String


'*****************************************************************************
'*               Programmer-Defined Subs & Functions                         *
'*****************************************************************************

'-----------------------------------------------------------------------------
Public Sub ConnectToDB()
'-----------------------------------------------------------------------------

    Set mobjConn = New ADODB.Connection
    'mobjConn.ConnectionString = "Provider=MSDAORA.1;" _
                              & "Data Source=OracleServiceHMS;" _
                              & "User ID=scott;" _
                              & "Password=tiger" _
                              & "Persist Security Info=True"
                              
                              
 'mobjConn.ConnectionString = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=localhost)(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=DB1)));User Id=scott;Password=tiger"
'mobjConn.ConnectionString = "Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True"
'mobjConn.ConnectionString = "HOST =localhost; Data Source=ORADB;Password=tiger;User ID=scott; Integrated Security = no" 'Persist Security Info=True"
    
    
    
    'mobjConn.Provider = "OraOLEDB.Oracle"
    'mobjConn.Properties("Data Source") = "ORADB"
    'mobjConn.Properties("User Id") = "scott"
    'mobjConn.Properties("Password") = "tiger"
    
    
    'mobjConn.ConnectionString = "Driver={Oracle in XE}; User Id=system;Password=pass"
    'mobjConn.ConnectionString = "Driver={Oracle in XE}; User Id=pragya;Password=pragya"
    mobjConn.ConnectionString = strConnectionString
    mobjConn.Open

    Set mobjCmd = New ADODB.Command
    Set mobjCmd.ActiveConnection = mobjConn
    
End Sub


'------------------------------------------------------------------------
Public Sub ClearCommandParameters()
'------------------------------------------------------------------------

    Dim lngX    As Long
    
    For lngX = (mobjCmd.Parameters.Count - 1) To 0 Step -1
        mobjCmd.Parameters.Delete lngX
    Next

End Sub

'-----------------------------------------------------------------------------
Public Sub DisconnectFromDB()
'-----------------------------------------------------------------------------

    Set mobjCmd = Nothing
    
    mobjConn.Close
    Set mobjConn = Nothing

End Sub


Public Function AuthenticateUser(username As String, password As String) As Boolean
    ConnectToDB
    ClearCommandParameters
    mobjCmd.CommandType = adCmdText
        
        'mobjCmd.CommandText = "select UserName, Passwords from Login where username  = @UserName and passwords = @Password"
        'mobjCmd.Parameters.Append mobjCmd.CreateParameter("@UserName", adVarChar, adParamInput, 50, username)
        'mobjCmd.Parameters.Append mobjCmd.CreateParameter("@Password", adVarChar, adParamInput, 100, password)
        Dim query As String
        query = "select UserName, Passwords from Login where username  = '" & username & "' and passwords = '" & password & "'"
        mobjCmd.CommandText = query
        Set mobjRst = mobjCmd.Execute
        If mobjRst.EOF Then
    'If mobjRst.RecordCount = 0 Then
    AuthenticateUser = False
    Else
        AuthenticateUser = True
        
    End If
    DisconnectFromDB

End Function



Public Function SetBedMaster(BedID As String, WardName As String, WardType As String, HOD As String, Charge As Double, Status As String) As Boolean
    ConnectToDB
    ClearCommandParameters
    mobjCmd.CommandType = adCmdText
    Dim query As String
    query = "insert into bed_record (BED_ID, WARD_NAME, WARD_TYPE,HOD,CHARGE,STATUS) values ('" & BedID & "' , '" & WardName & "' , '" & WardType & "' , '" & HOD & "' , " & Charge & " , '" & Status & "')"
    mobjCmd.CommandText = query
    mobjCmd.Execute
    SetBedMaster = True
    DisconnectFromDB

End Function


Public Function GetBedMaster() As ADODB.Recordset
    ConnectToDB
    ClearCommandParameters
    mobjCmd.CommandType = adCmdText
    Dim query As String
    query = "select * from bed_record"
    mobjCmd.CommandText = query
   Set mobjRst = mobjCmd.Execute
   Set GetBedMaster = mobjRst
    DisconnectFromDB

End Function


