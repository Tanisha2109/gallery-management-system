Attribute VB_Name = "mainModule"
Sub Main()
    On Error GoTo ErrHandler
    ' assign connection string
    Dim FileName As String
    FileName = App.Path & "\config.txt"
    Open FileName For Input As #1
    Contents = Input(LOF(1), #1)
    Close #1
    strConnectionString = Contents
    'connect to db
    ConnectToDB
    'disconnect to db
    DisconnectFromDB
    frmSplash.Show
   
    Exit Sub
ErrHandler:
    MsgBox "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
    
    End Sub
