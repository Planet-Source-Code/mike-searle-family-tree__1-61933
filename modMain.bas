Attribute VB_Name = "modMain"
Option Explicit

Public fMainForm As frmMain

Public gApp As clsApp

Sub Main()
Dim mainCN As ADODB.Connection

    Set gApp = New clsApp

    'Now get the connection to the database
    If GetAccessConnection(mainCN, "Admin", "", App.Path & "\FamTree.mdb", "", "{Microsoft Access Driver (*.mdb)}") Then
        Set gApp.cn = mainCN
    Else
        MsgBox "Failed to establish Connect to the database." & vbCrLf & "The application will now close.", vbCritical, "SQL Connection Failed"
        End
    End If

    frmSplash.Show
    frmSplash.Refresh
    Set fMainForm = New frmMain
    Load fMainForm
    Unload frmSplash

    fMainForm.Show
End Sub

Public Function GetAccessConnection(mainCN As ADODB.Connection, strUser As String, strPassword As String, strDatabase As String, strServer As String, Optional strDriver As String) As Boolean
Dim sConnect As String
    
    GetAccessConnection = False
    On Error GoTo ErrSub

    sConnect = "UID=" + Trim(strUser) + ";PWD=" + Trim(strPassword) _
        + ";DBQ=" + Trim(strDatabase) + ";Driver=" + Trim(strDriver)
    
    Set mainCN = New ADODB.Connection
    mainCN.ConnectionString = sConnect
    mainCN.Open

    GetAccessConnection = True
    
Exit Function
ErrSub:
    'FIXIT - can't use custom MsgBox as no database connection established
    MsgBox "Unable to connect to Access database", vbExclamation + vbOKOnly, _
        "Error Connecting to Database"
    GetAccessConnection = False

End Function



