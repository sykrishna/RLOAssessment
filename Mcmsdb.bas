Attribute VB_Name = "mDBBackOffice"
' mDBBackOffice
'
'
' Uses ErrorLog functionality - be sure to include 'modError.bas' module in your project.
'
'
' Modified by: Bob Byrne 06/06/1997 to provide the following functions:
'               OpenSIM0001DSN01 - returns a rdoConnection after opening DSN=SIMDSN01.
'               CloseConnection - closes supplied rdoConnection.
'              Bob Byrne 07/07/1997 - OpenConnection('DSN CONSTANT') function created
'
Option Explicit

Private iCnt As Integer     'may not require these two
Dim con As rdoConnection    'declarations if using gConnection below

Public gConnection As rdoConnection

'Function used by many SIM apps to expand drop down list box
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' The following declare the WinExec function allowing execution of apps from VB
Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Public Function OpenSISDSNIPM() As rdoConnection
'
' Returns connection if successful in opening the connection
'  otherwise 'nothing'
'
'   ---- Allows Connection pooling
'

Dim cn As rdoConnection

On Error GoTo ErrorHandler
    
    Set cn = New rdoConnection
    
    cn.CursorDriver = rdUseOdbc
    cn.Connect = "dsn=DKSLS01"
    cn.EstablishConnection rdDriverNoPrompt
        
    Set OpenSISDSNIPM = cn
        
    Set cn = Nothing
        
    Exit Function
ErrorHandler:
    LogError "mDBBackOffice", "OpenSISDSNIPM", False
    err.Raise err.Number, err.Source, err.Description
    
End Function
Public Sub CloseCMSDB()
'this corresponds to bCMSDBOpen() function

    On Error Resume Next
    gConnection.Close
    Set gConnection = Nothing
    
End Sub

Public Function nRowCount(rs As rdoResultset) As Long
'returns correct RowCount in a resultset
    Dim lngPosition As Long
    
    nRowCount = 0
    
    'check if there are any record in the resultset
    If Not (rs.BOF And rs.EOF) Then
        'save the current position in the resultset
        lngPosition = rs.AbsolutePosition
    
        rs.MoveLast
        nRowCount = rs.RowCount
    
        'reposition the resultset
        rs.AbsolutePosition = lngPosition
    End If
    
End Function


Public Function CMSDBOpen() As rdoConnection
    ' Note that I have not used prepared statements here because I need to
    ' test for IS NULL in some cases and = 9999 in others.
    '
    If iCnt = 0 Then
        Set con = rdoEnvironments(0).OpenConnection("", rdDriverNoPrompt, False, "dsn=sim0001dsn01;uid=;pwd=;database=cmsdb")
    End If

    iCnt = iCnt + 1
    Set CMSDBOpen = con

End Function

Public Function sPadString(ByVal sString As String, intPaddedStrLen As Integer, _
                           Optional sPadChar As String = " ", _
                           Optional bLeftPadding As Boolean = True)
'I have placed this function here as I don't know where to put
'will be better placed in some common/standard module or class
'returns a string padded with the desired character

    Dim sTmp As String
    Dim intLen As Long
    
    On Error GoTo Error_PS
    
    sTmp = String(intPaddedStrLen, sPadChar)
    intLen = Len(sString)
    
    If intLen > 0 Then
        If bLeftPadding Then
            Mid(sTmp, intPaddedStrLen - intLen + 1) = sString
        Else
            Mid(sTmp, 1, intPaddedStrLen) = sString
        End If
    End If
    
    sPadString = sTmp
    
Exit_PS:
    Exit Function
    
Error_PS:
    MsgBox Error
    Resume Exit_PS
    
End Function


Public Function CMSDBClose()

    iCnt = iCnt - 1

    If iCnt = 0 Then
      con.Close
    End If

End Function

Public Function bCMSDBOpen() As Boolean
'returns true if successful in opening the connection
'otherwise false

    On Error GoTo Error_CDO
    
    bCMSDBOpen = False
    
'BB    Set gConnection = rdoEnvironments(0).OpenConnection("", rdDriverNoPrompt, False, "dsn=simdsn01;uid=;pwd=;database=cmsdb")
    Set gConnection = rdoEnvironments(0).OpenConnection("sim0001dsn01", rdDriverNoPrompt, False)
    
    bCMSDBOpen = True

Exit_CDO:
    Exit Function

Error_CDO:
    Resume Exit_CDO

End Function


Public Function OpenSIM0001DSN01() As rdoConnection
' Returns connection if successful in opening the connection
'  otherwise 'nothing'
'
' Done this way so that more than 1 connection may be opened.
'  It is assumed that the mormal call will be:
'           set gConnection = OpenSIM0001DSN01()
'   thus setting the global variable defined in this module.
'
On Error GoTo Error_OpenSIM0001DSN01
    
    Set OpenSIM0001DSN01 = rdoEnvironments(0).OpenConnection(INVENTORY_ODBC_1, rdDriverNoPrompt, False)
    'Set OpenSIM0001DSN01 = rdoEnvironments(0).OpenConnection("", rdDriverNoPrompt, False, "dsn=simdsn01;uid=;pwd=;database=cmsdb")
    
    Exit Function
Error_OpenSIM0001DSN01:
    LogError "mDBBackOffice", "OpenSIM0001DSN01", False
    err.Raise err.Number
End Function

Public Function OpenConnection(sODBCDSN As String) As rdoConnection
' Returns connection if successful in opening the connection
'  otherwise 'nothing'
'
' Done this way so that more than 1 connection may be opened.
'  It is assumed that the mormal call will be:
'           set gConnection = OpenConnection()
'   thus setting the global variable defined in this module.
'
On Error GoTo Error_OpenConnection

    Set OpenConnection = rdoEnvironments(0).OpenConnection(sODBCDSN, rdDriverNoPrompt, False)
    
    Exit Function
Error_OpenConnection:
    LogError "mDBBackOffice", "OpenConnection", False
    err.Raise err.Number
End Function

Public Sub CloseConnection(conClosing As rdoConnection)
On Error GoTo Error_CloseConnection

    conClosing.Close
    Set conClosing = Nothing
    
    Exit Sub
Error_CloseConnection:
    LogError "mDBBackOffice", "CloseConnection", False
    err.Raise err.Number
End Sub

' ****************************************************************************
'
' *********** To be moved to 'MAXio' ****************************************
'
' ****************************************************************************
Public Function debug_log(s As String)

    Dim iFnum As Integer
    Static sFilename As String
    Static sOpenDone As String

    iFnum = FreeFile
    If sOpenDone <> "Y" Then
        sFilename = App.Path & "\debug.log"
    End If

    If sOpenDone = "Y" Then
       Open sFilename For Append As #iFnum
    Else
       Open sFilename For Output As #iFnum
       sOpenDone = "Y"
    End If

    Print #iFnum, Time, " ", s
    Close #iFnum

End Function

Public Function OpenIPSDSN() As rdoConnection
'
' Returns connection if successful in opening the connection
'  otherwise 'nothing'
'
'   ---- Allows Connection pooling
'

Dim cn As rdoConnection

On Error GoTo ErrorHandler
    
    Set cn = New rdoConnection
    
    cn.CursorDriver = rdUseOdbc
    cn.Connect = "dsn=IPSDSN"
    cn.EstablishConnection rdDriverNoPrompt
        
    Set OpenIPSDSN = cn
        
    Set cn = Nothing
    
    bIPSOpen = True
        
    Exit Function
ErrorHandler:
    bIPSOpen = False
    LogError "mDBBackOffice", "OpenIPSDSN", False
    err.Raise err.Number, err.Source, err.Description
    
End Function
