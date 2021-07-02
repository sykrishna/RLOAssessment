VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUtilsDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'   Class:    clsUtilsDB
'
'   Creator : Dean Lane   3/4/1998
'
'   Overall Function:
'   This common class will provide methods for accessing specific computer and user
'   information as well as an interface to the CML_REGISTRY table.
'   It is to
'
'   Usage:
'   Computer/user specific properties will be avaiable by instantiating this class.
'   To use the 'GetSetting' method (provides access to the computer registry and
'   CML_REGISTRY table) the DSN property must be set.
'
'   Example:
'
'     Private mobjUtilsDB as new clsUtilsDB
'
'     mobjUtilsDB.DSN = INVENTORY_ODBC_4
'
'     sServerName = mobjUtilsDB.GetSetting(0, 0, "ApplicationServer")
'     sFilePath = mobjUtilsDB.GetSetting(SIMApplicationId, STOCKTAKE_PC,"ReportsPath")
'
'   Dependencies:
'   This class depends on the following Modules and Classes being included in the
'   current project.
'
'   Object              Description
'   ---------------     -------------------------------------------------------------
'   clsStdGlobals       Provides common VB/API type and constant declarations
'   modError            Provide error handling facilities for writing to error log
'
'   Change Control
'
'   Name                Date        Description
'   ------------------  ----------  -------------------------------------------------------------
'   Dean Lane           3/4/1998    Clean up and comment ready for others to keep neat (hint hint)
'   Fariba Mokarram     14/4/98     add some functions and fix error handling
'   Dean Lane           15/4/98     Remove function level connections
'   Tonia Pedersen(TLP) 18/01/2000  Add back function level connection, due to too many connections open problem
'   Tonia Pedersen(TLP) 27/01/2000  New procedure CreateEnv: Creates module level environment for use with procedure level connections
'
Option Explicit
Private Const DEFAULT_TIME_OUT = 180  '3Min
Private Const CONNECTION_ENV_NAME = "clsUtils"
Private Const DEFAULT_MAX_RETRIES = 3

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

'Private mconDB              As rdoConnection TLP; 27/01/2000; Not used any more
Private menvDB              As rdoEnvironment

Private msDSN               As String
Private mbConnected         As Boolean
Private msDomainName        As String
Private msUserID            As String
Private msComputerName      As String
'has been added 14/4/98
Private miConflictAppId      As Integer
Private miConflictFuncId     As Integer
Private msConflictCompName   As String
Private mlProcessID          As Long

Public Property Get ConflictAppId() As Integer
    ConflictAppId = miConflictAppId
End Property
Public Property Get ConflictFuncId() As Integer
    ConflictFuncId = miConflictFuncId
End Property
Public Property Get ConflictCompName() As String
    ConflictCompName = msConflictCompName
End Property
Public Property Let DSN(sValue As String)
    msDSN = sValue
End Property
Public Property Get DomainName() As String
    DomainName = msDomainName
End Property
Public Property Get UserID() As String
    UserID = msUserID
End Property
Public Property Get ComputerName() As String
    ComputerName = msComputerName
End Property
Public Property Get DSN() As String
    DSN = msDSN
End Property
Private Sub Class_Initialize()
    
   On Error GoTo Err_Class_Initialize
   
    Dim szComputerName      As String * 100
    Dim szUserName          As String * 100
     
    msDomainName = Environ$("USERDOMAIN")
    msComputerName = Environ$("COMPUTERNAME")
    If ComputerName = "" Then  'Windows 95
       GetComputerName szComputerName, 100
       msComputerName = CStringtoVBString(szComputerName)
    End If
    
    msUserID = Environ$("USERNAME")
    If UserID = "" Then        'Windows 95
        GetUserName szUserName, 100
        msUserID = CStringtoVBString(szUserName)
    End If
    
    msDSN = ""
    mbConnected = False
    
    Exit Sub
    
Err_Class_Initialize:
    LogError "clsUtilsDB", "Class_Initalise", False, False
    err.Raise err.Number
    
End Sub
Public Function RemoveCarraigeReturn(sTxt As String) As String
'Author: Ian Frazer
'Date: 8/6/99
'Purpose: Removes carraige returns from strings (ie Chr(10) and Chr(13)). These are causing problems
'when the string is used in Net Sends

    Dim iPos As Integer
    Dim sResult As String
    
    On Error GoTo Err_RemoveCarraigeReturn
    
    If InStr(1, sTxt, Chr(10)) > 0 Or InStr(1, sTxt, Chr(13)) > 0 Then
        For iPos = 1 To Len(sTxt)
            If Mid$(sTxt, iPos, 1) <> Chr(10) And Mid$(sTxt, iPos, 1) <> Chr(13) Then
                sResult = sResult & Mid$(sTxt, iPos, 1)
            End If
        Next
        RemoveCarraigeReturn = sResult
    Else
        RemoveCarraigeReturn = sTxt
    End If
    
    Exit Function
    
Err_RemoveCarraigeReturn:
    LogError "modSISUtils", "RemoveCarraigeReturn Parm:" & sTxt, False
    err.Raise err.Number

End Function
Public Sub KillApplication(sAppName As String)
'Author: Ian Frazer 25/2/2000
'Purpose: Kills an application on the server
'Arguments: sAppName - The name of the application in the setting_name field of the
'           cml_registry table eg to kill MaxCount.exe sAppName should be CountEXE.
'=================================================================================
' WHO  : Alex S
' WHAT : Mod
' WHY  : To Close Connection
' WHEN : June 2000
'==================================================================================

    
    
    On Error GoTo err_KillApplication
    
    Dim sExe As String
    Dim con As clsConnect ' AS was NEW
    
    
    Set con = New clsConnect
    
    sExe = GetSetting(SIMApplicationId, STOCKTAKE_BATCH_PROCESSING_PC, sAppName)
    If Len(sExe) = 0 Then 'Setting not in registry
        WriteToErrorLog "clsUtilsDB, KillApplication, The setting " & sAppName & " is not in the cml_registry table"
    Else
        con.OpenConnection (msDSN)
        con.Execute "EXEC master.dbo.xp_cmdshell 'Kill " & sExe & "'"
    End If
      
      
    con.CloseConnection   ' AS, June 2000
    Set con = Nothing     ' AS, June 2000
      
    Exit Sub
    
err_KillApplication:
    LogError "clsUtilsDB", "Class_Initalise", False, False
    If Not con Is Nothing Then   ' AS , Next three lines
        Set con = Nothing
    End If
    err.Raise err.Number
    
End Sub
Public Sub RaiseAlert(ctlMsg As Control, lAppId As Long, iErrorNbr As Integer, iMsgId As Integer, _
                                                    Optional sMsgArg1 As String, _
                                                    Optional sMsgArg2 As String, _
                                                    Optional sMsgArg3 As String, _
                                                    Optional sMsgArg4 As String)
                                                    
'Author:    Ian Frazer
'Date:      22/3/2000
'Purpose:   Creates an alert that will be handled by Tivoli or email depending on registry setting and writes
'           details to log file.

    On Error GoTo Err_RaiseAlert

    Dim sMsg                As String
    Dim sCmd                As String
    Dim sAppName            As String
    Static sUseTivoli       As String
    Static sTivoliPath      As String
    Static bTivoliExists    As Boolean
    Dim sServer             As String


    'Determine whether to use Tivoli or email
    If Len(sUseTivoli) = 0 Then
        sUseTivoli = GetSetting(lAppId, 0, "UseTivoli")
    End If

    If sUseTivoli = "1" Then
        If Len(sTivoliPath) = 0 Then
            sServer = GetSetting(0, 0, "ApplicationServer")
            sTivoliPath = "\\" & sServer & GetSetting(lAppId, 0, "TivoliPath")
        End If

        If Not bTivoliExists Then
            'Check that Tivoli executable exists
            If Len(Dir(sTivoliPath & "\opsmsg.exe")) > 0 Then
                bTivoliExists = True
            Else
                WriteToErrorLog "Tivoli executable not found in expected location: " & sTivoliPath
            End If
        End If

        If bTivoliExists Then
            With ctlMsg
                If Len(Trim(sMsgArg1)) = 0 And Len(Trim(sMsgArg2)) = 0 And Len(Trim(sMsgArg3)) = 0 And Len(Trim(sMsgArg4)) = 0 Then
                    .MessageLoad CInt(lAppId), iMsgId
                ElseIf Len(Trim(sMsgArg1)) > 0 And Len(Trim(sMsgArg2)) > 0 And Len(Trim(sMsgArg3)) > 0 And Len(Trim(sMsgArg4)) > 0 Then
                    .MessageLoad CInt(lAppId), iMsgId, sMsgArg1, sMsgArg2, sMsgArg3, sMsgArg4
                ElseIf Len(Trim(sMsgArg1)) > 0 And Len(Trim(sMsgArg2)) > 0 And Len(Trim(sMsgArg3)) > 0 Then
                    .MessageLoad CInt(lAppId), iMsgId, sMsgArg1, sMsgArg2, sMsgArg3
                ElseIf Len(Trim(sMsgArg1)) > 0 And Len(Trim(sMsgArg2)) > 0 Then
                    .MessageLoad CInt(lAppId), iMsgId, sMsgArg1, sMsgArg2
                ElseIf Len(Trim(sMsgArg1)) > 0 Then
                    .MessageLoad CInt(lAppId), iMsgId, sMsgArg1
                End If
                sMsg = .Text
            End With

            'Remove any hard returns in the message string because this could cause problems for Tivoli
            'Ian Frazer 8/6/00 - Lines below commented out because they are VB6 specific. Wont compile in VB5!
            'sMsg = Replace(sMsg, Chr(13), "")
            'sMsg = Replace(sMsg, Chr(10), "")
            sMsg = Me.RemoveCarraigeReturn(sMsg)

            sCmd = sTivoliPath & "\opsmsg -a stocktake -g job " & Chr(34) & Left(iErrorNbr & ", " & sMsg, 255) & Chr(34)
            ExecAndWait sCmd, 0, True
        End If

    Else
        'Send email - This will occur if sUseTivoli is not equal to 1 or the 'UseTivoli' setting has not been created
        With ctlMsg
            If Len(Trim(sMsgArg1)) = 0 And Len(Trim(sMsgArg2)) = 0 And Len(Trim(sMsgArg3)) = 0 And Len(Trim(sMsgArg4)) = 0 Then
                .Mail CInt(lAppId), iMsgId
            ElseIf Len(Trim(sMsgArg1)) > 0 And Len(Trim(sMsgArg2)) > 0 And Len(Trim(sMsgArg3)) > 0 And Len(Trim(sMsgArg4)) > 0 Then
                .Mail CInt(lAppId), iMsgId, sMsgArg1, sMsgArg2, sMsgArg3, sMsgArg4
            ElseIf Len(Trim(sMsgArg1)) > 0 And Len(Trim(sMsgArg2)) > 0 And Len(Trim(sMsgArg3)) > 0 Then
                .Mail CInt(lAppId), iMsgId, sMsgArg1, sMsgArg2, sMsgArg3
            ElseIf Len(Trim(sMsgArg1)) > 0 And Len(Trim(sMsgArg2)) > 0 Then
                .Mail CInt(lAppId), iMsgId, sMsgArg1, sMsgArg2
            ElseIf Len(Trim(sMsgArg1)) > 0 Then
                .Mail CInt(lAppId), iMsgId, sMsgArg1
            End If
        End With
    End If

    Exit Sub

Err_RaiseAlert:
    LogError "clsUtilsDB", "RaiseAlert", False, False

End Sub
Public Function GetSetting(lAppId As Long, lFuncId As Long, sSettingName As String) As Variant
On Error GoTo Err_GetSetting
            
Dim conDB           As rdoConnection
Dim vRegString      As Variant
Dim sSQL            As String
Dim rsRegistry      As rdoResultset
Dim iRetry          As Integer
        
    'TLP; 18/01/2000; Use procedure local Connect to the database.
    If Not mbConnected Then
        CreateEnv
    End If
    Set conDB = menvDB.OpenConnection(msDSN, rdDriverNoPrompt, True, "")
    
    vRegString = RegistryQueryValue("SOFTWARE\CML\" & Trim(Str(lAppId)) & "\" & Trim(Str(lFuncId)), sSettingName)
    If Len(vRegString) > 0 Then
        GetSetting = CStringtoVBString(CStr(vRegString))
        GoTo Exit_GetSetting
    End If
    
    vRegString = RegistryQueryValue("SOFTWARE\CML\" & Trim(Str(lAppId)), sSettingName)
    If Len(vRegString) > 0 Then
        GetSetting = CStringtoVBString(CStr(vRegString))
        GoTo Exit_GetSetting
    End If
    
    vRegString = RegistryQueryValue("SOFTWARE\CML", sSettingName)
    If Len(vRegString) > 0 Then
        GetSetting = CStringtoVBString(CStr(vRegString))
        GoTo Exit_GetSetting
    End If
'    If Not ConnectToDB() Then
'        WriteToErrorLog "clsUtilsDB: GetSetting: Connection to database not available"
'        GetSetting = ""
'        Exit Function
'    End If
    
    sSQL = " SELECT SETTING_VALUE " & _
           " FROM CML_REGISTRY " & _
           " WHERE APPL_ID = " & lAppId & " AND " & _
           " FUNC_ID = " & lFuncId & " AND " & _
           " SETTING_NAME = '" & sSettingName & "'"
    Set rsRegistry = conDB.OpenResultSet(sSQL, rdOpenStatic, rdConcurReadOnly)
    If Not rsRegistry.EOF Then
        GetSetting = "" & rsRegistry!setting_value
        rsRegistry.Close
        GoTo Exit_GetSetting
    End If
    rsRegistry.Close
    
    sSQL = " SELECT SETTING_VALUE " & _
           " FROM CML_REGISTRY " & _
           " WHERE APPL_ID = " & lAppId & " AND " & _
           " FUNC_ID Is Null AND " & _
           " SETTING_NAME = '" & sSettingName & "'"
    Set rsRegistry = conDB.OpenResultSet(sSQL, rdOpenStatic, rdConcurReadOnly)
    If Not rsRegistry.EOF Then
        GetSetting = "" & rsRegistry!setting_value
        rsRegistry.Close
        GoTo Exit_GetSetting
    End If
    rsRegistry.Close
    
    sSQL = " SELECT SETTING_VALUE " & _
           " FROM CML_REGISTRY " & _
           " WHERE APPL_ID Is Null AND " & _
           " FUNC_ID Is Null AND " & _
           " SETTING_NAME = '" & sSettingName & "'"
    Set rsRegistry = conDB.OpenResultSet(sSQL, rdOpenStatic, rdConcurReadOnly)
    If Not rsRegistry.EOF Then
        GetSetting = "" & rsRegistry!setting_value
        rsRegistry.Close
        GoTo Exit_GetSetting
    End If
    rsRegistry.Close
    
    GetSetting = ""
                           
Exit_GetSetting:
Dim lError  As Long
    lError = err.Number
On Error Resume Next

    'TLP; 18/01/2000; Close connections
    conDB.Close
    
    Set conDB = Nothing
    
    'Put Raise here after clean up
    If lError <> 0 Then
        err.Raise lError
    End If
    Exit Function
    
Err_GetSetting:
    If err.Number = 40002 Then
        If iRetry <= 20 Then
            'Jo Gault / Ian Frazer 30 Nov 1999
            'This is to fix a strange intermittant bug which seems to occur
            'if the resultset is opened too quickly after the connection is
            'initiated.  When the error occurs we wait a bit & then try again.
            'After 20 tries, we give up.
            SleepEx 100, 0
            '*** Alex S, added jan 2000 to Log conn properties
            WriteToErrorLog "DROPS CONNECTION : " & Time & " : " & conDB.Connect
            LogError "clsUtilsDB", "GetSetting - Retry: " & iRetry, False, False
            iRetry = iRetry + 1
            Resume
        End If
    End If
    LogError "clsUtilsDB", "GetSetting", False, False
    
    GoTo Exit_GetSetting
    
End Function
Public Sub SetSetting(lAppId As Long, lFuncId As Long, sSetting As String, sValue As String)
On Error GoTo Err_SetSetting

Dim conDB           As rdoConnection
Dim sSQL            As String
Dim sWhereClause    As String
Dim rs              As rdoResultset


    'TLP; 18/01/2000; Use procedure local Connect to the database.
    If Not mbConnected Then
        CreateEnv
    End If
    Set conDB = menvDB.OpenConnection(msDSN, rdDriverNoPrompt, True, "")
    
    sWhereClause = "WHERE  SETTING_NAME = '" & sSetting & "' AND " & _
        "       APPL_ID = " & lAppId & " AND " & _
        "       FUNC_ID = " & lFuncId

    Set rs = conDB.OpenResultSet("SELECT SETTING_VALUE FROM CML_REGISTRY " & sWhereClause)
    If rs.EOF And rs.BOF Then

        sSQL = "INSERT INTO CML_REGISTRY (APPL_ID, FUNC_ID, STORE_SPECIFIC, SETTING_NAME, SETTING_VALUE) " & _
                    " VALUES  (" & lAppId & ", " & _
                    lFuncId & ", " & _
                    "'N'," & _
                    "'" & sSetting & "', " & _
                    "'" & sValue & "')"
    Else
        sSQL = "UPDATE CML_REGISTRY SET SETTING_VALUE = '" & sValue & "' " & sWhereClause
    End If

    conDB.Execute sSQL
    
Exit_SetSetting:
Dim lError  As Long
    lError = err.Number
On Error Resume Next
    
    'Close connection
    conDB.Close
    
    Set conDB = Nothing
    
    'Put Raise here after clean up
    If lError <> 0 Then
        err.Raise lError
    End If
    Exit Sub
    
Err_SetSetting:
    LogError "clsUtilsDB", "SetSetting", False, False
    GoTo Exit_SetSetting
    
End Sub

'TLP; 27/01/2000; Not used any more - tried to fix multi connection problem
'Private Function ConnectToDB() As Boolean
'
'    On Error GoTo Err_ConnectToDB
'
'    Dim sConName    As String
'    Dim iEnvCount   As Integer
'
'    If mbConnected Then
'       ConnectToDB = True
'       Exit Function
'    End If
'
'    If msDSN = "" Then
'       WriteToErrorLog "clsUtilsDB: ConnectToDB: DSN property has not been set - unable to connect to database."
'       ConnectToDB = False
'       Exit Function
'    End If
'
'    iEnvCount = rdoEngine.rdoEnvironments.Count
'    Do
'        iEnvCount = iEnvCount + 1
'        sConName = "clsUtilsDB" & Trim$(CStr(iEnvCount))
'        On Error Resume Next
'        Set menvDB = rdoEngine.rdoCreateEnvironment(sConName, rdoDefaultUser, rdoDefaultPassword)
'        If err.Number = 0 Then
'            Exit Do
'        End If
'    Loop While True
'
'    On Error GoTo Err_ConnectToDB
'
'
'
'
'    Set mconDB = menvDB.OpenConnection(msDSN, rdDriverNoPrompt, True, "")
'
'    'This don't work
'    '**** Alex S, Nov 99
'    '**** We should reshake connection, otherwise err 40002
'    '**** It is just workaround
'
'
'    'mconDB.Close
'    'mconDB.EstablishConnection rdDriverNoPrompt
'
'   ' Set mconDB = menvDB.OpenConnection(msDSN, rdDriverNoPrompt, True, "")
'
'    '******************
'
'
'
'
'    mbConnected = True
'    ConnectToDB = True
'
'    Exit Function
'
'Err_ConnectToDB:
'    LogError "clsUtilsDB", "ConnectToDB", False, False
'    err.Raise err.Number
'
'End Function

Public Function RegistryQueryValue(sKeyName As String, sValueName As String) As Variant
    '  Get Registry Values
    '
    '  eg. sKeyName := "SOFTWARE/CML/SMS/1.0" sValueName := "LogFilePath"
    On Error GoTo Err_RegistryQueryValue
    
    Dim lRetVal         As Long
    Dim lKey            As Long
    Dim vValue          As Variant
    
    lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyName, 0, KEY_QUERY_VALUE, lKey)
    If lRetVal <> ERROR_NONE Then
        RegistryQueryValue = Empty
        Exit Function
    End If
    lRetVal = QueryValueEx(lKey, sValueName, vValue)
    RegCloseKey (lKey)
    
    If Len(vValue) > 0 Then
        vValue = CStringtoVBString(CStr(vValue))
    End If
    
    RegistryQueryValue = vValue
    
    Exit Function
    
Err_RegistryQueryValue:
    LogError "clsUtilsDB", "RegistryQueryValue", False, False
    err.Raise err.Number
    
End Function

Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
' Support function for RegistryQueryValue
    On Error GoTo Err_QueryValueEx
    
    Dim cch         As Long
    Dim lrc         As Long
    Dim lType       As Long
    Dim lValue      As Long
    Dim sValue      As String
    
    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then
       vValue = Empty
       Exit Function
    End If
    
    Select Case lType
        Case REG_SZ:    ' For strings
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch)
            Else
                vValue = Empty
            End If
    
        Case REG_DWORD: ' For DWORDS
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
            If lrc = ERROR_NONE Then vValue = lValue
        Case Else
            'all other data types not supported
            lrc = -1
    End Select
    
    QueryValueEx = lrc
    
    Exit Function
    
Err_QueryValueEx:
    LogError "clsUtilsDB", "QueryValueEx", False, False
    err.Raise err.Number
    
End Function

Private Sub Class_Terminate()

    On Error Resume Next
    
    'TLP; 27/01/2000; Not used any more
'    If mbConnected Then
'        mconDB.Close
'    End If
    If mbConnected Then
        menvDB.Close
        Set menvDB = Nothing
    End If
    
End Sub

Public Function ConflictingFunctions(appl_id As Integer, func_id As Integer) As Boolean
On Error GoTo Err_ConflictingFunctions

Dim conDB       As rdoConnection
Dim rs          As rdoResultset
Dim SQL         As String

    miConflictAppId = 0
    miConflictFuncId = 0
    msConflictCompName = ""

    'TLP; 18/01/2000; Use procedure local Connect to the database.
    If Not mbConnected Then
        CreateEnv
    End If
    Set conDB = menvDB.OpenConnection(msDSN, rdDriverNoPrompt, True, "")
'    If Not ConnectToDB() Then
'        WriteToErrorLog "clsUtilsDB: ConflictingFunctions: Connection to database not available"
'        ConflictingFunctions = True
'        Exit Function
'    End If
    ' setup, and query the SQL table for conflicting functions and
    ' active functions matching the one specified
    
    SQL = " select cf.appl_id, cf.func_id, cf.conflicting_appl_id, " & _
          "    cf.conflicting_func_id, af.appl_id af_appl_id, af.func_id af_func_id, " & _
    "     af.computer_name af_computer_name" & _
     "  from CML_Conflicting_Functions cf, CML_Active_Functions af " & _
     "  where cf.appl_id = " & appl_id & " and cf.func_id = " & func_id & " and " & _
    "     cf.conflicting_appl_id = af.appl_id and " & _
     "     cf.conflicting_func_id = af.func_id"
    Set rs = conDB.OpenResultSet(SQL, rdOpenStatic, rdConcurReadOnly)

    If rs.EOF Then    ' no rows found - no conflict
        ConflictingFunctions = False
    Else
        'Save the first conflicting application and function id's
        ConflictingFunctions = True
        miConflictAppId = NullConvert(rs!af_appl_id, 0)
        miConflictFuncId = NullConvert(rs!af_func_id, 0)
        msConflictCompName = NullConvert(rs!af_computer_name, "")
    End If
    rs.Close
    
Exit_ConflictingFunctions:
Dim lError  As Long
    lError = err.Number
On Error Resume Next
    
    'Close connection
    conDB.Close
    
    Set conDB = Nothing
    
    'Raise Error here
    If lError <> 0 Then
        err.Raise lError
    End If
    Exit Function

Err_ConflictingFunctions:
    LogError "clsUtilsDB", "ConflictingFunctions", False, False
    GoTo Exit_ConflictingFunctions

End Function

Public Function Register(iAppId As Integer, iFuncid As Integer) As Boolean
    'Register the function to the cml_active_functions table if there are no
    'conflicting functions
    'Would like Convert to use a stored procedure but used by too many
    'other apps & don't know who would require the new sp
On Error GoTo Err_Register

Const MAX_EMP_ID = 12

Dim conDB       As rdoConnection
Dim sSQL        As String
        
    'TLP; 18/01/2000; Create procedure local connection
    If Not mbConnected Then
        CreateEnv
    End If
    Set conDB = menvDB.OpenConnection(msDSN, rdDriverNoPrompt, True, "")
    
    If ConflictingFunctions(iAppId, iFuncid) Then
       Register = False
       GoTo Exit_Register
    Else
       Register = True
    End If

    mlProcessID = GetCurrentProcessId
    
'    conDB.Execute ("EXEC usp_CML_RegisterApp '" & msComputerName & "', '" & _
'            Left$(msUserID, MAX_EMP_ID) & "', " & _
'            mlProcessID & ", " & _
'            iAppId & ", " & _
'            iFuncid & ", '" & _
'            App.EXEName & "'")
        

    sSQL = " DELETE FROM CML_ACTIVE_FUNCTIONS " & _
           " WHERE COMPUTER_NAME  = '" & msComputerName & "'" & _
           " AND EMP_ID <> '" & Left$(msUserID, MAX_EMP_ID) & "'"
    conDB.Execute sSQL

    sSQL = " delete from cml_active_functions " & _
           " where computer_name = '" & msComputerName & "' and " & _
           " process_id = " & mlProcessID
    conDB.Execute sSQL

    sSQL = "insert into CML_Active_Functions " & _
           "(appl_id, func_id, emp_id, computer_name, " & _
           " start_time, process_id, image_name) " & _
           " values (" & iAppId & "," & iFuncid & ", '" & Left$(msUserID, MAX_EMP_ID) & "', '" & _
          msComputerName & "','" & DateTimeStringFormat(Now) & "'," & mlProcessID & ",' " & App.EXEName & "')"
    conDB.Execute sSQL

Exit_Register:
Dim lError  As Long
    lError = err.Number
On Error Resume Next
    
    'Close connection
    conDB.Close
    
    Set conDB = Nothing
    
    'Raise error here
    If lError <> 0 Then
        err.Raise lError
    End If
    Exit Function

Err_Register:
    LogError "clsUtilsDB", "Register", False, False
    GoTo Exit_Register
    
End Function

Public Sub Unregister()
On Error GoTo Err_Unregister

Dim conDB   As rdoConnection
Dim sSQL    As String
    
    
    'TLP; 18/01/2000; Use local connection
    If Not mbConnected Then
        CreateEnv
    End If
    Set conDB = menvDB.OpenConnection(msDSN, rdDriverNoPrompt, True, "")
'    If Not ConnectToDB() Then
'        WriteToErrorLog "clsUtilsDB: Unregister: Connection to database not available"
'        Exit Sub
'    End If
    
    sSQL = " delete from CML_Active_Functions " & _
           " where Computer_Name = '" & msComputerName & "' and Process_ID = " & mlProcessID
    conDB.Execute sSQL
    
Exit_Unregister:
Dim lError  As Long
    lError = err.Number
On Error Resume Next
    
    'Close connection
    conDB.Close
    
    Set conDB = Nothing
    
    'Raise error here
    If lError <> 0 Then
        err.Raise lError
    End If
    Exit Sub

Err_Unregister:
    LogError "clsUtilsDB", "UnRegister", False, False
    GoTo Exit_Unregister

End Sub

Public Function UserPermission(iApplId As Integer, iFuncid As Integer)
    ' UserPermission
    ' Verify that the user has permission to run the function
    '
    ' Returns:
    '   True : user has permission
    '   False : user does not have permission
    Dim sApplicationServer  As String
    Dim sPath               As String
    Dim fhInput             As Integer
    
    On Error GoTo Err_UserPermission
    
    ' Get Application Server name from the registry
    sApplicationServer = GetSetting(0, 0, "ApplicationServer")
    '
    ' If the user has permission to read the function's file then
    ' they can run the function.
    '
    fhInput = FreeFile
    sPath = "\\" & sApplicationServer & "\SCS$\Perm\" & Format(iApplId, "0000") & "\" & Format(iFuncid, "0000")
            
    On Error GoTo AccessDenied
    
    Open sPath For Input As #fhInput
    Close #fhInput

    UserPermission = True

    Exit Function

AccessDenied:
    UserPermission = False
    Exit Function

Err_UserPermission:
    LogError "clsUtilsDB", "UserPermission", False, False
    err.Raise err.Number

End Function
Public Function sGetFuncDescr(iAppId As Integer, iFuncid As Integer) As String
'Created:   Dean Lane 8/1/1998  Given application and function id return the description
On Error GoTo Err_sGetFuncDescr

Dim conDB           As rdoConnection
Dim strConName      As String
Dim rsFuncDesr      As rdoResultset
Dim sSQL            As String
    
    'TLP; 18/01/2000; Use local connection
    If Not mbConnected Then
        CreateEnv
    End If
    Set conDB = menvDB.OpenConnection(msDSN, rdDriverNoPrompt, True, "")
'    If Not ConnectToDB() Then
'        WriteToErrorLog "clsUtilsDB: sGetFuncDescr: Connection to database not available"
'        sGetFuncDescr = ""
'        Exit Function
'    End If
    
    sGetFuncDescr = ""
    sSQL = " Select Func_Descr From APPL_FUNCTION " & _
           " Where Appl_Id = " & iAppId & " and Func_Id = " & iFuncid
    Set rsFuncDesr = conDB.OpenResultSet(sSQL, rdOpenStatic, rdConcurReadOnly)
    If Not rsFuncDesr.EOF Then
        sGetFuncDescr = NullConvert(rsFuncDesr!Func_Descr, "")
    Else
        sGetFuncDescr = ""
    End If
    rsFuncDesr.Close
    
Exit_sGetFundDescr:
Dim lError  As Long
    lError = err.Number
On Error Resume Next
    
    'Close connection
    conDB.Close
    
    Set conDB = Nothing
    
    'Raise error here
    If lError <> 0 Then
        err.Raise lError
    End If
    Exit Function

Err_sGetFuncDescr:
    LogError "clsUtilsDB", "sGetFuncDescr", False, False
    GoTo Exit_sGetFundDescr
    
End Function

Public Function FunctionIsActive(appl_id As Integer, func_id As Integer) As Boolean
On Error GoTo Err_FunctionIsActive
        
Dim conDB       As rdoConnection
Dim rsView      As rdoResultset
Dim SQL         As String
    
    'TLP; 18/01/2000; Use local connection
    If Not mbConnected Then
        CreateEnv
    End If
    Set conDB = menvDB.OpenConnection(msDSN, rdDriverNoPrompt, True, "")
'    If Not ConnectToDB() Then
'        WriteToErrorLog "clsUtilsDB: FunctionIsActive: Connection to database not available"
'        FunctionIsActive = True
'        Exit Function
'    End If
    
    SQL = "select appl_id, func_id from CML_Active_Functions where appl_id = " _
        & appl_id & " and func_id = " & func_id
    Set rsView = conDB.OpenResultSet(SQL, rdOpenStatic, rdConcurReadOnly)

    If rsView.EOF Then
        FunctionIsActive = False
    Else
        FunctionIsActive = True
    End If
    rsView.Close
    
Exit_FunctionIsActive:
Dim lError  As Long
    lError = err.Number
On Error Resume Next
    
    'Close connection
    conDB.Close
    
    Set conDB = Nothing
    
    'Raise error here
    If lError <> 0 Then
        err.Raise lError
    End If
    Exit Function

Err_FunctionIsActive:
    LogError "clsUtilsDB", "FunctionIsActive", False, False
    GoTo Exit_FunctionIsActive

End Function

Public Function FunctionIsActiveLocal(appl_id As Integer, func_id As Integer) As Boolean
On Error GoTo Err_FunctionIsActiveLocal

Dim conDB       As rdoConnection
Dim rsView      As rdoResultset
Dim SQL         As String


    'TLP; 18/01/2000; Use local connectiopn
    If Not mbConnected Then
        CreateEnv
    End If
    Set conDB = menvDB.OpenConnection(msDSN, rdDriverNoPrompt, True, "")
    
    SQL = "select appl_id, func_id from CML_Active_Functions where appl_id = " & appl_id & _
        " and func_id = " & func_id & _
        " and computer_name = '" & msComputerName & "'"
    Set rsView = conDB.OpenResultSet(SQL, rdOpenKeyset, rdConcurRowVer)
    If rsView.EOF Then
        FunctionIsActiveLocal = False
    Else
        FunctionIsActiveLocal = True
    End If
    rsView.Close
    
Exit_FunctionIsActiveLocal:
Dim lError  As Long
    lError = err.n
On Error Resume Next
    
    'Close connection
    conDB.Close
    
    Set conDB = Nothing
    
    'Raise error here
    If lError <> 0 Then
        err.Raise lError
    End If
    Exit Function

Err_FunctionIsActiveLocal:
    LogError "clsUtilsDB", "FunctionIsActiveLocal", False, False
    GoTo Exit_FunctionIsActiveLocal
    
End Function

'TLP; 27/01/2000; Creates module level environment for use with procedure level connections
Private Sub CreateEnv()
On Error GoTo Err_CreateEnv
    
Dim iEnvCount   As Integer
Dim sConName    As String
    
    iEnvCount = rdoEngine.rdoEnvironments.Count
    Do
        iEnvCount = iEnvCount + 1
        sConName = CONNECTION_ENV_NAME & Trim$(CStr(iEnvCount))
        On Error Resume Next
        Set menvDB = rdoEngine.rdoCreateEnvironment(sConName, rdoDefaultUser, rdoDefaultPassword)
        If err.Number = 0 Then
            Exit Do
        End If
    Loop While True
        
    mbConnected = True
        
Exit Sub
    
Err_CreateEnv:
    mbConnected = False
    LogError "clsUtilsDB", "CreateEnv", False
    err.Raise err.Number
    
End Sub

