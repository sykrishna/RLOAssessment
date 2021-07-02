VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' This class contains routines which are used by all VB programs.
'
'
' $History: clsStd.cls $
'
'*****************  Version 4  *****************
'User: Kchan        Date: 6/06/05    Time: 15:54
'Updated in $/Store Inventory System [SIS]/Software/Source/SAMIM/PDE_MIGRATION/PDE Migration - Inc 2/MaxGroupCapture
'Roll Back Changes
'
'*****************  Version 2  *****************
'User: Kchan        Date: 29/03/04   Time: 11:22
'Updated in $/Store Inventory System [SIS]/Software/Source/SAMIM/PDE_MIGRATION/PDE Migration - Inc 2/MaxGroupCapture
'Enhance clsSTD.GetStoreID logic
'
'*****************  Version 51  *****************
'User: Tgiannon     Date: 1/03/01    Time: 11:07
'Updated in $/Applications/Data Reporting System [DRS]/Software/Planning/Source/SPS0015Reps
'Converted back to VB5.
'
'*****************  Version 50  *****************
'User: Ntrigg       Date: 23/10/00   Time: 15:23
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/PDE_MIGRATION/DataMigration/DMSyncRequest
'For DataMigration Kmart 001.001.060 build compilation
'
'*****************  Version 49  *****************
'User: Ntrigg       Date: 23/10/00   Time: 15:15
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/PDE_MIGRATION/DataMigration/DMSyncRequest
'For DataMigration Kmart 001.001.060 build compilation
'
'*****************  Version 48  *****************
'User: Ntrigg       Date: 22/10/00   Time: 15:57
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/PDE_MIGRATION/DataMigration/DMCleanProcess
'version 2.2.60 - check in and out to get latest version
'
'*****************  Version 47  *****************
'User: Sjabbour     Date: 18/07/00   Time: 14:56
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/waste/Waste
'Fixef speed and GST quantity calculation
'
'*****************  Version 46  *****************
'User: Sjabbour     Date: 18/07/00   Time: 14:21
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/waste/MaxWaste
'Fixed speed and GST calculation
'
'*****************  Version 45  *****************
'User: Pdudley      Date: 1/06/00    Time: 13:25
'Updated in $/Applications/Store Merchandise System [SMS]/Software/Current/SMSFindOCX/Source
'Added conditional compile for SMSFINDI
'
'*****************  Version 42  *****************
'User: Cpayne       Date: 1/05/00    Time: 10:25
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/GeneralCapture
'Version 6
'
'*****************  Version 40  *****************
'User: Kng          Date: 14/04/00   Time: 10:29
'Updated in $/Applications/Store Merchandise System [SMS]/Software/LabelEngine/Source
'Fixed Label Print problem due to SQL page blocking cause by large
'transaction batch (Commits every 100). Workaround by reducing commit
'batch to 1
'
'*****************  Version 39  *****************
'User: Kng          Date: 5/04/00    Time: 14:26
'Updated in $/Applications/Store Merchandise System [SMS]/Software/HostMaintenanceApply(HMApply)/Source
'
'*****************  Version 38  *****************
'User: Mxatkins     Date: 7/03/00    Time: 10:33
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/imdc0000 SIS Batch/SISBatch
'
'*****************  Version 37  *****************
'User: Mxatkins     Date: 6/03/00    Time: 17:11
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/imdc0000 SIS Batch/SISBatch
'
'*****************  Version 36  *****************
'User: Mxatkins     Date: 6/03/00    Time: 17:06
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/imdc0000 SIS Batch/SISBatch
'
'*****************  Version 35  *****************
'User: Mxatkins     Date: 21/02/00   Time: 14:13
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/imdc0000 SIS Batch
'
'*****************  Version 34  *****************
'User: Mxatkins     Date: 7/02/00    Time: 19:13
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/imdc0000 SIS Batch
'
'*****************  Version 33  *****************
'User: Mxatkins     Date: 7/02/00    Time: 19:08
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/imdc0000 SIS Batch
'
'*****************  Version 33  *****************
'User: MXAtkins     Date: 7/02/00    Time: 16:18
'Added process Sleep Function

'*****************  Version 32  *****************
'User: Rcuthill     Date: 7/02/00    Time: 16:18
'Updated in $/Applications/Store Merchandise System [SMS]/Software/PLUDailySales/Source/PLU Daily Sales
'
'*****************  Version 30  *****************
'User: Rcuthill     Date: 28/01/00   Time: 15:48
'Updated in $/Applications/Store Merchandise System [SMS]/Software/MerchPLU/Software
'Change to utilise Coles Myer naming conventions
'
'*****************  Version 29  *****************
'User: Atea         Date: 28/01/00   Time: 10:57
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/PDE_MIGRATION/TransferToOS2
'removing "clsstd." from clsstd.UnexpectedErrorFatal ... This is
'unneccessary and creates dependency on global clsstd....
'
'*****************  Version 28  *****************
'User: Rcuthill     Date: 13/01/00   Time: 8:32
'Updated in $/Applications/Store Merchandise System [SMS]/Software/MerchPLU/Software
'
'*****************  Version 27  *****************
'User: Rcuthill     Date: 7/01/00    Time: 10:29
'Updated in $/Applications/Store Merchandise System [SMS]/Software/PLUDailySales/Source/PLU Daily Sales
'
'*****************  Version 26  *****************
'User: Rcuthill     Date: 31/12/99   Time: 13:49
'Updated in $/Applications/Store Merchandise System [SMS]/Software/PLUDailySales/Source/PLU Daily Sales
'
'*****************  Version 25  *****************
'User: Rcuthill     Date: 29/11/99   Time: 13:58
'Updated in $/Applications/Store Merchandise System [SMS]/Software/ProfitRetentionSystem[PRS]/Source
'GST Amendments
'
'*****************  Version 24  *****************
'User: Kng          Date: 23/11/99   Time: 15:37
'Updated in $/Applications/Store Merchandise System [SMS]/Software/Range/Source
'
'*****************  Version 23  *****************
'User: Kng          Date: 17/11/99   Time: 16:55
'Updated in $/Applications/Store Merchandise System [SMS]/Software/ScheduleApply/Source
'
'*****************  Version 22  *****************
'User: Joliver      Date: 13 10 99   Time: 16:42
'Updated in $/Applications/Online Shopping [OLS]/Software/NSE/CCFSend
'1. Modified ExecAndWait() return type from Integer to Long.
'
'*****************  Version 21  *****************
'User: Akorneyk     Date: 11/10/99   Time: 15:23
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/CMLStocktake/MaxStkTk
'
'*****************  Version 20  *****************
'User: Kng          Date: 26/08/99   Time: 11:58
'Updated in $/STORE MERCHANDISE SYSTEMS/src/In Store Maintenance/DissipationBusinessObject/CurrentBuild
'
'*****************  Version 19  *****************
'User: Trogan       Date: 4-12-98    Time: 9:06
'Updated in $/Store Merchandise Systems/src/Unknown Waste Percentages Table Load
'Remove ShowGUI from Logerror calls
'
'*****************  Version 17  *****************
'User: Bbyrne       Date: 31-08-98   Time: 13:42
'Updated in $/Store Merchandise Systems/src/Utilities/SQLCComp
'New function 'RunningAsEXE'
'
'*****************  Version 16  *****************
'User: Joliver      Date: 16/07/98   Time: 10:05
'Updated in $/Store Administration Systems/DRS/src/DocketReport
'clsStd.cls has effectively been rolled back to Verison 8.
'This is due to concerns by other members of the NSE team
'regarding changes made to the getSetting() function.
'As a result, I have replaced the previous version (Version 15)
'with this version.
'
'   John Oliver     16 July, 1998
'
'*****************  Version 8  *****************
'User: Jatabak      Date: 5/06/98    Time: 11:20
'Updated in $/Store Administration Systems/FDM/src/Balance
' STOREID property Added.
'
'*****************  Version 7  *****************
'User: Joliver      Date: 18/05/98   Time: 11:25
'Updated in $/Store Administration Systems/DRS/src/EFTReconciliation
'UnexpectedErrorFatal now raises error to caller for FIN apps.
'
'*****************  Version 6  *****************
'User: Pbryson      Date: 1-05-98    Time: 10:10
'Updated in $/Store Planning Systems/src/General  Reports/SPSREP
'Fixed a problem with the images directory path name.
'Added a description to the operator column i.e. 'Operator & Id'
'
'*****************  Version 5  *****************
'User: Bbyrne       Date: 15-04-98   Time: 8:06
'Updated in $/Store Merchandise Systems/src/Host Maintenence/Apply/HMApply
'Fix bug - 'SettingNotFound' was not reset upon entry to the
''GetSetting' call.
'
'*****************  Version 4  *****************
'User: Walbury      Date: 4-03-98    Time: 3:39p
'Updated in $/Store Inventory Systems/SAMIM/imdc0450
'
'*****************  Version 3  *****************
'User: Lromagna     Date: 1-28-98    Time: 3:38p
'Updated in $/Store Administration Systems/FDM/src/Balance
'
'*****************  Version 2  *****************
'User: Lromagna     Date: 12/22/97   Time: 9:21a
'Updated in $/Store Administration Systems/FDM/src/Batch
'
'*****************  Version 1  *****************
'User: Awatson      Date: 12/01/97   Time: 10:57a
'Created in $/Store Systems Shared Area/NSE/src/Classes
'New Birth for ClsStd.cls
'

Option Explicit

Private msDomainName As String
Private msUserID As String
Private miStoreID As Integer
Private msComputerName As String
Private mbSettingNotFound As Boolean  ' Used by GetSetting
Private mbShowGUI As Boolean

' Following API call used to determine running environment.
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal HMODULE As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

' Halt the current process for n Milliseconds
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub ProcessSleep(lMilliSecs As Long)
Attribute ProcessSleep.VB_Description = " Halt the current process for n Milliseconds"
' Halt the current process for n Milliseconds
    Sleep lMilliSecs
End Sub

Public Sub Tile(Background As Object, Tile As Object, StartY As Long)
' Author  : Daniel Siew (11/3/97)
' Last Modified : 11/3/97
'
' Description :
' This function tiles a backgound with a given image.
'
' Usage :
' Tile me, image1, 0
' Use this in the Form Load sub module.
'
' Note:
' The Form's AutoRedraw Property must be set to True.
' There exists a memory leak in this function due to the "pallet"
On Error GoTo errHandle

    Dim i As Integer
    Dim j As Integer
    Dim TwipsX As Integer
    
    TwipsX = Screen.TwipsPerPixelX
    
    For i = 0 To Background.ScaleWidth Step Tile.Width - TwipsX
        For j = StartY To Background.ScaleHeight Step Tile.Height
        
            Background.PaintPicture Tile.Picture, i, j, Tile.Width, Tile.Height
        Next j
    Next i
    
    Exit Sub
    
errHandle: UnexpectedErrorFatal "clsStd", "Tile", err.Source, err.Number, err.Description
End Sub

Public Function GetSetting(appl_id As Integer, func_id As Integer, sName As String) As Variant

    '
    ' Routine to return the value of a given setting from either the local
    ' PC's registry or the database CML_Registry table.
    '
    ' Returns Null if the setting was not found.
    '
    ' Modified :
    '    D.Siew 10/6/97 : Striped off null character from registry setting value.
    '    B.Byrne 15/4 98: Reset 'SettingNotFound' flag upon entry.
    '
    On Error GoTo errHandle
    
    Dim vRegString As Variant
    
    ' Ensure error flag is reset upon each call.                BB 15/4/98
    SettingNotFound = False
    
    '
    ' Try going through the local PC's registry first.
    '
    vRegString = RegistryQueryValue("SOFTWARE\CML\" & Trim(Str(appl_id)) & "\" & Trim(Str(func_id)), sName)
    If Len(vRegString) > 0 Then
        GetSetting = CStringtoVBString(CStr(vRegString))
        Exit Function
    End If
    vRegString = RegistryQueryValue("SOFTWARE\CML\" & Trim(Str(appl_id)), sName)
    If Len(vRegString) > 0 Then
        GetSetting = CStringtoVBString(CStr(vRegString))
        Exit Function
    End If
    vRegString = RegistryQueryValue("SOFTWARE\CML", sName)
    If Len(vRegString) > 0 Then
        GetSetting = CStringtoVBString(CStr(vRegString))
        Exit Function
    End If
    
    '
    ' No luck in the local PC's registry so try the CML_Registry database
    ' table.
    '
    Dim con As rdoConnection
    Dim rs As rdoResultset
    
    '
    ' Note that I have not used prepared statements here because I need to
    ' test for IS NULL in some cases and = 9999 in others.
    '
    Set con = rdoEnvironments(0).OpenConnection("SCS0001DSN01", rdDriverNoPrompt, False, "")
    ' Function level
    Set rs = con.OpenResultset("select setting_value from CML_registry where appl_id = " & appl_id & " and func_id = " & func_id & " and setting_name = '" & sName & "'", rdOpenKeyset, rdConcurRowVer)
    If (rs.EOF And rs.BOF) Then
        rs.Close
        ' Application level
        Set rs = con.OpenResultset("select setting_value from CML_registry where appl_id = " & appl_id & " and func_id is null and setting_name = '" & sName & "'", rdOpenKeyset, rdConcurRowVer)
        If (rs.EOF And rs.BOF) Then
            rs.Close
            ' Site level
            Set rs = con.OpenResultset("select setting_value from CML_registry where appl_id is null and func_id is null and setting_name = '" & sName & "'", rdOpenKeyset, rdConcurRowVer)
        End If
    End If
        
    If (rs.EOF And rs.BOF) Then
        SettingNotFound = True
    Else
        SettingNotFound = False
        GetSetting = rs!setting_value
    End If
    
    rs.Close
    con.Close
        
    Exit Function
    
errHandle: UnexpectedErrorFatal "clsStd", "GetSetting", err.Source, err.Number, err.Description
    
End Function

Public Sub UnexpectedErrorWarning(sModule As String, sRoutine As String, sSource As String, iNumber As Long, sDescription As String)
    
    '
    ' Routine which gets called when we get an unexpected error in a
    ' routine and we wish to warn the user but not terminate the program.
    '
    On Error GoTo errHandle
    Dim msg As String
    msg = "An unexpected error occurred, please contact your system support help desk." & vbCrLf
    msg = msg & vbCrLf
    msg = msg & "Module: " & sModule & vbCrLf
    msg = msg & "Routine: " & sRoutine & vbCrLf
    msg = msg & vbCrLf
    msg = msg & "Error Source: " & sSource & vbCrLf
    msg = msg & "Error Number: " & iNumber & vbCrLf
    msg = msg & "Error Description: " & sDescription & vbCrLf
    
    If mbShowGUI Then
        MsgBox msg, vbExclamation + vbApplicationModal
    End If
    
    Exit Sub
    
errHandle: UnexpectedErrorFatal "clsStd", "UnexpectedErrorWarning", err.Source, err.Number, err.Description
    
End Sub

Public Sub UnexpectedErrorFatal(sModule As String, sRoutine As String, sSource As String, iNumber As Long, sDescription As String)
' Routine which gets called when we get an unexpected error in a
' routine and we wish to terminate the program.
On Error Resume Next
    Dim msg As String
    
    
    ' ALL FIN programs handle errors by raising the error back to the caller.
    ' The caller then deals with the error appropriately.
    #If FIN Then
        Call err.Raise(iNumber, sSource, sDescription, "", 0)
        Exit Sub
    #End If
    
    AbortProgram
    
    msg = "An unexpected error occurred, please contact your system support help desk." & vbCrLf
    msg = msg & vbCrLf
    msg = msg & "Module: " & sModule & vbCrLf
    msg = msg & "Routine: " & sRoutine & vbCrLf
    msg = msg & vbCrLf
    msg = msg & "Error Source: " & sSource & vbCrLf
    msg = msg & "Error Number: " & iNumber & vbCrLf
    msg = msg & "Error Description: " & sDescription & vbCrLf
    
    '
    ' Show message box if ShowGUI
    '
    If mbShowGUI Then
        MsgBox msg, vbCritical + vbApplicationModal
        MsgBox "Due to the previous error this program will now abnormally terminate.", vbCritical + vbApplicationModal
    End If
#If SMSFINDI Then
#Else
    End
#End If

End Sub

Public Function CStringtoVBString(s As String)
    
    '
    ' Do some weird stuff to get a null terminated string into a normal string
    '
    
    On Error GoTo errHandle
    Dim i As Integer
    Dim c As Byte
    i = 1
    CStringtoVBString = ""
    
    While Asc(Mid(s, i, 1)) <> 0
    
        CStringtoVBString = CStringtoVBString & Mid$(s, i, 1)
        i = i + 1
    
    Wend
    
    Exit Function
    
errHandle: UnexpectedErrorFatal "clsStd", "CStringtoVBString", err.Source, err.Number, err.Description
    
End Function

Public Sub FormCenter(frmForm As Form)
    
    '
    ' Class Module to Center a form, Form Name as argument
    '
    ' Move form once for optimization.
    ' Syntax: FORM.Move Left, Top
    
    On Error GoTo errHandle
    frmForm.Move (Screen.Width - frmForm.Width) / 2, (Screen.Height - frmForm.Height) / 2
    Exit Sub
    
errHandle: UnexpectedErrorFatal "clsStd", "FormCenter", err.Source, err.Number, err.Description
    
End Sub

Public Sub FormPrint(frmForm As Form)
    
    '
    ' Class Module to print the screen onto the printer
    '
    
    On Error GoTo errHandle
    frmForm.PrintForm
    Exit Sub
    
errHandle: UnexpectedErrorFatal "clsStd", "FormPrint", err.Source, err.Number, err.Description
    
End Sub

Public Function NullConvert(field, Optional NullVal, Optional NotNullVal)
    
    '
    ' Class Module to test if a field is null and allow user to
    ' assign it a different value if is or not...
    '
    On Error GoTo errHandle
    
    If IsNull(field) Then
    
        If IsMissing(NullVal) Then
        
            NullConvert = ""
        
        Else
        
            NullConvert = NullVal
        
        End If
        
    Else
    
        If IsMissing(NotNullVal) Then
        
            NullConvert = field
        
        Else
        
            NullConvert = NotNullVal
        
        End If
        
    End If
    
    Exit Function
    
errHandle: UnexpectedErrorFatal "clsStd", "NullConvert", err.Source, err.Number, err.Description
    
End Function

Public Sub ControlSetReadOnly(obj As Object)
    
    ' Class Module to set a contorl to a read only state
    On Error GoTo errHandle
    
    If TypeOf obj Is TextBox Then
    
        obj.Locked = True
        obj.BackColor = vbMenuBar
    
    ElseIf TypeOf obj Is ComboBox Then
    
        ' For ComboBoxes you need to have set up a TextBox of the
        ' same size and position with the ComboBox. The name of the
        ' TextBox should be the same as the ComboBox except that it
        ' must have a txt prefix rather than a cbo prefix.
        Dim ctl As Control
        Dim bNotSet As Boolean
        
        bNotSet = True
        
        For Each ctl In Screen.ActiveForm.Controls
        
            If bNotSet Then
            
                If "txt" & Right(obj.Name, Len(obj.Name) - 3) = ctl.Name Then
                
                    ctl = obj
                    obj.Visible = False
                    ctl.Visible = True
                    ctl.Locked = True
                    ctl.BackColor = vbMenuBar
                    bNotSet = False
                
                End If
                
            Else
            
                Exit Sub
            
            End If
            
        Next
        
    End If
    
    Exit Sub
    
errHandle: UnexpectedErrorFatal "clsStd", "ControlSetReadOnly", err.Source, err.Number, err.Description
    
End Sub

Public Sub ControlSetUpdatable(obj As Object)
    
    ' Class Module to Change a control to be in an updatable state.
    '
    On Error GoTo errHandle
    
    If TypeOf obj Is TextBox Then
    
        obj.Locked = False
        obj.BackColor = vbWhite
    
    ElseIf TypeOf obj Is ComboBox Then
    
        '
        ' For ComboBoxes we just have to make them visible
        ' and the associated TextBox invisible.
        Dim ctl As Control
        Dim bNotSet As Boolean
        
        bNotSet = True
        
        For Each ctl In Screen.ActiveForm.Controls
        
            If bNotSet Then
            
                If "txt" & Right(obj.Name, Len(obj.Name) - 3) = ctl.Name Then
                
                    obj.Visible = True
                    ctl.Visible = False
                    bNotSet = False
                
                End If
                
            Else
            
                Exit Sub
            
            End If
            
        Next
        
    End If
    
    Exit Sub
    
errHandle: UnexpectedErrorFatal "clsStd", "ControlSetUpdatable", err.Source, err.Number, err.Description
    
End Sub

Public Sub ControlFill(Control As ComboBox, Connection As rdoConnection, SQL As String, Column As String, ClearControl As Boolean, Optional FormatStyle)
    
    ' Fill a combo box control with a column of data from a select on a table.
    '
    On Error GoTo errHandle
    Dim rs As rdoResultset
    
    If ClearControl Then
    
        Control.Clear
    
    End If
    
    Set rs = Connection.OpenResultset(SQL, rdOpenKeyset)
    
    rs.MoveLast
    rs.MoveFirst
    
    Do Until rs.EOF
    
        If IsMissing(FormatStyle) Then
        
            Control.AddItem rs(Column)
        
        Else
        
            Control.AddItem Format(rs(Column), FormatStyle)
        
        End If
        
        rs.MoveNext
    
    Loop
    
    rs.Close
    
    Exit Sub
    
errHandle: UnexpectedErrorFatal "clsStd", "ControlFill", err.Source, err.Number, err.Description
    
End Sub

Public Function RegistryQueryValue(sKeyName As String, sValueName As String) As Variant
    
    '
    '  Get Registry Values
    '
    '  eg. sKeyName := "SOFTWARE/CML/SMS/1.0" sValueName := "LogFilePath"
    '
    On Error GoTo errHandle
    Dim lRetVal As Long
    Dim hKey As Long
    Dim vValue As Variant
    
    lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyName, 0, KEY_QUERY_VALUE, hKey)
    If lRetVal <> ERROR_NONE Then
        RegistryQueryValue = Empty
        Exit Function
    End If
    lRetVal = QueryValueEx(hKey, sValueName, vValue)
    RegCloseKey (hKey)
    RegistryQueryValue = vValue
    
    Exit Function
    
errHandle: UnexpectedErrorFatal "clsStd", "RegistryQueryValue", err.Source, err.Number, err.Description
    
End Function

Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
    
    '
    ' Support function for RegistryQueryValue
    '
    
    On Error GoTo errHandle
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String
    
    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then
       vValue = Empty
       Exit Function
    End If
    
    Select Case lType
    ' For strings
    
        Case REG_SZ:
    
        sValue = String(cch, 0)
        lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
        
        If lrc = ERROR_NONE Then
        
            vValue = Left$(sValue, cch)
        
        Else
        
            vValue = Empty
        
        End If
        
        ' For DWORDS
    
        Case REG_DWORD:
    
        lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
        If lrc = ERROR_NONE Then vValue = lValue
    
    Case Else
    
        'all other data types not supported
        lrc = -1
    
    End Select
    
    QueryValueEx = lrc
    
    Exit Function
    
errHandle: UnexpectedErrorFatal "clsStd", "QueryValueEx", err.Source, err.Number, err.Description
    
End Function

Public Sub Create3DBorder(txtObj As TextBox, obj As Control)
    
    '
    ' Fix for the 3D look for ListView Control
    '
    ' Apart from calling this function you need to,
    '      1. Set the ListView Object to a 2D view
    '      2. Create a text box
    '      3. Remember to put the text box in the same form and ,if applicable,
    '         the same frame as the ListView Control.
    '      4. Call the function when the form is loaded.
    '         clsstd.Create3DBorder frmMain.text1, frmMain.lvwMain
    '
    On Error GoTo errHandle
    txtObj.Move obj.Left - 30, _
    obj.Top - 30, _
    obj.Width + 60, _
    obj.Height + 60
    obj.ZOrder 0
    Exit Sub
    
errHandle: UnexpectedErrorFatal "clsStd", "Create3DBorder", err.Source, err.Number, err.Description
    
End Sub

Public Function ExecAndWait(CmdLine As String, secTimeout As Integer, _
            Optional xPos As Long = -1, Optional yPos As Long = -1) As Long
' Executes the command line and waits until it has finished
'
' commmand line is a complete command and timeout is in seconds, x and y are in pixels
'
' Returns the exit code of the function
'
' if Xpos or ypos is set then the startup position is used.
On Error GoTo errHandle
    
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim lRet As Long
    Dim lRetCode As Long
    Dim lTimeout As Long
    Dim tc As Long
    
    ' Initialize time out
    lTimeout = CLng(secTimeout) * 10
    
' Initialize the STARTUPINFO structure:
    start.cb = Len(start)
    If xPos <> -1 Or yPos <> -1 Then ' use the startup position specified
        start.dwX = xPos
        start.dwY = yPos
        start.dwFlags = start.dwFlags Or STARTF_USEPOSITION
    End If
    
' Start the shelled application:
    lRet = CreateProcessA(0&, CmdLine, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    
' Wait for the shelled application to finish:
    If lRet Then
        tc = 0
        
        Do While lRet And (tc < lTimeout Or lTimeout = 0) ' repeat till time out
            lRet = WaitForSingleObject(proc.hProcess, 100)
            DoEvents
            tc = tc + 1
        Loop
    
        lRet = GetExitCodeProcess(proc.hProcess, lRetCode)
        lRet = TerminateProcess(proc.hProcess, 0&)
        lRet = CloseHandle(proc.hProcess)
        lRet = CloseHandle(proc.hThread)
    
        If tc >= lTimeout And lTimeout > 0 Then ' we timed out!
            ExecAndWait = -1 'Fail
        Else
            ExecAndWait = lRetCode              'Return with Return Code
        End If
    Else
            ExecAndWait = -1  'Fail
    End If
    
Exit Function
errHandle:
    err.Raise err.Number, err.Source & " WinExecutionFunctions.ExecAndWait", err.Description, err.HelpFile, err.HelpContext
End Function

Private Sub Class_Initialize()
    
    On Error GoTo errHandle
    
    DomainName = Environ$("USERDOMAIN")
    
    ComputerName = Environ$("COMPUTERNAME")
    If ComputerName = "" Then   ' Windows 95
        Dim szComputerName As String * 100
        GetComputerName szComputerName, 100
        ComputerName = CStringtoVBString(szComputerName)
    End If
    
    UserID = Environ$("USERNAME")
    If UserID = "" Then     ' Windows 95
        Dim szUserName As String * 100
        GetUserName szUserName, 100
        UserID = CStringtoVBString(szUserName)
    End If
    
    '
    ' Default to show GUI
    '
    ShowGUI = True
    
    Exit Sub
    
errHandle: UnexpectedErrorFatal "clsStd", "Class_Initialize", err.Source, err.Number, err.Description
    
End Sub

Public Property Get DomainName()
    
    DomainName = msDomainName
    
End Property

Private Property Let DomainName(vNewValue)
    
    msDomainName = vNewValue
    
End Property

Public Property Get UserID()
    
    UserID = msUserID
    
End Property

Private Property Let UserID(vNewValue)
    
    msUserID = vNewValue
    
End Property

Public Property Get ComputerName()
    
    ComputerName = msComputerName
    
End Property

Private Property Let ComputerName(vNewValue)
    
    msComputerName = vNewValue
    
End Property

Public Property Get SettingNotFound()

    SettingNotFound = mbSettingNotFound

End Property

Private Property Let SettingNotFound(vNewValue)

    mbSettingNotFound = vNewValue
    
End Property

Public Sub FormsUnload()
    
    '
    ' Unload all the forms in the program
    '
    
    Dim i As Integer
    i = Forms.Count
    Do While i > 0
        Unload Forms(i - 1)
        If i = Forms.Count Then Exit Do
        i = i - 1
    Loop

End Sub

Public Function FindInCollection(cln As Variant, sKey As String) As Boolean
'
' Search for an item in the collection by it's key.
' NOTE : Keys in collections have to be a string.
'
On Error GoTo errHandle
    
    FindInCollection = False
    If cln.Count > 0 Then
        If Not IsNull(cln(sKey)) Then
            FindInCollection = True
        End If
    End If
    
    Exit Function

errHandle:
'
' ATTENTION : LEAVE THIS BLANK. THIS MEANS THAT IT IS NOT (THAT RIGHT, NOT!) IN THE COLLECTION !
'
End Function

Public Sub OpqBackStyle(CmdButton As Control)
' Author  : John Atabak (22/05/97)
' Last Modified : 22/05/97
' Description :
'       This function changes the back style of passed button to Opaque style .
'
' Usage : sscommandButton_MouseEnter event
'       Use this in the Form Level for every command buttons.
    
    On Error GoTo errHandle
    
    CmdButton.BackStyle = 0  ' ssOpaque
    'since some applications do not use sheridan controls,
    'the actual value of "ssOpaque" is used for backstyle property
    CmdButton.BackColor = 14671839
  
    Exit Sub

errHandle: UnexpectedErrorFatal "clsstd", "OpqBackStyle", err.Source, err.Number, err.Description
End Sub

Public Sub TransBackStyle(CmdButton As Control)
' Author  : John Atabak (22/05/97)
' Last Modified : 22/05/97
' Description :
'       This function changes the back style of passed button to transparent style .
' Usage : sscommandButton_MouseExit
'       Use this in the Form Level for every command buttons.
   
    On Error GoTo errHandle
    
    CmdButton.BackStyle = 1 'ssTransparent
    'since some applications do not use sheridan controls,
    'the actual value of "ssTransparent" is used
    
    Exit Sub

errHandle: UnexpectedErrorFatal "clsStd", "TransBackStyle", err.Source, err.Number, err.Description
End Sub

Public Sub ClearCollection(PassedCollection As Collection)
    
' Author  : Giorgia Stasis (22/05/97)
' Last Modified : 22/05/97
' Description :
' This function removes all members of the passed collection
' Usage : Use this in the Form Level for clearing a collection
    
    On Error GoTo errHandle
    
    Dim iNum As Integer
    
    For iNum = 1 To PassedCollection.Count
         PassedCollection.Remove 1
    Next iNum
    
    Exit Sub
    
errHandle: UnexpectedErrorFatal "clsStd", "ClearCollection", err.Source, err.Number, err.Description
End Sub

Public Function ExtractString(sRecord As String, sFieldDelimiter As String, iField As Integer, bRight As Boolean) As String
' Author  : John Atabak (24/06/97)
' Last Modified : 24/06/97
' Description :
'   This function returns partial of the passed text, according to the number of fieldDelimeter appearances
'   if bRight is true  then it returns the whole text after  ifield th appearance of the field Delimeter
'   if bRight is false then it returns the text from ifield-1  appearance of the field Delimeter upto ifield-1  appearance of the field Delimeter
'   if ifield is out of rage or delimeter not found the last part of the text is returned
' Usage : Use for getting the part of the text
 
 On Error GoTo errHandle
    
    Dim sRemainderRecord As String
    Dim i As Integer
    
    ExtractString = ""
    sRemainderRecord = Trim(sRecord)

  For i = 2 To iField
        If InStr(sRemainderRecord, sFieldDelimiter) = 0 And i <= iField Then
            sRemainderRecord = ""
        Else
            sRemainderRecord = Mid(sRemainderRecord, InStr(sRemainderRecord, sFieldDelimiter) + 1)
        End If
    Next i
      'where code gets here,ExtractString starts from the "iField"th instant of delimeter
      ' e.g if sRecord=" text1-text2-text3-text4" and  FieldDelimiter is "-" and field is "2"
      'then ExtractString is "text2-text3-text4"
    If InStr(sRemainderRecord, sFieldDelimiter) = 0 Then
        'FieldDelimiter not found so return the remain text
           If sRemainderRecord <> "" Then
                If bRight = True Then
                    sRemainderRecord = ""
                End If
            End If
    ElseIf bRight = False Then
        ' find next appearance of the delimeter and return the left part of the text
        sRemainderRecord = Left(sRemainderRecord, InStr(sRemainderRecord, sFieldDelimiter) - 1)
    Else
        ' find next appearance of the delimeter and return the right part of the text upto the end
        sRemainderRecord = Mid(sRemainderRecord, InStr(sRemainderRecord, sFieldDelimiter) + 1, Len(sRemainderRecord))
    End If
    ExtractString = Trim(sRemainderRecord)
    
    Exit Function
errHandle: UnexpectedErrorFatal "clsStd", "ExtractString", err.Source, err.Number, err.Description
End Function

Public Function LoadImage(iStoreId As Integer, iApplId As Integer, sImageCode As String, Optional lReturnImageMaskColor As Long = -1) As Variant
' Author : Daniel Siew 7/97
' Last Modified : Daniel Siew 23/7/97
' Description:
' This is the function to load an image for an application
' from the path specified in the table, optionally returns the mask color
' by the mask color parameter.  Returns the picture
' Usage :
'   set control.picture = LoadImage(StoreId,ApplicationId,ImageCode[,MaskColor])
On Error GoTo errHandle
    Dim con As rdoConnection
    Dim qd As rdoQuery
    Dim rs As rdoResultset
    Dim sSQL As String
    Dim sPath As String
    Dim sFileName As String
    Dim sExtension As String
    Dim sFullFileName As String
    Dim bFound As Boolean
    Dim iOriginalStoreId As Integer
    
    Set LoadImage = Nothing
    
    '
    ' Set a few variables
    '
    bFound = False
    iOriginalStoreId = iStoreId
    
    
    '
    ' Open Connection
    '
    Set con = rdoEnvironments(0).OpenConnection("SCS0001DSN01", rdDriverNoPrompt, False, "")

    If con Is Nothing Then
        Exit Function
    End If

    '
    ' Search at which level
    '

    '
    ' Check Store and Application Specific
    '
    If iStoreId <> 0 And iApplId <> 0 Then
        sSQL = "select img_filename, img_ext, mask_color from cml_images " _
             & "where str_id = ? and appl_id = ? and img_cd = ?"
        Set qd = con.CreateQuery("", sSQL)
        qd(0) = iStoreId
        qd(1) = iApplId
        qd(2) = sImageCode
        Set rs = qd.OpenResultset()
        If Not (rs.EOF And rs.BOF) Then
            bFound = True
        Else
            iStoreId = 0
            rs.Close
            qd.Close
        End If
     End If
    
    '
    ' Check Application Wide
    '
    If iStoreId = 0 And iApplId <> 0 And Not bFound Then
        sSQL = "select img_filename, img_ext, mask_color from cml_images " _
             & "where str_id is null and appl_id = ? and img_cd = ?"
        Set qd = con.CreateQuery("", sSQL)
        qd(0) = iApplId
        qd(1) = sImageCode
        Set rs = qd.OpenResultset()
        If Not (rs.EOF And rs.BOF) Then
            bFound = True
        Else
            iApplId = 0
            rs.Close
            qd.Close
        End If
    End If
        
    '
    ' Check Store Wide
    '
    If iOriginalStoreId <> 0 And iApplId = 0 And Not bFound Then
        sSQL = "select img_filename, img_ext, mask_color from cml_images " _
             & "where str_id = ? and appl_id is null and img_cd = ?"
        Set qd = con.CreateQuery("", sSQL)
        qd(0) = iOriginalStoreId
        qd(1) = sImageCode
        Set rs = qd.OpenResultset()
        If Not (rs.EOF And rs.BOF) Then
            bFound = True
        Else
            iApplId = 0
            rs.Close
            qd.Close
        End If
    End If
    
    '
    ' Check Store Wide and Application Wide
    '
    If iStoreId = 0 And iApplId = 0 And Not bFound Then
        sSQL = "select img_filename, img_ext, mask_color from cml_images " _
             & "where str_id is null and appl_id is null and img_cd = ?"
        Set qd = con.CreateQuery("", sSQL)
        qd(0) = sImageCode
        Set rs = qd.OpenResultset()
        If Not (rs.EOF And rs.BOF) Then
            bFound = True
        Else
            rs.Close
            qd.Close
        End If
    End If


    If Not bFound Then
        con.Close
        Exit Function
    End If
    
    If Not (rs.EOF And rs.BOF) Then
        '
        ' Set Filename
        '
        sFullFileName = ""
        sFileName = NullConvert(rs(0).Value, "")
        sExtension = NullConvert(rs(1).Value, "")
        If Len(sFileName) > 0 And Len(sExtension) > 0 Then
            sFullFileName = sFileName & "." & sExtension
            If lReturnImageMaskColor <> -1 Then
                lReturnImageMaskColor = NullConvert(rs(2), 0)
            End If
        End If
    End If
    
    '
    ' Close data connection
    '
    rs.Close
    qd.Close
    con.Close
    
    '
    ' Build Path
    '
    If Len(sFullFileName) > 0 Then
        '
        ' Build Path
        '
        sPath = GetSetting(0, 0, "ApplicationServer")
        If mbSettingNotFound Then
            Exit Function
        End If
        
        sPath = "\\" & sPath & GetSetting(iApplId, 0, "ImagePath")
        If mbSettingNotFound Then
            Exit Function
        End If
    End If
    
    '
    ' Load Picture
    '
    If Dir(sPath & sFullFileName) <> "" Then
       Set LoadImage = LoadPicture(sPath & sFullFileName)
    End If
    
    Exit Function
errHandle: UnexpectedErrorFatal "clsStd", "LoadImage", err.Source, err.Number, err.Description
End Function

Public Sub SetButtonImage(cmd As Control, iStoreId As Integer, iApplId As Integer, sImageCode As String)
' Creator : Daniel Siew 7/97
' Last Modified : Daniel Siew 23/7/97
' Description :
'   This function is a small encapsulation of the LoadImage
'   function that specifically sets the picture property
'   of a sheridan control and uses the mask color.
' Usage:
'   SetButtonImage SScontrol,StoreId,ApplicationId,ImageCode
On Error GoTo errHandle

Dim lMaskColor As Long
    
    Set cmd.Picture = LoadImage(iStoreId, iApplId, sImageCode, lMaskColor)
    If cmd.Picture Is Nothing Then
        cmd.Enabled = True
        Exit Sub
    End If

    cmd.PictureMaskColor = lMaskColor
    cmd.PictureUseMask = True
    
    Exit Sub
    
errHandle: UnexpectedErrorFatal "clsStd", "SetButtonImage", err.Source, err.Number, err.Description
End Sub

Public Property Get ShowGUI() As Boolean
' Creator : Daniel Siew 7/97
' Last Modified : 5/8/97
' Description : Set to TRUE when you want a GUI to be displayed.
'               Set to FALSE when you do not want a GUI to be displayed.
' Usage:
'   ShowGUI = TRU
'
    ShowGUI = mbShowGUI
End Property

Public Property Let ShowGUI(ByVal vNewValue As Boolean)
' Creator : Daniel Siew 7/97
' Last Modified : 5/8/97
' Description : Set to TRUE when you want a GUI to be displayed.
'               Set to FALSE when you do not want a GUI to be displayed.
' Usage:
'   ShowGUI = TRUE
'
    mbShowGUI = vNewValue
End Property

Public Property Get StoreID() As Integer
' Property:
'   This procedure gets the store_id from store table
'   the assumption is there would be only one store per database
'   It uses the standard SCS0001DSN01 DSN to get to the data base
'   It goes to the database to retrieve the store ID only in the first call
'   it also raise an error if the Store table is empty

' Created By John Atabak on 05 June 1998
    On Error GoTo errStoreID
    
    Dim con As New rdoConnection
    Dim rsResult As rdoResultset
    Dim sSQL As String
    
        
    If miStoreID = 0 Then
        Set con = rdoEnvironments(0).OpenConnection("SCS0001DSN01", rdDriverNoPrompt, False, "")
        
        sSQL = "SELECT Str_ID FROM store_Status"

        Set rsResult = con.OpenResultset(sSQL, rdOpenKeyset)

        If rsResult.BOF And rsResult.EOF = True Then
           Call err.Raise(50001, "Store Specification", "The Store table is empty")

        Else
            rsResult.MoveFirst
            miStoreID = rsResult!Str_ID
        End If

        rsResult.Close
        con.Close
    End If

    StoreID = miStoreID

    Set con = Nothing

    Exit Property
       
errStoreID:
    Call err.Raise(err.Number, err.Source, err.Description, "", 0)

End Property

'
' Function to indicate whether program is running within the VB IDE environment
'                           or as a stand-alone executable.
'
' CAVEAT - it is assumed that no-one names their program "VBn......"
'           where the first two characters are "VB' and the third is a numeric.
'
Public Function RunningAsEXE() As Boolean                   ' 31st August 1998 - BByrne
    Dim sFileName As String
    Dim lCount As Long
    Dim sProgName As String
    Dim iIndex As Integer
    
    RunningAsEXE = True ' Assume running as an EXE
    ' Clear buffer
    sFileName = String(255, 0)
    ' Get FileName
    lCount = GetModuleFileName(App.hInstance, sFileName, 255)
    ' Strip crap off
    sFileName = Left(sFileName, lCount)
    
    ' Backup to the last \ character
    For iIndex = Len(sFileName) To 0 Step -1
        If Mid(sFileName, iIndex, 1) = "\" Then
            Exit For
        End If
    Next iIndex
    ' Extract Program Name
    sProgName = UCase(Mid(sFileName, iIndex + 1, Len(sFileName) - iIndex))
    ' Test program name is 'VB' followed by a numeric
    If Left(sProgName, 2) = "VB" And IsNumeric(Mid(sProgName, 3, 1)) Then
        RunningAsEXE = False
    End If
End Function

