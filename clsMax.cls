VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class: ClsMax
' Creator : Kien Chan 24/02/1997
'
' Overall Function:
' This class is used to handle the communication between out VB apps and the MAX
'
' Public Functions:
' TermKeyLoad
' MaxNumericOnSet
' FuncKeyPressCheck
' ProcessBackSpace
' ProcessScanOnCheck
' FuncNameGet
' FuncNameMaxGet
' FuncSeqGet
' FuncMaxSeqGet
' PrintTitle
' PrintStr
' AcceptStr
' ClearScr
' ClearMsgLine
' MsgLine
' DlgBox
' PrintStrPad
' ClearAfter
' HexDump
' mDebug_log
' RefreshScr
' EchoToClient
' NewScr
' CloseScr
' ScrollDown
' ScrollUp
' ScrollScr
' MultipleSelectScrollScr
' DisplayList
' SaveTerminatorOverstrike
' NewScrollScreen
' DisplayLine
' DisplayLineStr
' UnexpectedErrorFatal
' DlgVisible
'
' Last Modified : Pauline Tran 06/05/1997
'                 Dean Lane 13/10/1997 Expose DlgVisible property to show when dialogue box
'                                      is displayed
'                   Pauline Tran 09/02/1998 Modified the MsgLine function

' Last Modified : An Huynh 8/9/1998 Modified mDebug_Log Write to Log directory instead of Bin
'                   defect: 1001 tag 1001AH
Option Explicit

Const miMaxDisplayLines = 4
Const OneMegaByte = 1048576 '1001AH

Private msSTB(6)                    As String * 255   'State Table
Private msKSeq(MAXKEYSEQ)           As String * 6     'Key Seq table for esc seq
Private msFuncNm(MAXKEYSEQ)         As String * 20    'Function Name table
'
'for MAXNUMERIC_ON MAXNUMERIC_OFF etc
'similar to above for function keys
Private msMaxSTB(6)                 As String * 255   'Cannot mix with msSTB
Private msMaxKSeq(MAXMAXKEYSEQ)     As String * 6     '

Private msMaxFuncNm(MAXMAXKEYSEQ)   As String * 20    '

'
Private miKeySeq                    As Integer
Private msSpace20                   As String * 20
Private msDBLoadDone                As String
Private miScrCnt(MAXSCR)            As Integer
Private mbSavedTerminatorCheck      As Boolean
Private mbSavedOverStrike           As Boolean
Private msInputStr                  As String
Private msFuncKeyNm                 As String

'property variables declartions
Dim pbAlpha                         As Boolean
Dim piClnIndex                      As Integer
Dim piCurScreen                     As Integer
Dim piCurIndex                      As Integer
Dim pbDebugLogOff                   As Boolean
Dim piPrvScreen                     As Integer
Dim pbTerminatorCheck               As Boolean
Dim pbScanOnCheck                   As Boolean
Dim psScanType                      As String
Dim psDefaultValue                  As String
Dim piDefaultColumn                 As Integer
Dim piDefaultValueLen               As Integer
Dim pbOverStrike                    As Boolean
Dim psScrollTitle                   As String
Dim pbDlgVisible                    As Boolean  'dl

Public Sub ScrollRsListDown(rs As rdoResultset, _
                         iLineToPrintFirstRecord As Integer, _
                         intMaxLinesDisplayed As Integer, _
                         iCursorPos As Integer, vColumns As Variant)
'scrolls down the list on Max when list displayed is
'a part of the resultset passed

    Dim iRecordCur              As Long
    Dim sCharacterUnderCursor   As String
    
    On Error GoTo Error_SDRL
    
    MsgLine ""
    iRecordCur = rs.AbsolutePosition
    If iRecordCur = rs.RowCount Then               ' at last record
        MsgLine "At last record."
    ElseIf CurrentIndex = iLineToPrintFirstRecord + intMaxLinesDisplayed - 1 Then ' scroll down one
        DisplayRsList rs, iLineToPrintFirstRecord, iRecordCur - intMaxLinesDisplayed + 2, _
                      iRecordCur + 1, iCursorPos, vColumns
        
        rs.AbsolutePosition = iRecordCur + 1
    Else                                    ' move cursor down one
        CurrentIndex = CurrentIndex + 1
        rs.AbsolutePosition = iRecordCur + 1
    End If
    
    AcceptStr CurrentIndex, iCursorPos, ""    'display cursor
    
    Exit Sub
Error_SDRL:
    HandleErrorFatal "clsMax", "ScrollRsListDown", err.Source, err.Number, err.Description

End Sub

Public Sub ScrollRsListUp(rs As rdoResultset, _
                         iLineToPrintFirstRecord As Integer, _
                         intMaxLinesDisplayed As Integer, _
                         iCursorPos As Integer, vColumns As Variant)
'scrolls up the list on Max when list displayed is
'a part of the resultset passed

    Dim iRecordCur              As Long
    Dim sCharacterUnderCursor   As String
    
    On Error GoTo Error_SURL
    
    MsgLine ""
    
    iRecordCur = rs.AbsolutePosition
    If iRecordCur = 1 Then               ' at last record
        MsgLine "At first record."
    ElseIf CurrentIndex = iLineToPrintFirstRecord Then  ' scroll up one
        DisplayRsList rs, iLineToPrintFirstRecord, iRecordCur - 1, _
                      iRecordCur + intMaxLinesDisplayed - 1, iCursorPos, vColumns
        
        rs.AbsolutePosition = iRecordCur - 1
    Else                                    ' move cursor down one
        CurrentIndex = CurrentIndex - 1
        rs.AbsolutePosition = iRecordCur - 1
    End If
    
    AcceptStr CurrentIndex, iCursorPos, ""    'display cursor
    
    Exit Sub
Error_SURL:
    HandleErrorFatal "clsMax", "ScrollRsListUp", err.Source, err.Number, err.Description

End Sub

Public Sub TermKeyLoad()
'Get escape sequence and key names from database cmsdb.term_func_seq
'
' Modified By: Pauline Tran 06/05/1997
'              Dean Lane 22/9/1997 DefId:71 First screen initialised twice
     
    On Error GoTo Error
     
    Dim i  As Integer
 
    'initialize the sequence table to all '1'
    InitializeMaxKeyTable               'maxnumeric_on etc tables
    InitializeKeyTable                  'normal function keys
    
    'initialize scrcnt array to available
    For i = 1 To MAXSCR
        miScrCnt(i) = -1               'set to available
    Next i

    miKeySeq = 1                         ' starting check
    msSpace20 = String(MAXCOL, " ")      '20 spaces
    
    'Not required as Form_Term_Load will do initialisation
    piCurScreen = NewScr                 'Get a screen index should be 1
    piCurScreen = 1

    Exit Sub
Error:
    HandleErrorFatal "ClsMax", "TermKeyLoad", err.Source, err.Number, err.Description

End Sub

Private Sub InitializeMaxKeyTable()
' Loads psMAXKeyTB & msMaxSTB
'
' Modified By: Pauline Tran 06/05/1997

    On Error GoTo Error
    
    Dim iLen   As Integer
    Dim iSeq   As Integer
    Dim iKey   As Integer
    Dim sStr   As String
    Dim iKeyno As Integer
     
    iKeyno = DBLoadKeys("MAX")          'get no key definition from db
    For iSeq = 1 To 6
            msMaxSTB(iSeq) = String(255, "1")       'initialise to all 1
    Next iSeq
    
    'Load the escape sequence to a table map
    For iKey = 1 To iKeyno
        
        sStr = Trim(msMaxKSeq(iKey))
        iLen = Len(sStr)
        
        For iSeq = 1 To iLen
            
            If iSeq > 6 Then
                  HandleErrorFatal "ClsMax", "InitializeMaxKeyTable", "> 6 seq not supprtd", "121", "..Pls increase"
            End If
            
            If iSeq = iLen Then
                   Mid$(msMaxSTB(iSeq), Asc(Mid$(sStr, iSeq, 1)), 1) = ENDKEYSEQ 'end of list
            Else
                   Mid$(msMaxSTB(iSeq), Asc(Mid$(sStr, iSeq, 1)), 1) = iSeq + 1
            End If
          
        Next iSeq
        
    Next iKey
    
    Exit Sub
Error:
    HandleErrorFatal "ClsMax", "InitializeMaxKeyTable", err.Source, err.Number, err.Description
    
End Sub

Private Sub InitializeKeyTable()
' Normal function key tables

    On Error GoTo Error
    
    Dim iLen   As Integer
    Dim iSeq   As Integer
    Dim iKey   As Integer
    Dim sStr   As String
    Dim iKeyno As Integer
     
    iKeyno = DBLoadKeys("NONMAX")          'get no key definition from db
    
    For iSeq = 1 To 6
        msSTB(iSeq) = String(255, "1")       'initialise to all 1
    Next iSeq
    
    'Load the escape sequence to a table map
    For iKey = 1 To iKeyno
        
        sStr = Trim(msKSeq(iKey))
        iLen = Len(sStr)
        
        For iSeq = 1 To iLen
            
            If iSeq > 6 Then
                HandleErrorFatal "ClsMax", "InitializeKeyTable", "> 6 seq not supprtd", "120", "..Pls increase"
            End If
            
            If iSeq = iLen Then
                   Mid$(msSTB(iSeq), Asc(Mid$(sStr, iSeq, 1)), 1) = ENDKEYSEQ 'end of list
            Else
                   Mid$(msSTB(iSeq), Asc(Mid$(sStr, iSeq, 1)), 1) = iSeq + 1
            End If
          
        Next iSeq
        
    Next iKey
    
    Exit Sub
Error:
    HandleErrorFatal "ClsMax", "InitializeKeyTable", err.Source, err.Number, err.Description
    
End Sub

Private Function DBLoadKeys(sType As String) As Integer
' Get escape sequence and key names from database
' Two tables to load
' I have to load into 2 different tables because the escape sequence for
' MAX* function keys interfere with other function keys
' pskeySeq & psMaxkeySeq
'
' Modified By: Pauline Tran 06/05/1997
'
' Parameters:
' sType
    On Error GoTo Error
    
    Dim iLen   As Integer
    Dim iSeq   As Integer
    Dim iKey   As Integer
    Dim sStr   As String
    Dim iKeyno As Integer
    Dim con    As rdoConnection
    Dim rs     As rdoResultset
    Dim i      As Integer
    Dim sSQL   As String
     
    
    If sType = "MAX" Then
       sSQL = "select * from term_func_seq where func_key_nm like 'MAX%'"
    Else
       sSQL = "select * from term_func_seq where func_key_nm  not like 'MAX%'"
    End If
    
    'Set con = CMSDBOpen
    'closing of this at the end of this function closes con opened by
    'CMSDBOpen w/o decrementing the counter there and hence does not
    'return a connection on consequent cals.  CAN'T WORK THIS WAY - VS 07/05/97
'    Set con = rdoEnvironments(0).OpenConnection("", rdDriverNoPrompt, False, "dsn=simdsn01;uid=;pwd=;database=cmsdb")
    Set gConnection = OpenSIM0001DSN01
    Set rs = gConnection.OpenResultset(sSQL)
    
    If rs.EOF Then
        HandleErrorFatal "ClsMax", "Dbloadkeys", "NO Records", "110", "Load 'term_func_seq' with MAX function keys."
    End If
    
    iKey = 0
    
    Do While Not rs.EOF
        iKey = iKey + 1
        If sType = "MAX" Then
            If iKey > MAXMAXKEYSEQ Then
               
                HandleErrorFatal "ClsMax", "Dbloadkeys", "Max Key exceeded", "100", "..Pls increase"
                
            End If
        Else
           If iKey > MAXKEYSEQ Then
   
                HandleErrorFatal "ClsMax", "Dbloadkeys", "Max Key exceeded", "101", "..Pls increase"
  
           End If
        End If
        
        If sType = "MAX" Then
            msMaxKSeq(iKey) = rs!func_key_seq
            msMaxFuncNm(iKey) = Trim(UCase(rs!func_key_nm))
        Else
            msKSeq(iKey) = rs!func_key_seq
            msFuncNm(iKey) = Trim(UCase(rs!func_key_nm))
        End If
        rs.MoveNext
    Loop
    
    DBLoadKeys = iKey           'no of keys loaded
    rs.Close
'    con.Close
    msDBLoadDone = "Y"
    
    Exit Function
Error:
    HandleErrorFatal "ClsMax", "Dbloadkeys", err.Source, err.Number, err.Description

End Function

Public Function MaxNumericOnSet(vtStr As Variant, sFuncName As String) As Boolean
'Check to see server sents a Numeric on etc
'
' Modified By: Pauline Tran 06/05/1997
'
' Parameters:
' vtStr
' sFuncName
    
    On Error GoTo Error
    
    Dim sStr        As String
    Dim i           As Integer
    Dim iFuncLen    As Integer
    Dim sTmpstr     As String
    
    MaxNumericOnSet = False                       'not a func key
    iFuncLen = 1                                    'set to 1
    
    For i = 1 To Len(vtStr)
    
        miKeySeq = Mid$(msMaxSTB(miKeySeq), Asc(Mid(vtStr, i, 1)), 1)  'map into table
        
        If miKeySeq = ENDKEYSEQ Then               ' okay it is a escape sequence
            
            sStr = Mid(vtStr, i - iFuncLen + 1, iFuncLen) 'func seq
            FuncNameMaxGet sStr, sFuncName                ' what type of func key ?
            
            If Mid$(sFuncName, 1, 3) = "MAX" Then          'like MAX????
                MaxNumericOnSet = True                  'is a func key or cr lf
                sTmpstr = Mid(vtStr, 1, i - iFuncLen)   'remove func key seq
                vtStr = sTmpstr & Mid(vtStr, i + 1) 'strip seq
                
                miKeySeq = 1
                Exit For
            End If
            
            miKeySeq = 1                            'Reset check sequence to beginning
        Else
            iFuncLen = miKeySeq                 'func key len
        End If
    Next i

    Exit Function
Error:
    HandleErrorFatal "ClsMax", "MaxNumericOnSet", err.Source, err.Number, err.Description

End Function

Public Function FuncKeyPressCheck(vtStr As Variant, sFuncName As String) As Boolean
'-----------------------------------------------------------------------------
'Checking is driven from table
'Return true if e.g CRLF is entered etc
'can toggle not to check for terminator by using TerminatorCheck=false
'this is useful if e.g you want to user to just enter a character (i.e not entering
'character + <enter>)
'
' Modified By: Pauline Tran 06/05/1997
'
' Parameters:
' vtStr
' sFuncName
' Modified to include CLEAR S CLEAR P CLEAR N

    On Error GoTo Error
    
    Dim sStr                    As String
    Dim i                       As Integer
    Dim iSeqLen                 As Integer
    Dim sSavedStr               As String
    
    Static sPrvKey              As String
    
    
    FuncKeyPressCheck = False                       'not a func key
    iSeqLen = 1                                    'set to 1
    
    sSavedStr = vtStr
    For i = 1 To Len(vtStr)
    
        miKeySeq = Mid$(msSTB(miKeySeq), Asc(Mid$(vtStr, i, 1)), 1)  'map into table
        
        If miKeySeq = ENDKEYSEQ Then               ' okay it is END  escape sequence
            
            sStr = Mid$(vtStr, i - iSeqLen + 1, iSeqLen)  'Get Escape sequence
            FuncNameGet sStr, sFuncName             ' Get function key name ?
            FuncKeyPressCheck = True                'is a function key or C
            msFuncKeyNm = sFuncName                 'func key name
            miKeySeq = 1                            'Reset check sequence to beginning
            vtStr = Mid(vtStr, 1, i - iSeqLen)      'remove esc seq
            
            If ScanOnCheck Then                     'MaxScan_on sets scanoncheck=true
            
                If msFuncKeyNm = KEYRETURN Then
                    
                    vtStr = ProcessScanOnCheck(vtStr)    '?12345678,E
                    ScanOnCheck = False                  'turn off check
                End If
            
            End If
            
            Exit For                                 'found it
        
        Else
            iSeqLen = miKeySeq                        'seq key len
        End If
    
    Next i
    
    If Not FuncKeyPressCheck Then
        
        If Not TerminatorCheck Then               'No end of field check
            
            sFuncName = UCase(vtStr)
            msFuncKeyNm = sFuncName
            FuncKeyPressCheck = True
     
        End If
        
    Else
        
        If sFuncName = KEYBACKSPACE Then
            ProcessBackSpace                    'go back 1 and reset cursor
            FuncKeyPressCheck = False           'It's  not a func key for bS
        End If
 
    End If
       
    If sPrvKey = Chr(KEYCLEAR) Then  ' previous was a clear key
        If msFuncKeyNm = KEYDOWN Then
            msFuncKeyNm = KEYCLEAR_N
            FuncKeyPressCheck = True
        Else
            Select Case UCase(vtStr)
            Case "S", "5"
                msFuncKeyNm = KEYCLEAR_S
                FuncKeyPressCheck = True
            Case "P", "8"
                msFuncKeyNm = KEYCLEAR_P
                FuncKeyPressCheck = True
            Case "N"
                msFuncKeyNm = KEYCLEAR_N
                FuncKeyPressCheck = True
            End Select
        End If
        vtStr = ""
    End If
    sPrvKey = sSavedStr          'reset prv key
    
    Exit Function
Error:
    HandleErrorFatal "ClsMax", "FuncKeyPressCheck", err.Source, err.Number, err.Description
       
End Function

Public Sub ProcessBackSpace()
' procedure to handle the BackSpace key pressed on the max
'
' Modified By: Pauline Tran 06/05/1997

    On Error GoTo Error
    
    Dim X       As Integer
    Dim Y       As Integer
    Dim iIndex  As Integer
    
    iIndex = CurrentScreen
    
    X = StdScr(iIndex).Cur_x               'Current row
    
 
    Y = StdScr(iIndex).Cur_y
        
    If Y > DefaultColumn Then
        If OverStrike Then
            PrintStr X, Y, " "
            Y = Y - 1
        Else
            Y = Y - 1
            PrintStr X, Y, " "                    'Do it
        End If
        
        PrintStr X, Y, ""
    
    End If
    
    Exit Sub
Error:
    HandleErrorFatal "ClsMax", "ProcessBackSpace", err.Source, err.Number, err.Description
  
End Sub

Public Function ProcessScanOnCheck(vTmp As Variant) As String
' Function to Check the if the Max is ready for Scanning
'
' Modified By: Pauline Tran 06/05/1997
'
' Parameters:
' vTmp

    On Error GoTo Error
    
    Dim iPos As Integer
    
    iPos = InStr(vTmp, ",")                   ' "," found
    ScanType = ""                             'initialize
   ' vTmp = "088698076613,C"

    If iPos <> 0 Then                               'found ","
        'WriteToErrorLog "len vtemp in clsmax" & Len(vTmp)
        If Len(vTmp) <= 10 Then                     'Sell id 8 + ",F"
        'WriteToErrorLog "inside if in clsmax" & Len(vTmp)
            ScanType = Mid$(vTmp, iPos + 1, 1)       'get type scan type
            vTmp = Mid(vTmp, 1, iPos - 2)           'strip off checkdig &',' onwards
        End If
    End If
    
    ProcessScanOnCheck = vTmp                       'return value
 
    Exit Function
Error:
    HandleErrorFatal "ClsMax", "ProcessScanOnCheck", err.Source, err.Number, err.Description
 
End Function

Public Sub FuncNameGet(sFuncSeq As String, sFuncName As String)
'get the function key name when given the escape sequence
'e.g esc[B gives 'BS' (backspace)
'find the match in array done by TermKeyLoad
'
' Modified By: Pauline Tran 06/05/1997
'
' Parameters:
' sFuncSeq
' sFuncName

    On Error GoTo Error
    
    Dim i As Integer

    If msDBLoadDone <> "Y" Then
        DlgBox "", "DBKEYLOAD not done", " ..aborted", "", MSGTYP_YESNO
        End
    End If

    For i = 1 To MAXKEYSEQ
        
        If sFuncSeq = Trim(msKSeq(i)) Then
            sFuncName = Trim(msFuncNm(i))
            Exit Sub
        End If
    
    Next i

    Exit Sub
Error:
    HandleErrorFatal "clsMax", "FuncNameGet", err.Source, err.Number, err.Description

End Sub

Public Sub FuncNameMaxGet(sFuncSeq As String, sFuncName As String)
'get the function key name from max table
'find the match in array done by TermKeyLoad
'
' Modified By: Pauline Tran 06/05/1997
'
' Parameters:
' sFuncSeq
' sFuncName

    On Error GoTo Error
    
    Dim i As Integer

    If msDBLoadDone <> "Y" Then
        HandleErrorFatal "ClsMax", "funcNameMaxget", "DBLOAD not done", "0", "Aborted"
    End If

    For i = 1 To MAXMAXKEYSEQ
        
        If sFuncSeq = Trim(msMaxKSeq(i)) Then
            sFuncName = Trim(msMaxFuncNm(i))
            Exit Sub
        End If
    
    Next i

    Exit Sub
Error:
    HandleErrorFatal "clsMax", "FuncNameMaxGet", err.Source, err.Number, err.Description

End Sub

Public Sub FuncSeqGet(sFuncName As String, sFuncSeq As String)
'get the function escape sequence for a given function key name
'find the match in array done by TermKeyLoad
'
' Modified By: Pauline Tran 06/05/1997
'
' Parameters:
' sFuncName
' sFuncSeq
    
    On Error GoTo Error
    
    Dim i As Integer
    
    If msDBLoadDone <> "Y" Then
        HandleErrorFatal "ClsMax", "funcseqget", "DBLOAD not done", "0", "Aborted"
    End If

    For i = 1 To MAXKEYSEQ
        
        If sFuncName = Trim(msFuncNm(i)) Then
            sFuncSeq = Trim(msKSeq(i))
            Exit Sub
        End If
    
    Next i

    Exit Sub
Error:
    HandleErrorFatal "clsMax", "FuncSeqGet", err.Source, err.Number, err.Description

End Sub

Public Function GetInputNo(vData As Variant) As Variant
'Last modified by:Fariba Mokarram
'Function to get the ascii code for the input
    On Error GoTo Error
    
    Dim iCount As Integer

    For iCount = 1 To Len(vData)
        GetInputNo = GetInputNo & Format(Asc(Mid(vData, iCount, 1)), "000")
    Next

    Exit Function
Error:
    HandleErrorFatal "clsMax", "GetInputNo", err.Source, err.Number, err.Description

End Function

Public Sub FuncMaxSeqGet(sFuncName As String, sFuncSeq As String)
'get the function esc seq for a given function key name
'For MAXNUMERIC_ON MAXNUMERIC_OFF MAXSCAN_ON etc
'
' Modified By: Pauline Tran 06/05/1997
' Parameters:
' sFuncName
' sFuncSeq
    On Error GoTo Error
    
    Dim i As Integer
    
    If msDBLoadDone <> "Y" Then
       
        HandleErrorFatal "ClsMax", "funcMAxseqget", "DBLOAD not done", "0", "Aborted"
 
    End If

    For i = 1 To MAXMAXKEYSEQ
    
        If sFuncName = Trim(msMaxFuncNm(i)) Then
            sFuncSeq = Trim(msMaxKSeq(i))
            Exit Sub
        End If
    
    Next i

    Exit Sub
Error:
    HandleErrorFatal "clsMax", "FuncMaxSeqGet", err.Source, err.Number, err.Description

End Sub

Public Sub PrintTitle(sStr As String)
' Centers a string and print it as the Title on the MAX screen
'
' Modified By: Pauline Tran 06/05/1997
'
' Parameters:
' sStr  string to be printed on the Title line

    On Error GoTo Error
    
    Dim iCol As Integer

    'If Len(s) < MAXCOL Then
    'iCol will be 0 if len(s) = 19. to avoid this situation above line replaced with following
    If Len(sStr) < MAXCOL - 1 Then
        iCol = (MAXCOL - Len(sStr)) / 2      'center to line
    Else
        iCol = 1
    End If

    PrintStr 1, iCol, sStr

    Exit Sub
Error:
    HandleErrorFatal "clsMax", "PrintTitle", err.Source, err.Number, err.Description

End Sub

Public Sub PrintStr(X As Integer, Y As Integer, s As String)
'-------------------------------------------------------------------
'routine to print a string s at row x and column y
'The only escape sequence used in this project csiPx;PyH
'This is where everything sent to MAX is store in stdscr(i)
'
' Modified By: Pauline Tran 06/05/1997
'
' Parameters:
' x
' y
' s
'-------------------------------------------------------------------
    On Error GoTo Error
    
    Dim vtBuf   As Variant
    Dim iStrLen As Integer
    Dim iIndex  As Integer
    
    iStrLen = Len(s)
    
    If Y + iStrLen > MAXCOL Then
       
       iStrLen = MAXCOL - Y + 1
       
       If iStrLen < 0 Then
           
           iStrLen = 0
       
       End If
       
    End If
    
    s = Mid$(s, 1, iStrLen)                             'drop off if more than maxcol
    iIndex = CurrentScreen
    StdScr(iIndex).Cur_x = X                            'current x
    StdScr(iIndex).Cur_y = Y + iStrLen                  'current y
    If s <> "" And Y <= MAXCOL And iStrLen <= Len(s) Then
        Mid$(StdScr(iIndex).Scr(X), Y, iStrLen) = s      'store it in x,y
    End If
       
    vtBuf = Chr(27) & "[" & X & ";" & Y & "H" & s  'esc [ 'x';'y'H string
    frmMain.TCP1(giSockIndex).SendData vtBuf
    
    Exit Sub
Error:

    HandleErrorFatal "clsMax", "PrintStr", err.Source, err.Number, err.Description
    
End Sub

Public Sub AcceptStr(X As Integer, Y As Integer, _
                     Optional DefaultStr, Optional DefaultStrLen)
'Accept from x,y coordinates after clear the field
' DefaultStr is the default value for the string
' DefaultStrLen is the default length of the string
'
' Modified By: Pauline Tran 05/05/1997
'
' Parameters:
' x
' y
' DefaultStr Optional
' DefaultStrLen Optional
    On Error GoTo Error
    
    If IsMissing(DefaultStr) Then           'no default str
    
        DefaultValue = ""
        DefaultColumn = Y
        DefaultValueLen = MAXCOL - Y
        PrintStr X, Y, Mid$(String(MAXCOL, " "), 1, MAXCOL - Y + 1)
        
        If OverStrike Then                  'with default values
        
            If Y <= DefaultColumn Then
                PrintStr X, Y, ""
            Else
                PrintStr X, Y - 1, ""
            End If
        
        Else
        
            PrintStr X, Y, ""
        End If
    Else
        DefaultColumn = Y                   'Very Important for working was input by user
       
        If IsMissing(DefaultStrLen) Then
            DefaultValueLen = Len(DefaultStr) 'Use in clsmax.Inputstr
        Else
            DefaultValueLen = DefaultStrLen
        End If
       
    '          DefaultStr = Trim(DefaultStr)        'remove spaces & left justify
       'vs 24/04/97 not to trim as you may like to include spaces
        DefaultStr = DefaultStr        'remove spaces & left justify
        DefaultValue = DefaultStr
       
        If DefaultValueLen = 0 Then
            DefaultValueLen = MAXCOL - Y
        End If
        
        PrintStr X, Y, CStr(DefaultStr)
    
        If OverStrike Then
            If Y + Len(DefaultStr) - 1 = 0 Then
                PrintStr X, 1, ""
            Else
                PrintStr X, Y + Len(DefaultStr) - 1, "" 'overstrike
            End If
       
        Else
            PrintStr X, Y + Len(DefaultStr), ""     'next column
        End If
    End If
        
    Exit Sub
Error:
    HandleErrorFatal "clsMax", "AcceptStr", err.Source, err.Number, err.Description
        
End Sub

Public Sub ClearScr()
' Sub to clear the current MAX screen
'I do not want to introduce another escape sequence for clear screen
'so I use printstr
'
' Modified By: Pauline Tran 06/05/1997
 
    On Error GoTo Error
    
    Dim i As Integer
 
    For i = 1 To MAXROW
      PrintStr i, 1, msSpace20
    Next i
    
    Exit Sub
Error:
    HandleErrorFatal "clsMax", "ClearScr", err.Source, err.Number, err.Description

End Sub

Public Sub ClearMsgLine()
' procedure to clear or blank out the message line on the MAX screen
'
' Modified By: Pauline Tran 06/05/1997

    On Error GoTo Error
    
    Dim iSavedx          As Integer
    Dim iSavedy          As Integer
    Dim iIndex           As Integer
    
    iIndex = CurrentScreen                'get current screen
    
    If StdScr(iIndex).MsgCnt > 0 Then     'Is there one to be done ?
        
        StdScr(iIndex).MsgCnt = StdScr(iIndex).MsgCnt - 1
        iSavedx = StdScr(iIndex).Cur_x
        iSavedy = StdScr(iIndex).Cur_y
        PrintStr MSGROW, MSGCOL, msSpace20
        PrintStr iSavedx, iSavedy, ""          'restore previous cursor position
    
    End If
    
    Exit Sub
Error:
    HandleErrorFatal "clsMax", "ClearMsgLine", err.Source, err.Number, err.Description

End Sub

Public Sub MsgLine(sStr As String, Optional ErrBell As Boolean)
' Prints a string at the last line (line 8) of the MAX screen
'
' Parameters:
' sStr
' ErrBell Optional - to allow an Error Bell to sound.
'
' Modified By:  Pauline Tran 06/05/1997
'               Pauline Tran 09/02/1998 Added ErrBell parameter to activate Error Bell
    On Error GoTo Error
    
    Dim iIndex As Integer
    Dim iSpace As Integer
    
    iIndex = CurrentScreen

    iSpace = MAXCOL - Len(sStr)        '# of spaces to fill up MAXCOL remaining
    If iSpace < 0 Then
       iSpace = 1
    End If
    
    StdScr(iIndex).MsgCnt = StdScr(iIndex).MsgCnt + 1
    
    PrintStr MSGROW, MSGCOL, sStr & Mid$(String(MAXCOL, " "), 1, iSpace)
    
    If ErrBell Then
        MaxErrBell
    End If
    
    Exit Sub
Error:
    HandleErrorFatal "clsMax", "MsgLine", err.Source, err.Number, err.Description

End Sub

Public Function DlgBox(sInStr As String, sMsg1 As String, _
                  sMsg2 As String, sMsg3 As String, _
                  Optional vMsgTyp As Variant) As String
' Function to prints a message box onto the MAX screen
' Takes in 4 strings and a Message Type flag
'
' Modified By: Pauline Tran 06/05/1997
'
' Parameters:
' sInStr
' sMsg1
' sMsg2
' sMsg3
' vMsgTyp Optional
    On Error GoTo Error
    
    Dim sMsg4                       As String
    Static sDisplayDone             As String
    Static iTempScr                 As Integer
    Static bChangedAlpha            As Boolean
    
    If IsMissing(vMsgTyp) Then
        vMsgTyp = MSGTYP_YESNO      'default
    End If
    
    Select Case vMsgTyp
        Case MSGTYP_YESNO
            sMsg4 = "Press Y or N "
        Case MSGTYP_INFO
            sMsg4 = "Press <Enter>"
        Case Else
            HandleErrorFatal "ClsMax", "DLGBOX", "Msgtype", "0", "Invalid type"
    End Select
    
    If sDisplayDone <> "Y" Then                  'display msg

        mbSavedTerminatorCheck = TerminatorCheck 'saved it up
        sDisplayDone = "Y"
        DlgVisible = True  'dl
        iTempScr = NewScr                        'get a new screen
        PreviousScreen = CurrentScreen           'move current to prev
        CurrentScreen = iTempScr
        PrintStr 2, 1, String(MAXCOL, "-")
        PrintStrPad 3, 1, sMsg1
        PrintStrPad 4, 1, sMsg2
        PrintStrPad 5, 1, sMsg3
        PrintStrPad 6, 1, sMsg4
        PrintStr 7, 1, String(MAXCOL, "-")
        If Not MSGTYP_INFO Then
            If pbAlpha Then                'Is it alpha mode
                bChangedAlpha = False      'no changed
            Else
                MaxNumeric_Off             ' turn alpha on
                bChangedAlpha = True
            End If
        Else
            bChangedAlpha = False          'nothing changed
        End If
        
        MaxWarnBell                  'ring the warn bell
        
        TerminatorCheck = True       'wait until CRLF
        AcceptStr 6, Len(sMsg4) + 4

    Else
        If vMsgTyp = MSGTYP_YESNO Then
            If UCase(sInStr) = "Y" Or UCase(sInStr) = "N" Then
                DlgBox = UCase(sInStr)
                CloseScr                                 'release screen
                CurrentScreen = PreviousScreen           'restore saved screen
                RefreshScr                               'refresh previous screen
                sDisplayDone = "N"
                DlgVisible = False  'dl
                TerminatorCheck = mbSavedTerminatorCheck 'restore to original
                If bChangedAlpha Then
                  MaxNumeric_On                         'go back to numeric on
                End If
            Else
                MaxErrBell                               '3 beep err bell
                MsgLine "Invalid Input "
                'AcceptStr 6, Len(sMsg4) + 1
                'VS 30/05/97 cursor should be placed where input is required
                'which is 4th column after sMsg4 as in last line of code in this routine
                AcceptStr 6, Len(sMsg4) + 4
                MaxNumeric_Off     'FM  go back to aplha after error msg
            End If
        Else
            If FuncKeyName = KEYRETURN Then
                  '
                  '  For case = MSGTYP_INFO, default acknowledge is assumed if Keyreturn is deteched.
                  '      Before this design bug fix, the caller of this DLG routine has no means to know whether
                  '      user input is OK or not. (ie DlgBOX always return "" anyway.)
                  '
                  DlgBox = ""
                  CloseScr
                  CurrentScreen = PreviousScreen
                  RefreshScr                               'refresh previous screen
                  sDisplayDone = "N"
                  DlgVisible = False  'dl
                  TerminatorCheck = mbSavedTerminatorCheck 'restore to original
            Else
                PrintStr 2, 1, String(MAXCOL, "-")
                PrintStrPad 3, 1, sMsg1
                PrintStrPad 4, 1, sMsg2
                PrintStrPad 5, 1, sMsg3
                PrintStrPad 6, 1, sMsg4
                PrintStr 7, 1, String(MAXCOL, "-")
                AcceptStr 6, Len(sMsg4) + 4
            End If
        End If
    End If

Exit Function
Error:
    HandleErrorFatal "clsMax", "DlgBox", err.Source, err.Number, err.Description
    

End Function

Public Sub PrintStrPad(X As Integer, Y As Integer, s As String)
'pad the end of the input String with spaces
'
' Modified By: Pauline Tran 06/05/1997
'
' Parameters:
' x
' y
' s
    On Error GoTo Error
    
    If Len(s) < MAXCOL Then
        PrintStr X, Y, s & String(MAXCOL - Len(s), " ")
    Else
        PrintStr X, Y, Mid$(s, 1, MAXCOL)
    End If

Exit Sub
Error:
    HandleErrorFatal "clsMax", "PrintStrPad", err.Source, err.Number, err.Description

End Sub

Public Sub ClearAfter(X As Integer, Y As Integer)
' Sub to clear a line (y) after the position x on the MAX Screen
'
' Modified By: Pauline Tran 06/05/1997
'
' Parameters:
' x
' y
    On Error GoTo Error
    
    PrintStr X, Y, String(MAXCOL - Y, " ")

Exit Sub
Error:
    HandleErrorFatal "clsMax", "ClearAfter", err.Source, err.Number, err.Description

End Sub

Public Sub HexDump(sText As String, sHexStr As String)
'Dump out the string in Hexdecimal format
'eg. 12 ...4142
'
' Modified By: Pauline Tran 06/05/1997
'
' Parameters:
' sText
' sHexStr
    On Error GoTo Error
    
    Dim i    As Integer
    Dim sHex As String

    'output str in HEX decimal
    For i = 1 To Len(sText)
        sHex = Hex(Asc(Mid$(sText, i, 1)))
        
        If Len(sHex) = 1 Then
            sHex = "0" & sHex
        End If
        
        sHexStr = sHexStr & sHex & " "
    Next i

Exit Sub
Error:
    HandleErrorFatal "clsMax", "HexDump", err.Source, err.Number, err.Description

End Sub

Public Sub mDebug_Log(s As String)
' Debugging function for MAX apps developers.
' Writes to a log file 'appname'.log in c:\temp
'
' Modified By: Pauline Tran 06/05/1997
' Modified By: An Huynh 8/9/1998
' Write Log to Logs directory instead of Bin
' Parameters:
' s
    On Error GoTo Err_mDebug_Log

    If Not DebugLogOff Then  'if set not to logoff
  
        Static sOpenFlag As String
        Dim lSize As Long
        Dim sBUP As String
        Dim sLogDir As String
        Static sFileLogPath As String
        
        If sOpenFlag <> "Y" Then
            If sFileLogPath = "" Then
                sLogDir = sGetParentDir(App.Path) 'DL
                ' Ensure directory exists                           ' BB 2/2/98
                If Dir(sLogDir & "\logs", vbDirectory) = "" Then    ' Doesn't exist
                    MkDir (sLogDir & "\logs")
                End If
                sFileLogPath = sLogDir & "\logs\" & App.EXEName & ".log"
            End If
            If Dir(sFileLogPath) <> "" Then                     ' If logfile found
                lSize = FileLen(sFileLogPath)                   ' Get filesize
                If lSize > OneMegaByte Then                     ' If log is larger than 1 Mb copy to backup
                    If InStr(sFileLogPath, ".") Then            ' Concat BUP behind "."
                        sBUP = Mid$(sFileLogPath, 1, InStr(sFileLogPath, ".") - 1)
                    End If
                    sBUP = sBUP & ".BUP"
                    FileCopy sFileLogPath, sBUP                 ' Copy log file to .BUP
                    Kill sFileLogPath                           ' Remove existing log file (so it can be started afresh)
                    Open sFileLogPath For Output As #1 '1001AH
                    Close #1
                End If
                Open sFileLogPath For Append As #1   '1001AH
            Else
                Open sFileLogPath For Output As #1 '1001AH
            End If
        Else
            Open sFileLogPath For Append As #1   '1001AH

        End If
        sOpenFlag = "Y"

        Print #1, Date, Time, s & vbCrLf
        Close #1           'close so that you can view it as soon as it is written
    End If

Exit Sub
Err_mDebug_Log:
    LogError "clsMax", "mDebug_Log", False, False
    err.Raise err.Number
End Sub

Public Sub mScreenLog(sLine As String)
        
   On Error GoTo Error
   
   Static sOpenFlag As String
   Static fhOutput As Integer
    
   If sOpenFlag = "Y" Then
      Open App.Path & "\" & App.EXEName & ".txt" For Append As #fhOutput
   Else
      fhOutput = FreeFile
      Open App.Path & "\" & App.EXEName & ".txt" For Output As #fhOutput
      sOpenFlag = "Y"
  End If
  
  Print #fhOutput, sLine
  Close #fhOutput           'close so that you can view it as soon as it is written

Exit Sub
Error:
    HandleErrorFatal "clsMax", "mScreenLog", err.Source, err.Number, err.Description

End Sub

Public Sub RefreshScr()
' This function refreshes the Current screen on the MAX
'
' Modified By: Pauline Tran 06/05/1997

    On Error GoTo Error
    
    Dim i           As Integer
    Dim iPrvX       As Integer
    Dim iPrvY       As Integer
    Dim iIndex      As Integer
    
    iIndex = CurrentScreen               'refresh current screen
    
    iPrvX = StdScr(iIndex).Cur_x         'saved up current x y
    iPrvY = StdScr(iIndex).Cur_y         'for restoring later
    
    For i = 1 To MAXROW
        PrintStr i, 1, StdScr(iIndex).Scr(i)
    Next i
    PrintStr iPrvX, iPrvY, ""    'restore   cursor

Exit Sub
Error:
    HandleErrorFatal "clsMax", "RefreshScr", err.Source, err.Number, err.Description

End Sub

Public Sub EchoToClient(index As Integer, vtGetData As Variant, Optional bCase As Boolean)
' This function echos to the Client computer
'
' Modified By: Pauline Tran 06/05/1997
'              Weiyi Xu     15/02/1999
'
' Parameters:
' index
' vtGetData
    On Error GoTo Error
    
    Dim X       As Integer
    Dim Y       As Integer
    Dim iLen    As Integer
    Dim iIndex  As Integer
    Dim sTmp    As String
    
    iIndex = CurrentScreen
    
    X = StdScr(iIndex).Cur_x                                  '
    Y = StdScr(iIndex).Cur_y                                  '
    
    iLen = Len(vtGetData)                                     'output len
    If Y + iLen > MAXCOL Then   'no data allowed beyond 20th column
       iLen = MAXCOL - Y                                      'fit screen only
    End If
    
    'WX 15/02/1999
    'Re-enforce "no data allowed beyond 20th column" and avoid invalid procedure call error
    
    If iLen <= 0 Then Exit Sub
    
   'sTmp = UCase(CStr(vtGetData))                             'echo upper case
    'FM 2-11-98 to allow upper or lower case entry as input
    If bCase = True Then
       sTmp = Mid$(CStr(vtGetData), 1, iLen)                           'echo upper case
    Else
       sTmp = Mid$(UCase(CStr(vtGetData)), 1, iLen)
    End If
    frmMain.TCP1(index).SendData sTmp                         'echo it back
    StdScr(iIndex).Cur_y = StdScr(iIndex).Cur_y + iLen         '
    'Mid$(StdScr(iIndex).Scr(x), y) = Mid$(sTmp, 1, iLen) 'store in buffer
    Mid$(StdScr(iIndex).Scr(X), Y) = sTmp 'store in buffer
    
Exit Sub
Error:
    HandleErrorFatal "clsMax", "EchoToClient", err.Source, err.Number, err.Description

End Sub
Private Function sGetParentDir(sCurr As String) As String
'Creator:   Dean Lane 11/9/97   Find parent directory of current directory
'Modified   Dean Lane 12/9/97   Removed 'On Error' call to allow raising error upon return from this module

    Dim iLast As Integer
    Dim iPos As Integer
    
    iPos = InStr(sCurr, "\")
    While iPos > 0
        iLast = iPos
        iPos = InStr(iLast + 1, sCurr, "\")
    Wend
    sGetParentDir = Left$(sCurr, iLast - 1)
        
End Function
Private Sub LogMaxError(sModule As String, sRoutine As String, _
                        sSource As String, sErrNo As String, sErrDesc As String)
'jg 15 Oct 1999 clsMax requires own copy of this routine:  altered slightly
'   to avoid reliance on objects which may not be present in every project containing
'   clsMax

    Dim sMsgtxt      As String
    Dim sLogDir      As String
    Dim iFileNumber  As Integer
    Dim sComputerName As String
    Dim sMaxErrorLogFile As String
    'can't assume every project using clsMax has clsUtilsDB
    'so get computername here
    sComputerName = Environ$("COMPUTERNAME")

    ' Build message without cr/lf for logging
    sMsgtxt = Now() & " " & sComputerName & " Module: " & sModule & _
              ", Routine: " & sRoutine & _
              ", Source: " & sSource & " App inst: " & CStr(App.hInstance) & _
              ", ErrorNumber: " & sErrNo & _
              ", ErrorDescription: " & sErrDesc
    
    iFileNumber = FreeFile
    

    sLogDir = sGetParentDir(App.Path)  'DL
    sMaxErrorLogFile = sLogDir & "\logs\" & App.EXEName & ".log"
    Open sMaxErrorLogFile For Output As #iFileNumber
    
    Print #iFileNumber, sMsgtxt
    Close #iFileNumber
    
End Sub


Public Function NewScr() As Integer
' FUnction to get a new screen and return the new screen's number
'
' Modified By: Pauline Tran 06/05/1997

    On Error GoTo Error
    
    Dim i As Integer
    Dim iCnt As Integer
 
'===========================
  'jg 19 Oct 1999: The log file was incorrectly getting "No slots available" error instead
  ' when this routine is called from DlgBox in UnExpectedErrorFatal after an
  ' error occurs too early in Sub Main for the miScrCnt array to be set to -1
        '
    If miScrCnt(1) = 0 Then
        'don't use HandleErrorFatal as results in circular errors.
        LogMaxError "ClsMax", "Newscr", "", "0", "Screen array has not been initialised"
        'HandleErrorFatal "ClsMax", "Newscr", "", "100", "Initialisation not complete"
        End
    End If
'===========================
    
    For i = 1 To MAXSCR
        If miScrCnt(i) = -1 Then
            miScrCnt(i) = 1     'allocated for use
            Exit For
        End If
    Next i

    If i > MAXSCR Then                  'Use up all max=MAXSCR
        HandleErrorFatal "ClsMax", "Newscr", "", "100", "No slots available"
        End
    End If
    
    iCnt = i
    
    For i = 1 To MAXROW
        StdScr(iCnt).Scr(i) = msSpace20    'initialize buf
    Next i
    
    NewScr = iCnt
 
Exit Function
Error:
    HandleErrorFatal "clsMax", "NewScr", err.Source, err.Number, err.Description

End Function

Public Function CloseScr()
' Function to close the current screen by re-initialising it in the screen array
'
' Modified By: Pauline Tran 06/05/1997
 
    On Error GoTo Error
    
    miScrCnt(CurrentScreen) = -1   'make it available for reuse

Exit Function
Error:
    HandleErrorFatal "clsMax", "CloseScr", err.Source, err.Number, err.Description

End Function

' This function is no longer required... Delete when ready!!!!
Private Function ScrollScreenDownX(cln As Variant, iMaxDisp As Integer, bDispField As Boolean) As Boolean
' This function gives the effect that the screen is scrolling 1 down

    On Error GoTo Error
    
    If ClnIndex = cln.Count Then
        MsgLine "At last record."
        PrintStr 2 + CurrentIndex, 1, " "  'clear N or M
        AcceptStr 2 + CurrentIndex, 1, ""  'display cursor
        ScrollScreenDownX = False
        Exit Function
        
    ElseIf CurrentIndex >= iMaxDisp Then
        CurrentIndex = iMaxDisp
        ClnIndex = ClnIndex + 1
        DisplayList cln, False, ClnIndex - (iMaxDisp - 1), ClnIndex, bDispField
        
    Else
        PrintStr 2 + CurrentIndex, 1, " "  'clear N or M
        CurrentIndex = CurrentIndex + 1
        ClnIndex = ClnIndex + 1
        AcceptStr 2 + CurrentIndex, 1, ""    'display cursor
    End If
    
    ScrollScreenDownX = True
    
Exit Function
Error:
    HandleErrorFatal "clsMax", "ScrollScreenDownX", err.Source, err.Number, err.Description

End Function

Public Function ScrollDown(iCount As Integer, Optional iCursorPos, Optional sDefaultDisp) As Boolean
' function to reset the indexes to allow the effect of scrolling
' down one record in a given list
'
' Modified By: Pauline Tran 06/05/1997
'
' Parameters:
' iCount holds the Max number of records in the control (rs or cln)
' iCursorPos = where the cursor will land after scrolling, is 1 if missing
' sDefaultDisp = the default data to be displayed if any

    On Error GoTo Error
    
    If IsMissing(iCursorPos) Then   'set cursor pos = 1 when missing
        iCursorPos = 1
    End If
    
    If IsMissing(sDefaultDisp) Then
        sDefaultDisp = ""
    End If
    
    ScrollDown = False

    If ClnIndex = iCount Then               ' at last record
        MsgLine "At last record."
        PrintStr 2 + CurrentIndex, CInt(iCursorPos), " "  'clear N or M
        AcceptStr 2 + CurrentIndex, CInt(iCursorPos), sDefaultDisp  'display cursor
    ElseIf CurrentIndex >= miMaxDisplayLines Then  ' scroll down one
        CurrentIndex = miMaxDisplayLines           ' and re-display list
        ClnIndex = ClnIndex + 1
        ScrollDown = True
    Else                                    ' move cursor down one
        PrintStr 2 + CurrentIndex, CInt(iCursorPos), " "  'clear N or M
        CurrentIndex = CurrentIndex + 1
        ClnIndex = ClnIndex + 1
        AcceptStr 2 + CurrentIndex, CInt(iCursorPos), sDefaultDisp    'display cursor
    End If
    
    Exit Function
Error:
    HandleErrorFatal "clsMax", "ScrollDown", err.Source, err.Number, err.Description

End Function

Public Function ScrollUp(Optional iCursorPos, Optional sDefaultDisp) As Boolean
' This function gives the effect that the screen is scrolling 1 up
'
' Modified By: Pauline Tran 06/05/1997
'
' Parameters:
' iCursorPos = where the cursor will land after scrolling, is 1 if missing
' sDefaultDisp = the default data to be displayed if any

    On Error GoTo Error
    
    If IsMissing(iCursorPos) Then
        iCursorPos = 1
    End If
    
    If IsMissing(sDefaultDisp) Then
        sDefaultDisp = ""
    End If
    
    ScrollUp = False

    If ClnIndex = 1 Then            ' Already at first record.
        MsgLine "At first record."
        PrintStr 2 + CurrentIndex, CInt(iCursorPos), " "           'clear N or M
        AcceptStr 2 + CurrentIndex, CInt(iCursorPos), sDefaultDisp 'display cursor
    ElseIf CurrentIndex = 1 Then    ' scroll up one record
        ClnIndex = ClnIndex - 1     ' and re-display list
        ScrollUp = True
    Else                            ' move cursor up one
        PrintStr 2 + CurrentIndex, CInt(iCursorPos), " "  'clear N or M
        CurrentIndex = CurrentIndex - 1
        ClnIndex = ClnIndex - 1
        AcceptStr 2 + CurrentIndex, CInt(iCursorPos), sDefaultDisp 'display cursor
    End If
    
    Exit Function
Error:
    HandleErrorFatal "clsMax", "ScrollUp", err.Source, err.Number, err.Description
        
End Function

Public Function ScrollScr(sSeqKey As String, _
                          cln As Variant, _
                          bDispField As Boolean, _
                          Optional bCloseScreen, _
                          Optional bDispList, _
                          Optional sRawdata As String) As String
' Function to allow the effect of a list of records scrolling on the MAX screen
'
' Modified By: Pauline Tran 06/05/1997
'
' Parameters:
'sSeqKey is the function key pressed
'cln is the collection to be display
'bdispfield = true means cln.listfield must be present
'CloseScreen = false means screen will not be close on Enter or Quit
'bDispList present means Calls display list function.

    On Error GoTo Error
    
    Static bListDisplayed As Boolean
    
    ScrollScr = ""
    If IsMissing(bCloseScreen) Then
        bCloseScreen = True             'Close screen after scrolling
    End If
         
    If Not IsMissing(bDispList) Then
        bListDisplayed = False     'reset static variable to DisplayList
    End If
    
    If Not bListDisplayed Then
        
        If sSeqKey <> KEYLIST Then
          HandleErrorFatal "ClsMax", "ScrollScr", "KEYLIST", "0", "Not List Keys passed in"
        End If

        bListDisplayed = True
        DisplayList cln, True, 1, miMaxDisplayLines, bDispField
        mbSavedTerminatorCheck = TerminatorCheck      'saved it
        mbSavedOverStrike = OverStrike
        TerminatorCheck = False                     'can use N M
        
    End If

    ClearMsgLine

    If sSeqKey = "N" Then
        sSeqKey = KEYDOWN
    Else
        If sSeqKey = "M" Then
            sSeqKey = KEYUP
        End If
    End If
    'FM 14/10/98 the if statement below has been added to fix the problem of scanning in list screen
    If sSeqKey = KEYRETURN And sRawdata <> "" Then
      sRawdata = ""
      CloseScr                        'close current screen
      CurrentScreen = PreviousScreen  'restore previous screen
      DisplayList cln, True, 1, miMaxDisplayLines, bDispField 'display the list again
      mbSavedTerminatorCheck = TerminatorCheck
      mbSavedOverStrike = OverStrike
      TerminatorCheck = False
      Exit Function
    End If
    Select Case sSeqKey
    Case KEYDOWN
        If ScrollDown(cln.Count) Then
            DisplayList cln, False, ClnIndex - 3, ClnIndex, bDispField
        End If
    Case KEYUP
        If ScrollUp Then
            DisplayList cln, False, ClnIndex, ClnIndex + 3, bDispField
        End If
    Case KEYRETURN, KEYQUIT
                                          'redisplay screen
        ScrollScr = sSeqKey               'return the key press
        If bCloseScreen Then
            CloseScr                      'close current screen
            CurrentScreen = PreviousScreen 'restore previous screen
            bListDisplayed = False
        End If
    
        TerminatorCheck = mbSavedTerminatorCheck 'restore it
        OverStrike = mbSavedOverStrike           'restore it
    Case KEYLIST
    '
    Case Else
        MsgLine "Invalid Key"
        AcceptStr CurrentIndex + 2, 1, " "
    End Select
    
    Exit Function
Error:
    HandleErrorFatal "clsMax", "ScrollScr", err.Source, err.Number, err.Description
    
End Function

Public Function MultiSelectScrollScr(sSeqKey As String, _
                                     cln As Variant, _
                                     bDispField As Boolean, _
                                     Optional CloseScreen, _
                                     Optional vShowInvalid) As String
' Function to allow the effect of a list of records scrolling on the MAX screen
' for lists allowing multiple select
'
' Modified By: Pauline Tran 06/05/1997
'
' Parameters:
'sseqkey is the function key press
'cln  collection to be display
'bdispfield =true means cln.listfield must be present
'CloseScreen = false means screen will not be close on Enter or Quit

    On Error GoTo Error
    
    Static sDisplayList  As String

    MultiSelectScrollScr = ""
    
    If IsMissing(CloseScreen) Then
        CloseScreen = True             'Close screen after scroll
    End If
    If IsMissing(vShowInvalid) Then
        vShowInvalid = False
    End If
    
    If sDisplayList <> "Y" Then
        
        sDisplayList = "Y"
        DisplayList cln, True, 1, miMaxDisplayLines, bDispField
        mbSavedOverStrike = OverStrike
        mbSavedTerminatorCheck = TerminatorCheck      'saved it
        TerminatorCheck = False                       'can use N M
        
    End If
    
    ClearMsgLine
    
    If sSeqKey = "N" Then
        sSeqKey = KEYDOWN
    Else
        If sSeqKey = "M" Then
            sSeqKey = KEYUP
        End If
    End If
   
    Select Case sSeqKey
    Case KEYDOWN
        If ScrollDown(cln.Count) Then
            DisplayList cln, False, ClnIndex - 3, ClnIndex, bDispField
        End If
    Case KEYUP
        If ScrollUp Then
            DisplayList cln, False, ClnIndex, ClnIndex + 3, bDispField
        End If
    Case KEYRETURN
        MultiSelectScrollScr = sSeqKey           'return the key press
    Case KEYQUIT
        
        MultiSelectScrollScr = sSeqKey           'return the key press
        
        If CloseScreen Then
            CloseScr                      'close current screen
            CurrentScreen = PreviousScreen 'restore previous screen
        End If
        
        sDisplayList = "N"             'redisplay screen
        OverStrike = mbSavedOverStrike
        TerminatorCheck = mbSavedTerminatorCheck 'restore it
    Case KEYLIST
        '
    Case Else
        If vShowInvalid Then  'DL
            '**** Alex S, sep 99,
            '***** iivalid ref to itself (what if there are many instances?), should be Me
            '*****WAS:
            '***clsMax.MsgLine "Invalid Key"
            '***clsMax.AcceptStr clsMax.CurrentIndex + 2, 1, " "
        
            Me.MsgLine "Invalid Key"
            Me.AcceptStr Me.CurrentIndex + 2, 1, " "
            '*************END MOD***************
            
        End If
 
    End Select
    
    Exit Function
Error:
    HandleErrorFatal "clsMax", "MultiSelectScrollScr", err.Source, err.Number, err.Description
    
End Function

Public Sub DisplayList(cln As Variant, bNewScreen As Boolean, _
                       iStart As Integer, iEnd As Integer, bDispField As Boolean)
' Display the List on screen from the collection argument
'
' Modified By: Pauline Tran 06/05/1997
'
' Parameters:
' cln
' bNewScreen
' iStart
' iEnd
' bDispField

    On Error GoTo Error
    
    Dim i As Integer
    Dim iDisp As Integer
    
    iDisp = 3
      
    If bNewScreen Then
        PreviousScreen = CurrentScreen
        CurrentScreen = NewScr                'get a new screen id
        CurrentIndex = 1
        ClnIndex = CurrentIndex
        PrintTitle ScrollTitle                  ' prints title for scroll page
    End If
                        
    For i = iStart To iEnd         'display 4 rows from start to end
        If i > cln.Count Then      'stop if displaying more than cln
            Exit For
        End If
        
        ClearAfter iDisp, 1
        
        If bDispField Then
            PrintStr iDisp, 2, Trim(Mid$(cln(i).ListField, 1, 20 - 2))
        Else
            PrintStr iDisp, 2, Trim(Mid$(cln(i), 1, 20 - 2))
        End If
        
        iDisp = iDisp + 1
    Next i
    
    OverStrike = True
    AcceptStr CurrentIndex + 2, 1, " "   'set to first row

    Exit Sub
Error:
    HandleErrorFatal "clsMax", "DisplayList", err.Source, err.Number, err.Description

End Sub

Public Sub DisplayRsList(rs As rdoResultset, _
                         iLineToPrintFirstRecord As Integer, _
                         iRecordFirst As Integer, iRecordLast As Integer, _
                         iCursorPos As Integer, vColumns As Variant)
' Display the List on screen from the resultset argument
'
' Parameters:
' rs = the resultset containing all relevant data
' iLineToPrintFirstRecord specifies the line on Max to display the record specified by iRecordFirst
' iRecordFirst & iRecordLast = shows numbers of rows to be displayed
' iCursorPos = the position of the cursor after the list have been displayed
' vColumns = specifies the field name, size and leading spaces to be displayed in columns
'                    as follows:
'                    vColumns = Array(field_name, size, leading_spaces...)
'                    vColumns = Array("item_id", 7, 1, "item_name", 8, 2) will display
'                    7 characters of item_id starting at 2nd position and 8 characters of item_name
'                    starting at 11th position

    On Error GoTo Error
    
    Dim iRecord         As Integer
    Dim sStr            As String
    Dim iColumn         As Integer
    Dim intColumnCount  As Integer
    Dim intColumnSize   As Integer
    Dim sColumnValue    As String
    Dim iLineToPrintNextRecord As Integer
   
    iLineToPrintNextRecord = iLineToPrintFirstRecord
    
    intColumnCount = UBound(vColumns) \ 3
    
    rs.AbsolutePosition = iRecordFirst
    
    For iRecord = iRecordFirst To iRecordLast         'display 4 rows from start to end
        ClearAfter iLineToPrintNextRecord, 1
        
        If iRecord <= rs.RowCount Then   'stop if displaying more than cln
        
            sStr = ""
            
            For iColumn = 0 To intColumnCount
                intColumnSize = CInt(vColumns(3 * (iColumn + 1) - 2))
                sColumnValue = CStr(rs.rdoColumns(CStr(vColumns(3 * (iColumn + 1) - 3))).Value)
                                
                If Len(sColumnValue) >= intColumnSize Then
                    sColumnValue = Left$(sColumnValue, intColumnSize)
                Else
                    sColumnValue = sPadString(sColumnValue, intColumnSize, " ", False)
                End If
                
                sColumnValue = Space(CInt(vColumns(3 * (iColumn + 1) - 1))) & sColumnValue
                
                sStr = sStr & sColumnValue
            Next 'iColumn
        
            PrintStr iLineToPrintNextRecord, 1, sStr
            iLineToPrintNextRecord = iLineToPrintNextRecord + 1
        
            If Not rs.EOF Then
                rs.MoveNext
            End If
            
        End If
    Next 'iRecord
    
Exit Sub
Error:
    HandleErrorFatal "clsMax", "DisplayRsList", err.Source, err.Number, err.Description

End Sub

Public Sub SaveTerminatorOverstrike()
'
' Modified By: Pauline Tran 06/05/1997

    mbSavedTerminatorCheck = TerminatorCheck      'saved it
    mbSavedOverStrike = OverStrike
    TerminatorCheck = False 'can use N M

End Sub

Public Sub NewScrollScreen()
'
' Modified By: Pauline Tran 06/05/1997
        
    On Error GoTo Error
    
    PreviousScreen = CurrentScreen
    CurrentScreen = NewScr                'get a new screen id
    CurrentIndex = 1
    ClnIndex = CurrentIndex
    PrintTitle ScrollTitle                ' prints title for scroll page

Exit Sub
Error:
    HandleErrorFatal "clsMax", "NewScrollScreen", err.Source, err.Number, err.Description

End Sub

Public Sub DisplayLine(cln As Variant, bDisplayField As Boolean)
' Display a record on screen using the Current and Collection indexes
'
' Modified By: Pauline Tran 06/05/1997
'
' Parameters:
' cln
' bDisplayField

    On Error GoTo Error
    
    Dim i As Integer
    
    ClearAfter CurrentIndex + 2, 1
        
    If bDisplayField Then
         PrintStr CurrentIndex + 2, 2, Trim(Mid$(cln(ClnIndex).ListField, 1, 20 - 2))
    Else
         PrintStr CurrentIndex + 2, 2, Trim(Mid$(cln(ClnIndex), 1, 20 - 2))
    End If

    PrintStr CurrentIndex + 2, 1, ""    'set to first row

Exit Sub
Error:
    HandleErrorFatal "clsMax", "DisplayLine", err.Source, err.Number, err.Description

End Sub

Public Sub DisplayLineStr(sStr As String, iCursorPos As Integer)
' Display a record on screen using the Current and Collection indexes
'
' Modified By: Pauline Tran 06/05/1997
'
' Parameters:
' sStr
' iCursorPos
    On Error GoTo Error
    
    Dim i As Integer
    
    ClearAfter CurrentIndex + 2, 1
    PrintStr CurrentIndex + 2, 2, Trim(Mid$(sStr, 1, 20 - 2))
    PrintStr CurrentIndex + 2, iCursorPos, ""    'set to first row

Exit Sub
Error:
    HandleErrorFatal "clsMax", "DisplayLineStr", err.Source, err.Number, err.Description

End Sub

Public Sub UnexpectedErrorFatal(sProgram As String, sMsg As String, _
                                 sSource As String, errno As String, _
                                 errdesc As String)
'
'
' Modified By: Pauline Tran 06/05/1997
'
' Parameters:
' sProgram
' sMsg
' sSource
' errno
' errdesc
    
    On Error Resume Next
    
    Static sToEnd       As String
    Static iScrCount    As Integer
    Static sErrMsg(3)   As String
    Dim sTmpMsg(3)      As String
    Dim bMore           As Boolean
    Dim i               As Integer
    
    If sToEnd <> "Y" Then
       sErrMsg(1) = sSource
       sErrMsg(2) = errno
       sErrMsg(3) = errdesc
    End If
       
    Select Case FuncKeyName
        Case KEYLEFT, "K"
             If iScrCount > 0 Then
                iScrCount = iScrCount - 1
             End If
        Case KEYRIGHT, "L"
             iScrCount = iScrCount + 1
        Case KEYRETURN
            If sToEnd = "Y" Then
                frmMain.TCP1(giSockIndex).Close     'close connect
                AbortProgram
            End If
    End Select
    
    bMore = False
    
    For i = 1 To 3
      sTmpMsg(i) = ""
      If iScrCount * 20 + 1 <= Len(sErrMsg(i)) Then  'something to  display
           sTmpMsg(i) = Mid$(sErrMsg(i), iScrCount * 20 + 1)
           bMore = True
      End If
    Next i
    
     If bMore Then
        mDebug_Log sProgram & " " & sMsg & vbCrLf & sSource & " " & errno & errdesc
        ClearMsgLine
        DlgBox "", sTmpMsg(1), sTmpMsg(2), sTmpMsg(3), MSGTYP_INFO
    Else
        MsgLine "No More..."
    End If
    
    TerminatorCheck = False
    sToEnd = "Y"
    
End Sub

Public Function Unstring(Instring As String, sSeparator As String, bFound As Boolean) As String
'returns the first unstring continuous in instr
'eg.     instring="Ab;cd;e"  1st=unstring(instring,";",bfound) ('Ab')

    On Error GoTo Error
    
    Static iStart As Integer
    Dim iLen      As Integer
    Dim iFirst    As Integer
    Dim iNext     As Integer
    Dim sStr      As String
    Dim sTmp      As String

    iStart = iStart + 1
    iLen = Len(Instring)
    
    If iStart > iLen Then
        iStart = 0
        bFound = False
        Unstring = ""
        Exit Function
    End If
    
    sStr = Mid$(Instring, iStart)
    iFirst = InStr(sStr, sSeparator)
    
    Select Case iFirst
    Case 1
    'found
        bFound = True
        sTmp = Mid$(sStr, iFirst + 2)
        iNext = InStr(sTmp, sSeparator)
        If iNext = 0 Then
            Unstring = Mid$(sStr, iFirst + Len(sSeparator))
        Else
        'esc[12;3Hdasfdfffffffffffesc[1;4H "
            Unstring = Mid$(sStr, iFirst + Len(sSeparator), iNext - 1)
        End If
        iStart = iStart + Len(Unstring) + Len(sSeparator) - 1
    Case Is > 1
        'eg. 5676esc[
        bFound = False
        Unstring = Mid$(Instring, 1, iFirst - 1)
        iStart = iStart + iFirst - 2
    Case Else
        bFound = False
        Unstring = sStr
        iStart = iStart + Len(Unstring)
    End Select

Exit Function
Error:
    HandleErrorFatal "clsMax", "Unstring", err.Source, err.Number, err.Description

End Function

Public Property Get TerminatorCheck() As Boolean
    
    TerminatorCheck = pbTerminatorCheck

End Property

Public Property Let TerminatorCheck(bNewValue As Boolean)
    
    pbTerminatorCheck = bNewValue

End Property

Public Property Get FuncKeyName() As String

    FuncKeyName = msFuncKeyNm

End Property

Public Property Let FuncKeyName(sNewValue As String)
    
    msFuncKeyNm = sNewValue

End Property

Public Property Get ScanOnCheck() As Boolean
    
    ScanOnCheck = pbScanOnCheck

End Property

Public Property Let ScanOnCheck(bNewValue As Boolean)
    
    pbScanOnCheck = bNewValue

End Property

Public Property Get InputStr() As String
'Since Everything is kept in stdscr(i) I can make full use of it
'to return what is entered by user. Note if a default is supplied
'in Acceptstr x,y,"1234" defaultvaluelen defaultcolumn will be set
     
    Dim X As Integer
    Dim Y As Integer
    
    X = StdScr(CurrentScreen).Cur_x
    Y = DefaultColumn                         'starting column
    InputStr = Trim(Mid$(StdScr(CurrentScreen).Scr(X), Y, DefaultValueLen))

End Property

Public Property Let InputStr(sNewValue As String)
    'This is code is obsolete
    
    msInputStr = sNewValue

 End Property

Public Property Get ScanType() As String
    
    ScanType = psScanType

End Property

Public Property Let ScanType(sNewValue As String)
    
    psScanType = sNewValue

End Property

Public Property Get ClnIndex() As Integer
    
    ClnIndex = piClnIndex
    
End Property

Public Property Let ClnIndex(iNewvalue As Integer)
    
    piClnIndex = iNewvalue
    
End Property

Public Property Get CurrentIndex() As Integer
'should only be local to cls
    
    CurrentIndex = piCurIndex
    
End Property

Public Property Let CurrentIndex(iNewvalue As Integer)
    
    piCurIndex = iNewvalue
    
End Property

Public Property Get DebugLogOff() As Boolean
    
    DebugLogOff = pbDebugLogOff

End Property

Public Property Let DebugLogOff(bNewValue As Boolean)
  
  pbDebugLogOff = bNewValue

End Property

Public Property Get AlphaSet() As Boolean
    
    AlphaSet = pbAlpha

End Property

Public Property Let AlphaSet(bNewValue As Boolean)
  
  pbAlpha = bNewValue

End Property

Public Property Get OverStrike() As Boolean
    
    OverStrike = pbOverStrike

End Property

Public Property Let OverStrike(bNewValue As Boolean)
  
  pbOverStrike = bNewValue

End Property

Public Property Get DefaultValue() As String
    
    DefaultValue = psDefaultValue

End Property

Public Property Let DefaultValue(sNewValue As String)

  psDefaultValue = sNewValue

End Property

Public Property Get DefaultColumn() As Integer
    
    DefaultColumn = piDefaultColumn

End Property

Public Property Let DefaultColumn(iNewvalue As Integer)

    piDefaultColumn = iNewvalue

End Property

Public Property Get ScrollTitle() As String
    
    ScrollTitle = psScrollTitle

End Property

Public Property Let ScrollTitle(sNewValue As String)

  psScrollTitle = sNewValue

End Property

Public Property Get DefaultValueLen() As Integer
    
    DefaultValueLen = piDefaultValueLen

End Property

Public Property Let DefaultValueLen(iNewvalue As Integer)

    piDefaultValueLen = iNewvalue

End Property

Public Property Get PreviousScreen() As Integer
    
    PreviousScreen = piPrvScreen
    
End Property

Public Property Let PreviousScreen(iNewvalue As Integer)
    
    piPrvScreen = iNewvalue
    
End Property

Public Property Get CurrentScreen() As Integer
    
    CurrentScreen = piCurScreen
    
End Property

Public Property Let CurrentScreen(iNewvalue As Integer)
    
    piCurScreen = iNewvalue
    
End Property

Public Property Get CurrentColumn() As Integer
    
    CurrentColumn = StdScr(CurrentScreen).Cur_y
    
End Property

'Public Property Get ScanCheckDigit() As String
    
'    ScanCheckDigit = psScanCheckDigit

'End Property

'Public Property Let ScanCheckDigit(sNewValue As String)
    
'    psScanCheckDigit = sNewValue

'End Property

Public Property Get RowText(nRow As Integer) As String

    RowText = StdScr(CurrentScreen).Scr(nRow)

End Property

Public Property Get DlgVisible() As Boolean
    
    DlgVisible = pbDlgVisible

End Property

Private Property Let DlgVisible(bNewValue As Boolean)
    
    pbDlgVisible = bNewValue

End Property

Public Sub MaxNumeric_On()

    On Error GoTo Error
    
    pbAlpha = False
    MaxSendSeq "MAXNUMERIC_ON"

Exit Sub
Error:
    HandleErrorFatal "clsMax", "MaxNumeric_On", err.Source, err.Number, err.Description

End Sub

Public Sub MaxNumeric_Off()

    On Error GoTo Error
    
    pbAlpha = True
    MaxSendSeq "MAXNUMERIC_OFF"

Exit Sub
Error:
    HandleErrorFatal "clsMax", "MaxNumeric_Off", err.Source, err.Number, err.Description

End Sub

Public Sub MaxEan_On()
    
    On Error GoTo Error
    
    MaxSendSeq "MAXEAN_ON"

Exit Sub
Error:
    HandleErrorFatal "clsMax", "MaxEan_On", err.Source, err.Number, err.Description

End Sub

Public Sub MaxEan_Off()
    
    On Error GoTo Error
    
    MaxSendSeq "MAXEAN_OFF"

Exit Sub
Error:
    HandleErrorFatal "clsMax", "MaxEan_Off", err.Source, err.Number, err.Description

End Sub

Public Sub MaxScan_On()
    
    On Error GoTo Error
    
    ScanOnCheck = True
    MaxSendSeq "MAXSCAN_ON"

Exit Sub
Error:
    HandleErrorFatal "clsMax", "MaxScan_On", err.Source, err.Number, err.Description

End Sub

Public Sub MaxScan_Off()
'No need to scan off as MaX switches to OFF after scanning
'    ScanOnCheck = False
'    MaxSendSeq "MAXSCAN_OFF"
'
End Sub

Public Sub MaxAckBell()
    
    On Error GoTo Error
    
    MaxSendSeq "MAXACKBELL"

Exit Sub
Error:
    HandleErrorFatal "clsMax", "MaxAckBell", err.Source, err.Number, err.Description

End Sub

Public Sub MaxInit()
    
    On Error GoTo Error
    
    MaxSendSeq "MAXINIT"

Exit Sub
Error:
    HandleErrorFatal "clsMax", "MaxInit", err.Source, err.Number, err.Description

End Sub

Public Sub MaxClrf_On()
    
    On Error GoTo Error
    
    MaxSendSeq "MAXCLRF_ON"

Exit Sub
Error:
    HandleErrorFatal "clsMax", "MaxClrf_On", err.Source, err.Number, err.Description

End Sub

Public Sub MaxClrf_Off()
    
    On Error GoTo Error
    
    MaxSendSeq "MAXCLRF_OFF"

Exit Sub
Error:
    HandleErrorFatal "clsMax", "MaxClrf_Off", err.Source, err.Number, err.Description

End Sub

Public Sub MaxWarnBell()
    
    On Error GoTo Error
    
    MaxSendSeq "MAXWARNBELL"

Exit Sub
Error:
    HandleErrorFatal "clsMax", "MaxWarnBell", err.Source, err.Number, err.Description

End Sub

Public Sub MaxErrBell()
    
   On Error GoTo Error
   
    MaxSendSeq "MAXERRBELL"

Exit Sub
Error:
    HandleErrorFatal "clsMax", "MaxErrBell", err.Source, err.Number, err.Description

End Sub

Private Sub MaxSendSeq(s As String)
   
   On Error GoTo Error
   
   Dim vtBuf    As Variant
   Dim sSeqKey  As String
   
   FuncMaxSeqGet s, sSeqKey             'get the escape sequence
   
   If sSeqKey = "" Then                 'not found
 
      HandleErrorFatal "ClsMax", "Maxsendseq", "SeqKey not Found", "130", s
 
   End If
   
   vtBuf = sSeqKey
   frmMain.TCP1(giSockIndex).SendData vtBuf    'send it

Exit Sub
Error:
    HandleErrorFatal "clsMax", "MaxSendSeq", err.Source, err.Number, err.Description

End Sub

Public Function EnableEncorePrinter() As Boolean
   
On Error GoTo Error
   
   EnableEncorePrinter = False
   frmMain.TCP1(giSockIndex).SendData Chr$(27) & "[?5i"
   EnableEncorePrinter = True

Exit Function
Error:
    EnableEncorePrinter = False
    HandleErrorFatal "clsMax", "EnableEncorePrinter", err.Source, err.Number, err.Description

End Function
Public Function DisableEncorePrinter() As Boolean
   
On Error GoTo Error
   
   DisableEncorePrinter = False
   frmMain.TCP1(giSockIndex).SendData Chr$(27) & "[?4i"
   DisableEncorePrinter = True

Exit Function
Error:
    DisableEncorePrinter = False
    HandleErrorFatal "clsMax", "DisableEncorePrinter", err.Source, err.Number, err.Description

End Function

Public Function SendRawData(ByRef sMsg As String) As Boolean
   
On Error GoTo Error
   
   SendRawData = False
   frmMain.TCP1(giSockIndex).SendData sMsg
   SendRawData = True

Exit Function
Error:
    SendRawData = False
    HandleErrorFatal "clsMax", "SendRawData", err.Source, err.Number, err.Description

End Function


Private Sub Class_Initialize()

    pbScanOnCheck = False              '
    pbDebugLogOff = True              'Users must specify 'DEBUG
    pbTerminatorCheck = True           '

End Sub

Public Sub DumpScr(Optional iScreen)
'Dump out what is stored in the screen array
'
    On Error GoTo Error
    
    Dim i   As Integer
  
    If IsMissing(iScreen) Then
        iScreen = CurrentScreen
    End If
   
    mDebug_Log String(5, "-") & "screen=" & iScreen & String(5, "-") 'Draws -------999------"
   
    For i = 1 To MAXROW
        mDebug_Log StdScr(iScreen).Scr(i)  'throws to log row by row
    Next i
    
Exit Sub
Error:
    HandleErrorFatal "clsMax", "DumpScr", err.Source, err.Number, err.Description

End Sub

Public Sub DumpAllScreens(Optional iScreen)
    
    On Error GoTo Error
    
    Dim i    As Integer
    Dim iScr As Integer

    mScreenLog Format$(Now, "hh:nn") & " " & String(20, "=")
    If Not IsMissing(iScreen) Then
        iScr = CurrentScreen
        mScreenLog String(5, "-") & "screen= " & iScr & " " & String(5, "-")  'Draws -------999------"
        For i = 1 To MAXROW
            mScreenLog StdScr(iScr).Scr(i)  'throws to log row by row
        Next i
    Else
        For iScr = 1 To MAXSCR
            If miScrCnt(iScr) = 1 Then
                mScreenLog String(5, "-") & "screen= " & iScr & " " & String(5, "-") 'Draws -------999------"
                For i = 1 To MAXROW
                    mScreenLog StdScr(iScr).Scr(i)  'throws to log row by row
                Next i
            End If
        Next iScr
    End If
            
Exit Sub
Error:
    HandleErrorFatal "clsMax", "DumpAllScreens", err.Source, err.Number, err.Description

End Sub


Public Sub MaxScanEnable()
   ' The escape seq has been hard coded rather than getting it from the term_seq_func table to avoid the impact on too many applications
    
   On Error GoTo Error
   Dim sScanEnableEscapeSeq As String
   
   sScanEnableEscapeSeq = Chr(27) & "R0501" & Chr(27) & "\"
    
   frmMain.TCP1(giSockIndex).SendData sScanEnableEscapeSeq    'send the escape seq to the PDE


Exit Sub
Error:
    HandleErrorFatal "clsMax", "MaxEnable", err.Source, err.Number, err.Description

End Sub

Public Sub MaxScanDisable()
   ' The escape seq has been hard coded rather than getting it from the term_seq_func table to avoid the impact on too many applications
    
   On Error GoTo Error
   Dim sScanDisableEscapeSeq As String
   
   sScanDisableEscapeSeq = Chr(27) & "R0500" & Chr(27) & "\"
    
   frmMain.TCP1(giSockIndex).SendData sScanDisableEscapeSeq   'send the escape seq to the PDE


Exit Sub
Error:
    HandleErrorFatal "clsMax", "MaxDisable", err.Source, err.Number, err.Description

End Sub

