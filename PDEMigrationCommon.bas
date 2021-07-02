Attribute VB_Name = "modPDEMigrationCommon"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Const k_Max_QTYOK = 0
Public Const k_Max_ErrorBlankQTY = 1
Public Const k_Max_ErrorNumericQTY = 2
Public Const k_Max_ErrorQTYNegative = 3
Public Const k_Max_ErrorQTYDecimal = 4
Public Const k_Max_ErrorQTYTooLarge = 5
Public Const k_Max_WarningQTYTooLarge = 6
Public Const k_Max_WarningQTYDoubleDigit = 7
Public Const k_Max_WarningQTYTooLargeValue = 100

Public gCommandLineNCSSOCKParam    As String
Public Function StripNonPrintChar(ByVal sInput As String) As String

Dim iStrLen As Integer
Dim iCount As Integer

iStrLen = Len(sInput)
StripNonPrintChar = ""

If iStrLen > 0 Then
    
    iCount = 1
    
    Do While iCount <= iStrLen
            
      If (Asc(Mid(sInput, iCount, 1)) >= 32) And (Asc(Mid(sInput, iCount, 1)) <= 95) Then
        
        StripNonPrintChar = StripNonPrintChar + Mid(sInput, iCount, 1)
      
      End If
      
     iCount = iCount + 1
     
    Loop


End If
End Function
Sub EnforceUcaseFileExt(ByVal sFilename As String)

Const k_ExtTmp = ".TMP"
Const k_ExtDat = ".DAT"

Dim iDotPos             As Integer
Dim sSourceName         As String
Dim sDestinationName    As String

sFilename = Trim(sFilename)

sSourceName = sFilename
sDestinationName = Mid(sFilename, 1, Len(sFilename) - Len(k_ExtDat)) & k_ExtTmp

If FileExist(sDestinationName) Then
   Kill sDestinationName
End If
Name sSourceName As sDestinationName 'rename filename to have .TMP

sSourceName = Trim(sDestinationName)
sDestinationName = Mid(sFilename, 1, Len(sFilename) - Len(k_ExtDat)) & k_ExtDat

If FileExist(sDestinationName) Then
   Kill sDestinationName
End If
Name sSourceName As sDestinationName 'rename filename to have .DAT

End Sub

Sub EnforceUcaseFileExtPOS(ByVal sFilename As String)

Const k_ExtTmp = ".TMP"
Const k_ExtDat = ".POS"

Dim iDotPos             As Integer
Dim sSourceName         As String
Dim sDestinationName    As String

sFilename = Trim(sFilename)

sSourceName = sFilename
sDestinationName = Mid(sFilename, 1, Len(sFilename) - Len(k_ExtDat)) & k_ExtTmp

If FileExist(sDestinationName) Then
   Kill sDestinationName
End If
Name sSourceName As sDestinationName 'rename filename to have .TMP

sSourceName = Trim(sDestinationName)
sDestinationName = Mid(sFilename, 1, Len(sFilename) - Len(k_ExtDat)) & k_ExtDat

If FileExist(sDestinationName) Then
   Kill sDestinationName
End If
Name sSourceName As sDestinationName 'rename filename to have .DAT

End Sub

Function bFileRename(ByVal sBeforeFilename As String, ByVal sAfterFilename As String) As Boolean

bFileRename = True

On Error GoTo ErrorHandler

    If FileExist(sAfterFilename) Then
       Kill sAfterFilename
    End If
    Name sBeforeFilename As UCase(Trim(sAfterFilename)) 'rename filename
    
    Exit Function
ErrorHandler:

    bFileRename = False
    
End Function

Function bFileCopy(ByVal sBeforeFilename As String, ByVal sAfterFilename As String) As Boolean

bFileCopy = True

On Error GoTo ErrorHandler

    If FileExist(sAfterFilename) Then
       Kill sAfterFilename
    End If
    FileCopy sBeforeFilename, UCase(Trim(sAfterFilename))
    
    Exit Function
ErrorHandler:

    bFileCopy = False
    
End Function


Public Function TelxonBatchFileSeqNo(ByVal piBatchNo As Long) As Long

Dim sSQL            As String
Dim rs              As rdoResultset

On Error GoTo ErrorTelxonBatchFileSeqNo
  
    TelxonBatchFileSeqNo = 0
    
    sSQL = "Set Nocount On "
    sSQL = sSQL & "Insert into Cml_Group_Capture_Control "
    sSQL = sSQL & " (Crtd_tmstmp, Batch_no) "
    sSQL = sSQL & " Values("
    sSQL = sSQL & SQLText(Format(Now, "dd mmm yyyy hh:mm:ss")) + ","
    sSQL = sSQL & piBatchNo & " )"
    sSQL = sSQL & " Select @@Identity as BatchHDR_ID "
    sSQL = sSQL & " Set Nocount Off"
    Set rs = gCon.OpenResultset(sSQL)
    
    If Not rs.EOF Then
        If Not IsNull(rs!Batchhdr_id) Then
             TelxonBatchFileSeqNo = CLng(rs!Batchhdr_id)
        End If
    End If
              
    If Not (rs Is Nothing) Then
       rs.Close
       Set rs = Nothing
    End If
    
    Exit Function
ErrorTelxonBatchFileSeqNo:
    
    HandleErrorFatal "PDEMigrationCOmmon", "TelxonBatchFileSeqNo", err.Source, err.Number, err.Description

End Function


Public Sub LogMaxError(sProgram As String, sMsg As String, _
                                sSource As String, errno As String, _
                                errdesc As String)
    
Call HandleErrorFatal(sProgram, sMsg, sSource, errno, errdesc)

Exit Sub
End Sub

Public Sub HandleErrorFatal(sProgram As String, sMsg As String, _
                                sSource As String, errno As String, _
                                errdesc As String)
'This fuctions switches using clsstd or clsmax to handle
'errors according to kind of application
     '
     '  Disable LogMaxError, use modError.bas LogError Instead
     '  Kwok Chan
     '
     'LogMaxError sProgram, sMsg, sSource, errno, errdesc  'dl 13/10/1997
     '
     Call LogError(sProgram, sMsg, False, False)
     g_miState = 9999           'ERRORSTATE
     clsMax.UnexpectedErrorFatal sProgram, sMsg, sSource, errno, errdesc
    
End Sub

Function ValidateFixtureLocn(ByRef sLocn As String) As Boolean

Dim t_str   As String
Dim iPos    As Integer

Const k_Max_StartOfFixtureLocation = 1
Const k_Max_EndOfFixtureLocation = 9999

ValidateFixtureLocn = False

iPos = InStr(1, sLocn, ",")
If iPos > 0 Then
   sLocn = Left(sLocn, iPos - 1)
End If

Select Case Len(sLocn)
    Case Is = 4
         If ChkNumeric(sLocn) And Val(sLocn) >= k_Max_StartOfFixtureLocation And _
            Val(sLocn) <= k_Max_EndOfFixtureLocation Then
            ValidateFixtureLocn = True
            Exit Function
         End If
    
    Case Is = 8
         If ChkNumeric(sLocn) And Mid(sLocn, 1, 3) = "999" Then
            t_str = sLocn
            If clsBarCode.CalculateCheckDigit(t_str) = Right(sLocn, 1) Then
               '
               '  Extract The location
               '
               sLocn = Mid(sLocn, 4, 4)
               If IsNumeric(sLocn) And Val(sLocn) >= k_Max_StartOfFixtureLocation And _
                  Val(sLocn) <= k_Max_EndOfFixtureLocation Then
                  ValidateFixtureLocn = True
                  Exit Function
               End If
            End If
         End If
End Select
Exit Function


End Function
Function ValidateSellingFloorLocn(ByRef sLocn As String) As Boolean

Dim t_str   As String
Dim iPos    As Integer

Const k_Max_StartOfSellingFloorLocation = 2000
Const k_Max_EndOfSellingFloorLocation = 6999

ValidateSellingFloorLocn = False

iPos = InStr(1, sLocn, ",")
If iPos > 0 Then
   sLocn = Left(sLocn, iPos - 1)
End If

Select Case Len(sLocn)
    Case Is = 4
         If ChkNumeric(sLocn) And Val(sLocn) >= k_Max_StartOfSellingFloorLocation And _
            Val(sLocn) <= k_Max_EndOfSellingFloorLocation Then
            ValidateSellingFloorLocn = True
            Exit Function
         End If
    
    Case Is = 8
         If ChkNumeric(sLocn) And Mid(sLocn, 1, 3) = "999" Then
            t_str = sLocn
            If clsBarCode.CalculateCheckDigit(t_str) = Right(sLocn, 1) Then
               '
               '  Extract The location
               '
               sLocn = Mid(sLocn, 4, 4)
               If IsNumeric(sLocn) And Val(sLocn) >= k_Max_StartOfSellingFloorLocation And _
                  Val(sLocn) <= k_Max_EndOfSellingFloorLocation Then
                  ValidateSellingFloorLocn = True
                  Exit Function
               End If
            End If
         End If
End Select
Exit Function


End Function

Function ValidateStockroomLocn(ByRef sLocn As String) As Boolean

Dim t_str   As String
Dim iPos    As Integer
Const k_Max_StartOfStockroomLocn = 7000

ValidateStockroomLocn = False

iPos = InStr(1, sLocn, ",")
If iPos > 0 Then
   sLocn = Left(sLocn, iPos - 1)
End If

Select Case Len(sLocn)
    Case Is = 4
         If ChkNumeric(sLocn) And Val(sLocn) >= k_Max_StartOfStockroomLocn Then
            ValidateStockroomLocn = True
            Exit Function
         End If
    
    Case Is = 8
         If ValidateFixtureLocn(sLocn) Then
            If ChkNumeric(sLocn) And Val(sLocn) >= k_Max_StartOfStockroomLocn Then
               ValidateStockroomLocn = True
               Exit Function
            End If
         End If

End Select
Exit Function
End Function
Function StripLeadingZero(ByVal sBarcode As String) As String


If Len(sBarcode) > 2 Then   ' Make sure itS LENGTH > 2
    
    Do While Mid$(sBarcode, 1, 1) = "0" And Len(sBarcode) > 2
            
       sBarcode = Mid$(sBarcode, 2, Len(sBarcode) - 1)
    
    Loop


End If

StripLeadingZero = sBarcode

Exit Function
End Function

Function ChkNumeric(ByVal aSTR As String) As Boolean
Dim iLoop1 As Integer
Dim aChar As String * 1

ChkNumeric = False
For iLoop1 = 1 To Len(aSTR)
    aChar = Mid(aSTR, iLoop1, 1)
    Select Case aChar
           Case Is = "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                ChkNumeric = True
           Case Else
                ChkNumeric = False
                Exit Function
    End Select
Next iLoop1

Exit Function
End Function

Function DecodeQTYErrorMsg(ByVal iRet As Integer) As String
               
Select Case iRet
       Case Is = k_Max_ErrorBlankQTY
            DecodeQTYErrorMsg = "Cannot be Blank"
       Case Is = k_Max_ErrorNumericQTY
            DecodeQTYErrorMsg = "Must be Numeric"
       Case Is = k_Max_ErrorQTYNegative
            DecodeQTYErrorMsg = "Cannot Be Negative"
       Case Is = k_Max_ErrorQTYDecimal
            DecodeQTYErrorMsg = "Cannot Be Decimal"
       Case Is = k_Max_ErrorQTYTooLarge
            DecodeQTYErrorMsg = "Invalid QTY"
       Case Else
            DecodeQTYErrorMsg = "Invalid QTY"
End Select

End Function

Function ValidateItem(ByRef p_BarCodeStr As String) As Boolean

Dim t_str As String
Dim iPos As Integer
Dim sBarCodeStr As String
Dim sBarCodetype As String
    
ValidateItem = False
    '
    '  Strip ",????" from Barcode String if Exist
    '
    iPos = InStr(1, p_BarCodeStr, ",")
    Select Case iPos
       Case Is = 1
            ValidateItem = False
            Exit Function
       
       Case Is > 1
            '
            '  It is a scanned Barcode
            '
            sBarCodeStr = p_BarCodeStr
            sBarCodetype = Mid(sBarCodeStr, iPos + 1, 1) 'get the type of barcode
            '
            '  Further decode the barcode Type, If it is UPC , convert to APN
            '
            Select Case sBarCodetype
                   Case Is = "A"            'barcode is A type
                        sBarCodeStr = Trim(FindActualBarcode(sBarCodeStr))
                   Case Is = "F", "E", "C"
                        'sBarcode = sBarcode
                   Case Else
                        'CheckBarcodeType = "0"
            End Select
            iPos = InStr(1, sBarCodeStr, ",")
            If iPos > 0 Then
                   p_BarCodeStr = Left(sBarCodeStr, iPos - 1)
                   p_BarCodeStr = StripLeadingZero(p_BarCodeStr)
               Else
                   p_BarCodeStr = sBarCodeStr
            End If
       
       Case Else
            '
            ' Do nothing
            '
    End Select
    
    If Not ChkNumeric(p_BarCodeStr) Then     'Don't accept if alphanumeric
        Exit Function
    End If
    
    sBarCodeStr = clsBarCode.CleanBarcode(p_BarCodeStr)
    '
    ' Should be > 1 as BarCode + CheckDigit And APN expect is only allowed to have 13 or less digits
    '
    If Len(sBarCodeStr) <= 1 Or Len(sBarCodeStr) > 13 Then
       Exit Function
    End If
    
    If Val(sBarCodeStr) < 1 Then
       '---- 0 is not a valid apn or keycode
       'Although the checkdigit is valid
        Exit Function
    End If
    
    '
    ' Validate "BarCode Check Digit"
    '
    t_str = sBarCodeStr
    If clsBarCode.CalculateCheckDigit(t_str) = Right(sBarCodeStr, 1) Then
            
            'Select Case Len(sBarCodeStr)
            '       Case Is > 8
            '           t_str = sBarCodeStr
            '           If oUDB_Odbc.ConvertAPN2KeyCode(t_str) Then
            '              ValidateItem = True
            '           End If
            '       Case Else
            '           t_str = sBarCodeStr
            '           If oUDB_Odbc.ConvertKeyCode2APN(t_str) Then
            '                    ValidateItem = True
            '              Else
            '                  If oUDB_Odbc.ConvertAPN2KeyCode(t_str) Then
            '                          ValidateItem = True
             '                    Else
             '                         '
             '                         '  Force it tio ture fo rtesting
            '                          '
            '                          ValidateItem = True
            '                  End If
            '           End If
           'End Select
           ValidateItem = True
       Else
           ValidateItem = False
    End If

    Exit Function
    
End Function

Function ValidateItemForStyle(ByRef p_BarCodeStr As String) As Boolean

Dim t_str As String
Dim iPos As Integer
Dim sBarCodeStr As String
Dim sBarCodetype As String
    
ValidateItemForStyle = False
    '
    '  Strip ",????" from Barcode String if Exist
    '
    iPos = InStr(1, p_BarCodeStr, ",")
    Select Case iPos
       Case Is = 1
            ValidateItemForStyle = False
            Exit Function
       
       Case Is > 1
            '
            '  It is a scanned Barcode 
            '
            ValidateItemForStyle = False
            Exit Function            
            
       Case Else
            '
            ' Do nothing
            '
    End Select
    
    If Not ChkNumeric(p_BarCodeStr) Then     'Don't accept if alphanumeric
        Exit Function
    End If
    
    sBarCodeStr = p_BarCodeStr

    'A style number is 7 digits or less
    
    If Len(sBarCodeStr) <= 1 Or Len(sBarCodeStr) > 7 Then
       Exit Function
    End If
    
    If Val(sBarCodeStr) < 1 Then
       '---- 0 is not a valid style
        Exit Function
    End If
    
    ValidateItemForStyle = True
    
    Exit Function
    
End Function

Public Function FindActualBarcode(ByRef sBarcode As String) As String
'created by Fariba mokarram 9/10/98
'This function converts the product barcodes to standard ones (UPC-A to UPC-E)
    
    On Error GoTo Err_FindActualBarcode
    Dim sMan        As String
    Dim sItem       As String
    Dim sMan3Dig    As String
    Dim sInput      As String
    Dim sAPN        As String
    Dim sDigit      As String
    
    If Len(sBarcode) < 2 Then
       Exit Function
    End If
        
    sInput = sBarcode
    sMan3Dig = Mid$(sInput, 6, 1)

    Select Case Val(sMan3Dig)
        Case 0, 1, 2
            sAPN = "0" & Mid$(sInput, 1, 2) & sMan3Dig & "0000" & Mid$(sInput, 3, 3)
            sItem = sAPN + "9"
            sDigit = clsBarCode.CalculateCheckDigit(sItem)
            sAPN = sAPN & Trim(sDigit)
            
        Case 3
            sAPN = "0" & Mid$(sInput, 1, 3) & "00000" & Mid$(sInput, 4, 2)
            sItem = sAPN + "9"
            sDigit = clsBarCode.CalculateCheckDigit(sItem)
            sAPN = sAPN & Trim(sDigit)
        
        Case 4
            sAPN = "0" & Mid$(sInput, 1, 4) & "00000" & Mid$(sInput, 5, 1)
            sItem = sAPN + "9"
            sDigit = clsBarCode.CalculateCheckDigit(sItem)
            sAPN = sAPN & Trim(sDigit)
        
        Case Else
            sAPN = "0" & Mid$(sInput, 1, 5) & "0000" & sMan3Dig
            sItem = sAPN + "9"
            sDigit = clsBarCode.CalculateCheckDigit(sItem)
            sAPN = sAPN & Trim(sDigit)
        
    End Select
    FindActualBarcode = Trim(Format(sAPN, "#############"))
    
    Exit Function
Err_FindActualBarcode:
    HandleErrorFatal "clsMvmntBR", "FindActualBarcode", err.Source, err.Number, err.Description

End Function


Function ValidateQTY(ByVal psQTY As String, Optional psAllowNegative As String, Optional plMaxValue As Variant) As Integer
Dim t_str As String

If IsMissing(psAllowNegative) Then
        psAllowNegative = "N"
   Else
        If Trim(UCase(psAllowNegative)) = "Y" Then
                psAllowNegative = "Y"
            Else
                psAllowNegative = "N"
        End If
End If

If IsMissing(plMaxValue) Then
        plMaxValue = k_Max_WarningQTYTooLargeValue
    Else
        plMaxValue = plMaxValue
End If



ValidateQTY = k_Max_ErrorBlankQTY
t_str = psQTY
'
' Validate QTY
'
If Trim(psQTY) = "" Then
    ValidateQTY = k_Max_ErrorBlankQTY
    Exit Function
End If
              
If Not IsNumeric(psQTY) Then    'Numeric whether the entry is numeric
    ValidateQTY = k_Max_ErrorNumericQTY
    Exit Function
End If
              
If Len(psQTY) > 4 Then    'Numeric whether the entry is numeric
    ValidateQTY = k_Max_ErrorQTYTooLarge
    Exit Function
End If
              
If InStr(1, psQTY, ".") > 0 Then
    ValidateQTY = k_Max_ErrorQTYDecimal
    Exit Function
End If

If InStr(1, psQTY, ",") > 0 Then
    ValidateQTY = k_Max_ErrorNumericQTY
    Exit Function
End If
              
If Val(psQTY) < 0 And psAllowNegative <> "Y" Then
    ValidateQTY = k_Max_ErrorQTYNegative
    Exit Function
End If
              
If psAllowNegative = "Y" Then
       t_str = Str$(Abs(Val(psQTY)))
   Else
       t_str = psQTY
End If

If Val(t_str) > plMaxValue Then
    ValidateQTY = k_Max_WarningQTYTooLarge
    Exit Function
End If
              
If CheckDoubleDigits(t_str) Then
    ValidateQTY = k_Max_WarningQTYDoubleDigit
    Exit Function
End If
              
              
ValidateQTY = k_Max_QTYOK

Exit Function
End Function


Public Function CheckDoubleDigits(sInputstr As String) As Boolean
    Dim iCount      As Integer
    Dim sLetter     As String
    Dim sNextLetter As String
    Dim iCounter    As Integer
    
    
On Error GoTo Err_CheckDoubleDigits

    iCounter = 1
    For iCount = 1 To Len(sInputstr)
        sLetter = Mid(sInputstr, iCounter, 1)          'Get one letter from word at icounter position
        iCounter = iCounter + 1
        sNextLetter = Mid(sInputstr, iCounter, 1)
        If sLetter = sNextLetter Then                  'If letter equals the next letter
           CheckDoubleDigits = True                    'double digits are entered
           Exit Function
        Else
        CheckDoubleDigits = False
        End If
    Next iCount
    
    Exit Function

Err_CheckDoubleDigits:
    HandleErrorFatal "clsCountBR", "CheckDoubleDigits", err.Source, err.Number, err.Description
End Function



Public Function ExecSQL(sSQL As String) As Integer

Dim rs As rdoResultset

ExecSQL = False

On Error GoTo Error
 
    Set rs = gCon.OpenResultset(sSQL)

    ExecSQL = True

    Exit Function

Error:
    HandleErrorFatal "pdeMigrationCommon", "ExecSQL", err.Source, err.Number, err.Description
    Exit Function
End Function

Public Function FileExist(ByVal aFile As String) As Integer
Dim Result As String
    
On Error Resume Next
    
    FileExist = False
    Result = Dir$(aFile)
    If Result <> "" Then
        FileExist = True
    End If

End Function


Public Function DecodeBatchFilename(ByVal sBatchType As String) As String
    
On Error Resume Next
    
DecodeBatchFilename = "99"

Select Case Trim(UCase(sBatchType))
   Case Is = "NO"
        DecodeBatchFilename = "05"
   Case Is = "GC"
        DecodeBatchFilename = "07"
   Case Is = "RN"
        DecodeBatchFilename = "08"
   Case Is = "SI"
        DecodeBatchFilename = "11"
   Case Is = "SO"
        DecodeBatchFilename = "12"
   Case Is = "LP"
        DecodeBatchFilename = "14"
   Case Is = "SF"
        DecodeBatchFilename = "15"
   Case Is = "ML"
        '
        '  Multi Located Items
        '
        DecodeBatchFilename = "17"
   Case Is = "RF"
        '
        '  Range Flag
        '
        DecodeBatchFilename = "18"
   Case Else
        DecodeBatchFilename = "99"
End Select
Exit Function

End Function

Public Function SQLText(ByVal aText As String) As String
    Dim i As Integer, SQLStr As String
    SQLStr = "'"
    For i = 1 To Len(aText)
        If Mid$(aText, i, 1) = "'" Then
            SQLStr = SQLStr + "''"
        Else
            SQLStr = SQLStr + Mid$(aText, i, 1)
        End If
    Next i
    SQLText = SQLStr + "'"
End Function

Public Function TransferFileToOS2(ByVal sDecodedFilename As String, ByVal psSourceFilename As String) As Boolean

Const k_TmpFilenameExt = ".TMP"
Const k_FilenameExt = ".DAT"

'Dim clsTlcCopy              As New clsTlcCopy
Dim sDestinationPath        As String
Dim sDestinationFilename    As String
Dim sSourceFilename         As String

On Error GoTo ErrorHandler

TransferFileToOS2 = False
        
    'Construct File Name
    sDestinationPath = clsStd.GetSetting(2, 0, "OS2FilesDestinationDir")   'get path from Registry
    sDestinationFilename = Trim(sDestinationPath & "\" & sDecodedFilename & k_TmpFilenameExt)
        
    Call EnforceUcaseFileExt(Trim(psSourceFilename))
        
    'Check if Temp file exist
    If FileExist(sDestinationFilename) Then
       Kill sDestinationFilename
    End If
    '
    ' Copy to OS/2
    '
    '  The Reason behind to use clsTlcCopy instead of the FileCopy Statement
    '      clstlcCopy is in-house deveoped class, it copies the file byte by byte.
    '      So the copy process is happening within the PDE Apps process.
    '      Unlike FileCopy, PDE apps has no control when it will start, when it will finish !
    '
    'clsTlcCopy.FileSource = psSourceFilename
    'clsTlcCopy.FileDest = sDestinationFilename      'its .tmp at this point
    'Call clsTlcCopy.FileCopy
    FileCopy psSourceFilename, sDestinationFilename      'its .tmp at this point
    Call Sleep(500)

    'Delete existing file then rename the tmp file
    sSourceFilename = sDestinationFilename
    sDestinationFilename = Trim(sDestinationPath & "\" & sDecodedFilename & k_FilenameExt)
    If FileExist(sDestinationFilename) Then
       Kill sDestinationFilename
    End If
    Name sSourceFilename As UCase(sDestinationFilename)
    Call Sleep(500)

    TransferFileToOS2 = True

    Exit Function
ErrorHandler:

    TransferFileToOS2 = False
    
End Function

Public Function bFullFile(ByVal sFilenameAndPath As String) As Boolean

Dim iFileNumber     As Integer
Dim sBuffer         As String
Dim sLastBuffer     As String

On Error GoTo ErrorHandler

    bFullFile = False

    sFilenameAndPath = UCase(Trim(sFilenameAndPath))
    iFileNumber = FreeFile
    
    Open sFilenameAndPath For Input As #iFileNumber
    
    Do Until EOF(iFileNumber)
        sLastBuffer = sBuffer
        Line Input #iFileNumber, sBuffer
        If Trim(sBuffer) = "" Then
            Exit Do
        End If
    Loop

    Close #iFileNumber
    
    If Not Trim(sBuffer) = "" Then
        If InStr(1, Trim(sBuffer), "T") > 0 Then
            bFullFile = True
        End If
    Else
        If InStr(1, Trim(sLastBuffer), "T") > 0 Then
            bFullFile = True
        End If
    End If
    
    Exit Function
ErrorHandler:
    bFullFile = False
    
End Function



