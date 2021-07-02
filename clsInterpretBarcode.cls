VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInterpretBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Class InterpretBarcode is used to interpret scanned or entered barcodes converting
' the barcodes into PLU formats consistent with PLU_ID data stored in the system database.
' Other functality is available to clean the barcode of control characters and
' calculate check digits.
'
' NOTE: The barcode printed and the actual barcode encoded on a label may differ so
' the routines to determine the type of barcode may differ depending on whether the
' barcode is scanned or manually entered.
'
' ASSUMPTIONS: The PDE unit must be setup to scan barcodes and append the barcode
' symbology which will distinguish it as a scanned PLU (use MAXSCAN_ON).
'
' Summary of Barcode Symbology Appending character definitions
'
'   A   UPCE0           H   Code 39
'   B   UPCE1           I   Discrete 2 of 5
'   C   UPCA            J   Interleaved 2 of 5
'   D   MSI             K   Code 11
'   E   EAN8            L   Code 93
'   F   EAN13           M   Code 128
'   G   Codabar
'
' public rountines:
'                       ConvertPLUtoPLUId
'                       DeterminePLUType
'                       CalculateCheckDigit
'                       CleanBarcode
'                       IsVPCBarcode
'                       ExpandUPCEtoUPCA
'
' MODIFIED BY: Nicholas Trigg/NT DATE: 07/09/1998
' DEFECT NO: 1003
' DESCRIPTION: The logic of one elseif statement in the "EnteredPLUType" method
' was incorrect. It stated that a PLU of any length that preceded by "29" was
' of type EAN13. However this PLU must be 13 digits long first.
' SOLUTION: Added a nested if in the elseif and allowed the prefix checked to
' occur in the nested if, whereas the test to see if the PLU is of length 13 is
' in the elseif.
'
' MODIFIED BY: Nicholas Trigg/NT DATE: 07/09/1998
' DEFECT NO: Not logged as SQA problem
' DESCRIPTION: There needs to be logerror calls at the end of each function
' call in the clsInterpretBarcode class. Also there needs to be proper on
' error goto labels
' SOLUTION: Inclusion of logerror at the end of each function and correct
' on error goto lables.
'
Private Const APPEND_CHAR As String = ","

'Define the type of barcodes this class can interpret
'Barcode Symbology Constants used by the PDE units
Private Const UPCE_SYMBOLOGY As String = "A"
Private Const UPCA_SYMBOLOGY As String = "C"
Private Const EAN8_SYMBOLOGY As String = "E"
Private Const EAN13_SYMBOLOGY As String = "F"
Private Const EAN128_SYMBOLOGY As String = "M"

'The length of the barcode type when scanned
Private Const LEN_EAN13 As Integer = 13
Private Const LEN_EAN8 As Integer = 8
Private Const LEN_UPCA As Integer = 12
Private Const LEN_UPCE As Integer = 6
'The length of the barcodes printed that differ from the that which is scanned
Private Const LABEL_LEN_UPCA_WGT As Integer = 13
Private Const LABEL_LEN_UPCE As Integer = 8

'The prefix of the barcode type when scanned
Private Const PREFIX_EAN13 As String = "93"
Private Const PREFIX_EAN8 As String = "93"
Private Const PREFIX_EAN28 As String = "28"
Private Const PREFIX_EAN29 As String = "29"
Private Const PREFIX_UPCA As String = "0"
Private Const PREFIX_UPCA_WGT As String = "2"
'The prefix of the barcodes printed that differ from the that which is scanned
Private Const LABEL_PREFIX_UPCA_WGT As String = "02"
Private Const LABEL_PREFIX_UPCE As String = "0"

'The number of digits converted to zero from the right of the barcode number.
Private Const EAN28_29_ZERO_DIGITS As Integer = 6
Private Const UPCA_WGT_ZERO_DIGITS As Integer = 6

'The barcode types
Public Enum Barcode_Types
    barunknown
    barEAN13
    barEAN8
    barEAN13Code2829                'EAN13 Codes 28 & 29
    barUPCA                         'UPCA non weight
    barUPCAWeight                   'UPCA variable weight
    barUPCE
    barKeycode                      'MGB keycode
    barEAN128                       'Serial Shipping Container Code
End Enum

Public Function CleanBarcode(ByVal sBarcode As String) As String
 'This routine 'cleans' a barcode by stripping the appended symbology code returning
 ' the numeric barcode scanned.
    
    On Error GoTo Err_CleanBarcode

    Dim nIndex As Integer
''  FM 20-10-98 should have been done in this way which is commented out to save looping time:
''  Dim iIndex   As Integer
''
''  iIndex = InStr(sBarcode, ",")
''  If iIndex <> 0 Then
''     CleanBarcode =Left(sRawData, iIndex - 1)
''  Else
''     CleanBarcode = sBarcode
''  End If

    For nIndex = 1 To Len(sBarcode)
     'Just grab the numeric characters - ignore the rest
        If IsNumeric(Mid(sBarcode, nIndex, 1)) Then

            CleanBarcode = CleanBarcode & Mid(sBarcode, nIndex, 1)

        End If

    Next nIndex
    
    Exit Function
    
Err_CleanBarcode:
    LogError "clsInterpretBarcode", "CleanBarcode", False
    err.Raise err.Number
    
End Function

Public Function DeterminePLUType(ByVal sPLU As String) As Integer
 ' Determine the type of barcode but first determine whether the barcode has been
 ' manually entered or scanned.
 ' Modified By: An Huynh Provide Invalid Message when a valid plu is entered
 ' followed by a non standard character Def: 991 and 999
 
     On Error GoTo Err_DeterminePLUType

    Dim iIndex As Integer
    Dim sSymbology As String
    
    iIndex = InStr(1, sPLU, APPEND_CHAR)

    If iIndex > 0 Then
        sSymbology = Mid$(sPLU, iIndex + 1, 1)
        'Scanned entry determine barcode
        DeterminePLUType = ScannedPLUType(CleanBarcode(sPLU), sSymbology)
    Else
        'Manually entry determine barcode
        If CheckValidPLU(sPLU) Then                                 '991&999AH
            DeterminePLUType = EnteredPLUType(CleanBarcode(sPLU))   '991&999AH
        Else
            DeterminePLUType = Barcode_Types.barunknown             '991&999AH
        End If
    End If
    
    Exit Function
    
Err_DeterminePLUType:
    LogError "clsInterpretBarcode", "DeterminePLUType", False
    err.Raise err.Number

End Function

Private Function EnteredPLUType(sPLU As String) As Integer
' The barcode entered manually, determine barcode type.
' Last Modified By: Nicholas Trigg 07/09/1998 DefNo:1003

    On Error GoTo err_EnteredPLUType

    'EAN13 barcode , 13digits prefix '93'
    If Len(sPLU) = LEN_EAN13 And Left(sPLU, Len(PREFIX_EAN13)) = PREFIX_EAN13 Then
        
        EnteredPLUType = barEAN13
        
    'EAN8 barcode , 8digits prefix '93'
    ElseIf Len(sPLU) = LEN_EAN8 And Left(sPLU, Len(PREFIX_EAN8)) = PREFIX_EAN8 Then
    
        EnteredPLUType = barEAN8
        
    'EAN13 Code 28&29 barcode , 13digits prefix '28' or '29'
    'NT1003
    ElseIf Len(sPLU) = LEN_EAN13 Then
        If Left(sPLU, Len(PREFIX_EAN28)) = PREFIX_EAN28 Or _
        Left(sPLU, Len(PREFIX_EAN29)) = PREFIX_EAN29 Then
                EnteredPLUType = barEAN13Code2829
        End If
        
    'UPCA 'Non Weight' barcode , 12digits prefix '0'
    ElseIf Len(sPLU) = LEN_UPCA And Left(sPLU, Len(PREFIX_UPCA)) = PREFIX_UPCA Then
    
        EnteredPLUType = barUPCA
        
    'UPCA_Wgt 'Weight' barcode , 13digits prefix '02'
    ElseIf Len(sPLU) = LABEL_LEN_UPCA_WGT And _
        Left(sPLU, Len(LABEL_PREFIX_UPCA_WGT)) = LABEL_PREFIX_UPCA_WGT Then
    
        EnteredPLUType = barUPCAWeight
        
    'UPCE barcode , 8digits prefix '0'
    ElseIf Len(sPLU) = LABEL_LEN_UPCE And _
        Left(sPLU, Len(LABEL_PREFIX_UPCE)) = LABEL_PREFIX_UPCE Then
    
        EnteredPLUType = barUPCE
        
    'EAN8 barcode , 8digits prefix not '93'
    ElseIf Len(sPLU) = LEN_EAN8 Then
    
        EnteredPLUType = barKeycode
        
    Else 'Not a recognised barcode
    
        EnteredPLUType = barunknown
    
    End If
           
    Exit Function

err_EnteredPLUType:
    LogError "clsInterpretBarcode", "EnteredPLUType", False
    err.Raise err.Number
    
End Function

Public Function ScannedPLUType(sPLU As String, sSymbol As String) As Integer
' The barcode entered via scan, determine barcode type.
    
    On Error GoTo Err_ScannedPLUType

    'EAN13 barcode-non_weight , 13digits, prefix '93',symbol F
    If sSymbol = EAN13_SYMBOLOGY And Len(sPLU) = LEN_EAN13 And _
            Left(sPLU, Len(PREFIX_EAN13)) = PREFIX_EAN13 Then
        
        ScannedPLUType = barEAN13
        
    'EAN8 barcode , 8digits, prefix '93',symbol E
    ElseIf sSymbol = EAN8_SYMBOLOGY And Len(sPLU) = LEN_EAN8 And _
            Left(sPLU, Len(PREFIX_EAN8)) = PREFIX_EAN8 Then
    
        ScannedPLUType = barEAN8
        
    'EAN13 Code-weight 28&29 barcode , 13digits , prefix '28' or '29',symbol F
    ElseIf sSymbol = EAN13_SYMBOLOGY And Len(sPLU) = LEN_EAN13 And _
        Left(sPLU, Len(PREFIX_EAN28)) = PREFIX_EAN28 Or _
        Left(sPLU, Len(PREFIX_EAN29)) = PREFIX_EAN29 Then
    
        ScannedPLUType = barEAN13Code2829
        
    'UPCA 'Non Weight' barcode , 12digits, prefix '0',symbol C
    ElseIf sSymbol = UPCA_SYMBOLOGY And Len(sPLU) = LEN_UPCA And _
        Left(sPLU, Len(PREFIX_UPCA)) = PREFIX_UPCA Then
    
        ScannedPLUType = barUPCA
        
    'UPCA_Wgt 'Weight' barcode , 12digits prefix '2',symbol C
    ElseIf sSymbol = UPCA_SYMBOLOGY And Len(sPLU) = LEN_UPCA And _
        Left(sPLU, Len(PREFIX_UPCA_WGT)) = PREFIX_UPCA_WGT Then
    
        ScannedPLUType = barUPCAWeight
        
    'UPCE barcode , 6digits prefix none,symbol A
    ElseIf sSymbol = UPCE_SYMBOLOGY And Len(sPLU) = LEN_UPCE Then
    
        ScannedPLUType = barUPCE
        
    'EAN128 barcode,symbol M
    ElseIf sSymbol = EAN128_SYMBOLOGY Then
    
        ScannedPLUType = barEAN128
        
    'EAN8 barcode , 8digits prefix not '93',symbol E
    ElseIf sSymbol = EAN8_SYMBOLOGY And Len(sPLU) = LEN_EAN8 Then
    
        ScannedPLUType = barKeycode
        
    Else 'Not a recognised barcode
    
        ScannedPLUType = barunknown
    
    End If
           
    Exit Function

Err_ScannedPLUType:
    LogError "clsInterpretBarcode", "ScannedPLUType", False
    err.Raise err.Number

End Function

Public Function ConvertPLUtoPLUId(ByVal sPLU As String) As String
 ' Convert the PLU to PLU_Id format
 
    On Error GoTo Err_ConvertPLUtoPLUId
 
    Select Case DeterminePLUType(sPLU)
    
    Case barEAN13, barEAN8, barUPCA, barEAN128, barKeycode
        'No conversion required, PLU stored as is in database
        ConvertPLUtoPLUId = CleanBarcode(sPLU)
        
    Case barEAN13Code2829 'Symbology: F-weight
        ConvertPLUtoPLUId = ConvertEAN28_29ToPLUId(CleanBarcode(sPLU))
        
    Case barUPCAWeight 'symbology:C-weight
        ConvertPLUtoPLUId = ConvertUPCAWgttoPLUId(CleanBarcode(sPLU))
        
    Case barUPCE 'symbology: A
        ConvertPLUtoPLUId = ConvertUPCEtoPLUId(CleanBarcode(sPLU))
        
    Case Else   'barUnknown return given barcode
        ConvertPLUtoPLUId = sPLU
        
    End Select
    
    Exit Function
    
Err_ConvertPLUtoPLUId:
    LogError "clsInterpretBarcode", "ConvertPLUtoPLUId", False
    err.Raise err.Number

End Function

Public Function ConvertEAN28_29ToPLUId(sPLU As String) As String
 ' Replace price check digit, price and check digit with zeros
    
    On Error GoTo Err_ConvertEAN28_29ToPLUId

    
    ConvertEAN28_29ToPLUId = Left$(sPLU, Len(sPLU) - EAN28_29_ZERO_DIGITS) & _
                            String(EAN28_29_ZERO_DIGITS, "0")

    Exit Function

Err_ConvertEAN28_29ToPLUId:
    LogError "clsInterpretBarcode", "ConvertEAN28_29ToPLUId", False
    err.Raise err.Number

End Function

Public Function ConvertUPCAWgttoPLUId(sPLU As String) As String
 ' Replace price check digit, price and check digit with zeros
    
    On Error GoTo Err_ConvertUPCAWgttoPLUId

    ConvertUPCAWgttoPLUId = Left$(sPLU, Len(sPLU) - UPCA_WGT_ZERO_DIGITS) & _
                            String(UPCA_WGT_ZERO_DIGITS, "0")
    
    Exit Function
    
Err_ConvertUPCAWgttoPLUId:
    LogError "clsInterpretBarcode", "ConvertUPCAWgttoPLUId", False
    err.Raise err.Number

End Function

Private Function ConvertUPCEtoPLUId(sPLU As String) As String
 ' Strip the check digit from barcode
 
    On Error GoTo Err_ConvertUPCEtoPLUId

    If Len(sPLU) = LABEL_LEN_UPCE Then                      '8 digit printed barcode
        ConvertUPCEtoPLUId = Left$(sPLU, Len(sPLU) - 1)     ' remove check digit
    Else
        ConvertUPCEtoPLUId = sPLU
    End If
    
    Exit Function
    
Err_ConvertUPCEtoPLUId:
    LogError "clsInterpretBarcode", "ConvertUPCEtoPLUId", False
    err.Raise err.Number

End Function

Public Function CalculateCheckDigit(sBarcode As String) As String
 ' This routine calculates the check digit for the specified barcode.
 ' Caution the check digit for UPCE PLU is not calculated but derived from the original
 ' UPCA PLU. Trying to calculate the check digit for a UPCE PLU using this routine
 ' will not be correct.
    
    On Error GoTo Err_CalculateCheckDigit
  
    Dim nIndex As Integer
    Dim nRunningTotal As Integer
    Dim nDigit As Integer
    Dim nCheckDigit As Integer
    
    sBarcode = CleanBarcode(sBarcode)
    sBarcode = Left(sBarcode, Len(sBarcode) - 1)
    
    'Begin by getting every 2nd digit from the right and multiplying it 3
    For nIndex = Len(sBarcode) To 1 Step -2
        nDigit = Val(Mid(sBarcode, nIndex, 1))
        Debug.Print sBarcode, nDigit
        nRunningTotal = nRunningTotal + (nDigit * 3)
    Next nIndex
    
    'Now add every second digit starting from the 2nd rightmost digit
    For nIndex = Len(sBarcode) - 1 To 1 Step -2
        nDigit = Val(Mid(sBarcode, nIndex, 1))
        nRunningTotal = nRunningTotal + nDigit
    Next nIndex
    
    'We have a total - get it's Mod 10 for the final result
    nCheckDigit = ((10 - (nRunningTotal Mod 10)) Mod 10)
    
    CalculateCheckDigit = CStr(nCheckDigit)
    
    Exit Function
    
Err_CalculateCheckDigit:
    LogError "clsInterpretBarcode", "CalculateCheckDigit", False
    err.Raise err.Number

End Function

Public Function CheckValidPLU(sPLU)
'An Huynh 07/09/1998 function checks the PLU for non numeric characters

    On Error GoTo Err_CheckValidPLU
    
    Dim index As Integer
    Dim length As Long
    length = Len(sPLU)
    
    For index = 1 To length
        If Not IsNumeric(Mid(sPLU, index, 1)) Then
            CheckValidPLU = False
            Exit For
        End If
        CheckValidPLU = True
     Next index
     
Exit Function

Err_CheckValidPLU:
    LogError "clsInterpretBarcode", "CheckValidPLU", False
    err.Raise err.Number
End Function


