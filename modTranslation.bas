Attribute VB_Name = "modTranslation"
Option Explicit
Public Const k_Zero = "0"
Private Const k_Alphabets = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

Public Function GetGradingNumeric(ByVal sValue As String) As String
'
'   The value of the input param will be the key to
'   access the position of the character  in the Alphabetic const
'
On Error GoTo ErrorHandler

    GetGradingNumeric = GetNumericValue(sValue)

    Exit Function
ErrorHandler:
    GetGradingNumeric = ""
    err.Raise err.Number, err.Source, err.Description

End Function


Public Function GetHighPriceIndNumeric(ByVal sValue As String) As Integer

On Error GoTo ErrorHandler

    If Trim(sValue) = "Y" Then
        GetHighPriceIndNumeric = 1
    Else
        GetHighPriceIndNumeric = 0
    End If

    Exit Function
ErrorHandler:
    GetHighPriceIndNumeric = 0
    err.Raise err.Number, err.Source, err.Description

End Function

Public Function GetReOrderFlagNumeric(ByVal sValue As String) As Integer

On Error GoTo ErrorHandler

    If Trim(sValue) = "Y" Then
        GetReOrderFlagNumeric = 1
    Else
        GetReOrderFlagNumeric = 0
    End If

    Exit Function
ErrorHandler:
    GetReOrderFlagNumeric = 0
    err.Raise err.Number, err.Source, err.Description

End Function

Public Function GetItemStatusChar(ByVal sValue As String) As String
'
'   Also known by Host as Range Flag
'
On Error GoTo ErrorHandler

    '--- 3 = Quit Cycle, 4 = Allotment, 8 = Clearance, 6 = New Line, 7 = Return

    Select Case Val(Trim(sValue))
        Case 3
            GetItemStatusChar = "QC"
        Case 4
            GetItemStatusChar = "AL"
        Case 6
            GetItemStatusChar = "NL"
        Case 7
            GetItemStatusChar = "RT"
        Case 8
            GetItemStatusChar = "CL"
        
        Case Else
            GetItemStatusChar = ""      '---Normal... Valid
            
    End Select

    Exit Function
ErrorHandler:
    GetItemStatusChar = ""
    err.Raise err.Number, err.Source, err.Description

End Function

Public Function GetBestSellerChar(ByVal sValue As String) As String

Dim sRetVal     As String

On Error GoTo ErrorHandler

    sRetVal = ""
    sValue = Trim(sValue)
    
    '----Indication of Best Seller for Basic stock
    If Val(sValue) = 3 Then 'Always =  3 for best seller
        sRetVal = "*"
    End If
     
     '----25 FEB 2008 Flow Right Project
    If Val(sValue) = 1 Then '1 to display "*"
        sRetVal = "*"
    End If


    GetBestSellerChar = sRetVal

    Exit Function
ErrorHandler:
    GetBestSellerChar = ""
    err.Raise err.Number, err.Source, err.Description

End Function


Public Function GetSourceOfSupplyNumeric(ByVal sValue As String) As String

On Error GoTo ErrorHandler
    
    GetSourceOfSupplyNumeric = GetNumericValue(sValue)

    Exit Function
ErrorHandler:
    
    GetSourceOfSupplyNumeric = ""
    err.Raise err.Number, err.Source, err.Description

End Function

Private Function GetNumericValue(ByVal p_sValue As String) As String

Dim sAlphabets As String
        
On Error GoTo ErrorHandler

    sAlphabets = k_Alphabets
    p_sValue = UCase(Trim(p_sValue))
    
    If p_sValue <> "" Then
        GetNumericValue = InStr(sAlphabets, p_sValue)
    End If

    Exit Function
ErrorHandler:
    GetNumericValue = "0"
    err.Raise err.Number, err.Source, err.Description
    
End Function

Private Function GetAlphaValue(ByVal p_sValue As String) As String

Dim sAlphabets As String
        
On Error GoTo ErrorHandler

    sAlphabets = k_Alphabets
    p_sValue = UCase(Trim(p_sValue))
    
    GetAlphaValue = Mid(sAlphabets, Val(p_sValue), 1)
    
    Exit Function
ErrorHandler:
    GetAlphaValue = ""
    err.Raise err.Number, err.Source, err.Description

End Function

