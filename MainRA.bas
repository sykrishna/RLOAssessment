Attribute VB_Name = "MainRLOAssessment"
'Kmart PDE Migration Project
'
'   Name:       <RLOAssessment>
'
Option Explicit

Public Type udtPOSPrice
    
    sKeycode                As String
    sAPN                    As String
    dblPrice                As Double
    dblPriceWithDisc        As Double
    dblDealPrice            As Double
    dblDealPriceWithDisc    As Double
    sDiscPercent            As String
    sPricingMethod          As String
    iDealQTY                As Integer
    
    sItemStatus             As String
    sKmartItemFlag          As String
    sItemDescription        As String
    
    sDealDescr              As String
    iDealItemType           As Integer
    iNoOfDealInfoPages      As Integer
    
    iNoOfCopy               As Integer
    sWifiPrinterEnabled     As String
    sWifiPrinterType        As String
    sPrinterNumber          As String
    sServerName             As String
End Type

Public clsStd               As New clsStd                 'cls used in this application to access the database
Public clsRLOAssessment     As New clsRLOAssessment       'includes functions needed for Max applications
Public clsMax               As New clsMax
Public clsBarCode           As New clsInterpretBarcode
Public clsSOH               As New clsSOH

'variables
Public gcon                 As rdoConnection
Public gPosPrice            As udtPOSPrice

Public gsCheckAddCond       As String

'UDBAccess
'Public oUDB_Odbc            As udb_ODBC.clsUDBAccess

Public giSockIndex           As Integer
Public g_max_RequestID       As Long
Public g_miState             As Integer
Public g_miPrevState         As Integer
Public g_MaxBlocking         As Boolean
Public bIPSOpen              As Boolean
Public gbShelfReadyOnLabels  As Boolean

Private mbMaxLogOpened      As Boolean
Private msMaxErrorLogFile   As String

'States
Public Const k_Max_DisplayItemPromptState = 160
Public Const k_Max_GetUIItemState = 170
Public Const k_Max_DisplayRecallItemPromptState = 180
Public Const k_Max_GetUIRecallItemState = 190
Public Const k_Max_DisplayWriteOffItemPromptState = 200
Public Const k_Max_GetUIWriteOffItemState = 210
Public Const k_Max_DisplayClearanceItemPromptState = 220
Public Const k_Max_GetUIClearanceItemState = 230
Public Const k_Max_ProcessCheck_ClothingFootwearHeaterXmas = 240
Public Const k_Max_DisplaySCMPromptState = 250
Public Const k_Max_ProcessSCMState = 260
Public Const k_Max_DisplaySCMErrorMsgState = 270
Public Const k_Max_ProcessSCMErrorMsgState = 280
Public Const k_Max_ProcessDisplayContainHazards = 290
Public Const k_Max_ProcessContainHazards = 300
Public Const k_Max_ProcessDisplaySaleableOnClearanceTrolley = 310

Public Const k_Max_EndState = 900
Public Const k_Max_ErrorState = 9999

Public Const k_Claimable = "Claimable"
Public Const k_Salvage = "Salvage"
Public Const k_ReturnReasonCode_Claimable = 1
Public Const k_ReturnReasonCode_Salvage = 3
Public Const k_DefaultDept = 95

Private Enum BarcodeType
    barunknown
    barEAN13
    barEAN8
    barEAN13Code2829                'EAN13 Codes 28 & 29
    barUPCA                         'UPCA non weight
    barUPCAWeight                   'UPCA variable weight
    barUPCE
    barKeycode                      'MGB Item
    barEAN128                       'Serial Shipping Container Code
End Enum

Public Const ERR_DRIVE_CONNECTED = 84& 'AddConnection error

Function ValidateItem_RLOAssessment(ByRef p_BarCodeStr As String, ByRef p_sErrorMessage As String) As Integer

ValidateItem_RLOAssessment = 0

        If gIPMCon Is Nothing Then
           Set gIPMCon = OpenSISDSNIPM()
        End If
        
        If ValidAddedCondition(p_BarCodeStr, k_Func_RLOAssessment, p_sErrorMessage, gsCheckAddCond) Then
            ValidateItem_RLOAssessment = 0
        Else
            ValidateItem_RLOAssessment = 2
        End If
        
Exit Function

End Function

Public Sub Main()

gCommandLineNCSSOCKParam = Command
    
    Set gcon = OpenSIM0001DSN01
    Set gIPMCon = OpenSISDSNIPM()
    
    gsCheckAddCond = clsStd.GetSetting(2, 0, "AddOrderingValidation")

    gbShelfReadyOnLabels = True
    
    Call InitializeErrorHandler
    WriteToErrorLog "----- " & Now & " " & App.EXEName & " Started -----  " + gCommandLineNCSSOCKParam
    'clsMax.mDebug_Log "DB Activated ...."
    frmMain.Form_Term_Load
    
End Sub

Public Sub AbortProgram()

Dim iCount As Integer
    
On Error Resume Next
      
    WriteToErrorLog "----- " & Now & " " & App.EXEName & " Ended   -----  " + gCommandLineNCSSOCKParam
    
    CloseConnection gcon
    
    If Not gGStoreCon Is Nothing Then
        CloseConnection gGStoreCon
    End If
    
    If Not gIPMCon Is Nothing Then
        CloseConnection gIPMCon
    End If
    
    If Not clsStd Is Nothing Then
       Set clsStd = Nothing
    End If
    
    If Not clsMax Is Nothing Then
       Set clsMax = Nothing
    End If
    
    If Not clsRLOAssessment Is Nothing Then
        Set clsRLOAssessment = Nothing
    End If
    
    If Not clsSOH Is Nothing Then
        Set clsSOH = Nothing
    End If

    End
End Sub

Public Function GetSOHFromIPS(ByVal sKeycode As String, ByRef lSOH As Long) As Boolean
'
'   Get SOH value from the IPS system
'
Dim sSQL As String
Dim Cqy As rdoQuery

On Error GoTo ErrorHandler
    '
    '  Set return Code
    '
    GetSOHFromIPS = False
    lSOH = 0
    
    Set gIPSCon = OpenIPSDSN()
    Set Cqy = New rdoQuery
    
    sSQL = "{? = call dbo.usp_IPS_GetSOHSOOSITLastReceiptDateAWC(?,?,?,?,?,?,?)}"
    
    Set Cqy = gIPSCon.CreateQuery("", sSQL)
    
    Cqy.rdoParameters(0).Direction = rdParamReturnValue
    Cqy.rdoParameters(0).Type = rdTypeINTEGER
    Cqy.rdoParameters(1).Direction = rdParamInput
    Cqy.rdoParameters(1).Type = rdTypeINTEGER
    Cqy.rdoParameters(1).Value = clsStd.sToreID
    Cqy.rdoParameters(2).Direction = rdParamInput
    Cqy.rdoParameters(2).Type = rdTypeINTEGER
    Cqy.rdoParameters(2).Value = CLng(sKeycode)
    Cqy.rdoParameters(3).Direction = rdParamOutput
    Cqy.rdoParameters(3).Type = rdTypeINTEGER
    Cqy.rdoParameters(4).Direction = rdParamOutput
    Cqy.rdoParameters(4).Type = rdTypeINTEGER
    Cqy.rdoParameters(5).Direction = rdParamOutput
    Cqy.rdoParameters(5).Type = rdTypeINTEGER
    Cqy.rdoParameters(6).Direction = rdParamOutput
    Cqy.rdoParameters(6).Type = rdTypeVARCHAR
    Cqy.rdoParameters(7).Direction = rdParamOutput
    Cqy.rdoParameters(7).Type = rdTypeDOUBLE
   
    Cqy.Execute
   
    If IsNull(Cqy.rdoParameters(3).Value) Then
        lSOH = 0
    Else
        lSOH = Cqy.rdoParameters(3).Value
    End If
    
    If Cqy.rdoParameters(0) <> 0 Then
        WriteToErrorLog clsRLOAssessment.psCreateUser + " Error in function GetSOHFromIPS. Position 1."
        GetSOHFromIPS = False
        
        If Not gIPSCon Is Nothing Then
            CloseConnection gIPSCon
        End If
        
        Set Cqy = Nothing
        Set gIPSCon = Nothing
        LogError "MainRLOAssessment", "GetSOHFromIPS", False
'       err.Raise err.Number, err.Source, err.Description
'       HandleErrorFatal "MainRLOAssessment", "GetSOHFromIPS", err.Source, err.Number, err.Description
    Else
        GetSOHFromIPS = True
    End If
    
    If Not gIPSCon Is Nothing Then
       CloseConnection gIPSCon
    End If

    Set Cqy = Nothing
    Set gIPSCon = Nothing
    
    Exit Function
ErrorHandler:
    WriteToErrorLog clsRLOAssessment.psCreateUser + " Error in function GetSOHFromIPS. Position 2."
    GetSOHFromIPS = False

    If Not gIPSCon Is Nothing Then
        CloseConnection gIPSCon
    End If
    
    Set Cqy = Nothing
    Set gIPSCon = Nothing
    LogError "MainRLOAssessment", "GetSOHFromIPS", False
'   err.Raise err.Number, err.Source, err.Description
'   HandleErrorFatal "MainRLOAssessment", "GetSOHFromIPS", err.Source, err.Number, err.Description
End Function

Public Function ValidateRLOItemStep1(ByVal sKeycode As String, ByRef iReturnReasonCode As Integer, ByVal sRLO_ReasonSubCode As String, ByRef sRLOReturnType As String, _
                                     ByRef iADJCode As Integer, ByRef sErrMessage As String, ByRef dUnitCost As Double) As Integer
'
'   Validate the Item
'
Dim sSQL As String
Dim Cqy As rdoQuery

On Error GoTo ErrorHandler
    '
    '  Set return Code
    '
    ValidateRLOItemStep1 = 0
    
    Set gIPMCon1 = OpenSISDSNIPM()
    Set Cqy = New rdoQuery

    sSQL = "{? = call dbo.usp_CML_Validate_RLO_Item_Step1(?,?,?,?,?,?,?)}"

    Set Cqy = gIPMCon1.CreateQuery("", sSQL)
    
    Cqy.rdoParameters(0).Direction = rdParamReturnValue
    Cqy.rdoParameters(0).Type = rdTypeINTEGER
    Cqy.rdoParameters(1).Direction = rdParamInput
    Cqy.rdoParameters(1).Type = rdTypeINTEGER
    Cqy.rdoParameters(1).Value = CLng(sKeycode)
    Cqy.rdoParameters(2).Direction = rdParamInput
    Cqy.rdoParameters(2).Type = rdTypeINTEGER
    Cqy.rdoParameters(2).Value = iReturnReasonCode
    Cqy.rdoParameters(3).Direction = rdParamInput
    Cqy.rdoParameters(3).Type = rdTypeVARCHAR
    Cqy.rdoParameters(3).Value = sRLO_ReasonSubCode
    Cqy.rdoParameters(4).Direction = rdParamOutput
    Cqy.rdoParameters(4).Type = rdTypeVARCHAR
    Cqy.rdoParameters(5).Direction = rdParamOutput
    Cqy.rdoParameters(5).Type = rdTypeINTEGER
    Cqy.rdoParameters(6).Direction = rdParamOutput
    Cqy.rdoParameters(6).Type = rdTypeVARCHAR
    Cqy.rdoParameters(7).Direction = rdParamOutput
    Cqy.rdoParameters(7).Type = rdTypeDOUBLE    
  
    Cqy.Execute
   
    If IsNull(Cqy.rdoParameters(4).Value) Then
        sRLOReturnType = "NULL RLOReturnType"
    Else
        sRLOReturnType = Cqy.rdoParameters(4).Value
    End If
    
    If IsNull(Cqy.rdoParameters(5).Value) Then
        iADJCode = -47
    Else
        iADJCode = Cqy.rdoParameters(5).Value
    End If
 
    If IsNull(Cqy.rdoParameters(6).Value) Then
        sErrMessage = "Error Message"
    Else
        sErrMessage = Cqy.rdoParameters(6).Value
    End If

    If IsNull(Cqy.rdoParameters(7).Value) Then
        dUnitCost = 0.0
    Else
        dUnitCost = Cqy.rdoParameters(7).Value
    End If
    
    If (Cqy.rdoParameters(0) <> 0 And Cqy.rdoParameters(0) <> 1) Then
        WriteToErrorLog clsRLOAssessment.psCreateUser + " Error in function ValidateRLOItemStep1. Position 1."
        ValidateRLOItemStep1 = 0
        
        If Not gIPMCon1 Is Nothing Then
            CloseConnection gIPMCon1
        End If
        
        Set Cqy = Nothing
        Set gIPMCon1 = Nothing
        LogError "MainRLOAssessment", "ValidateRLOItemStep1", False
        err.Raise err.Number, err.Source, err.Description
        HandleErrorFatal "MainRLOAssessment", "ValidateRLOItemStep1", err.Source, err.Number, err.Description
    Else
        If IsNull(Cqy.rdoParameters(0).Value) Then
            ValidateRLOItemStep1 = -83
        Else
            ValidateRLOItemStep1 = Cqy.rdoParameters(0).Value
        End If
    End If
    
    If Not gIPMCon1 Is Nothing Then
       CloseConnection gIPMCon1
    End If

    Set Cqy = Nothing
    Set gIPMCon1 = Nothing
    
    Exit Function
ErrorHandler:
    WriteToErrorLog clsRLOAssessment.psCreateUser + " Error in function ValidateRLOItemStep1 " + err.Description
    ValidateRLOItemStep1 = 0

    If Not gIPMCon1 Is Nothing Then
        CloseConnection gIPMCon1
    End If
    
    Set Cqy = Nothing
    Set gIPMCon1 = Nothing
    LogError "MainRLOAssessment", "ValidateRLOItemStep1", False
    err.Raise err.Number, err.Source, err.Description
    HandleErrorFatal "MainRLOAssessment", "ValidateRLOItemStep1", err.Source, err.Number, err.Description
End Function

Public Function ValidateRLOItemStep2(ByVal sKeycode As String, ByRef iReturnReasonCode As Integer, ByRef sRLOReturnType As String, _
                                     ByRef iADJCode As Integer, ByRef sErrMessage As String) As Integer
'
'   Validate the Item
'
Dim sSQL As String
Dim Cqy As rdoQuery

On Error GoTo ErrorHandler
    '
    '  Set return Code
    '
    ValidateRLOItemStep2 = 0

    Set gIPMCon1 = OpenSISDSNIPM()
    Set Cqy = New rdoQuery
    
    sSQL = "{? = call dbo.usp_CML_Validate_RLO_Item_Step2(?,?,?,?,?)}"

    Set Cqy = gIPMCon1.CreateQuery("", sSQL)

    Cqy.rdoParameters(0).Direction = rdParamReturnValue
    Cqy.rdoParameters(0).Type = rdTypeINTEGER
    Cqy.rdoParameters(1).Direction = rdParamInput
    Cqy.rdoParameters(1).Type = rdTypeINTEGER
    Cqy.rdoParameters(1).Value = CLng(sKeycode)
    Cqy.rdoParameters(2).Direction = rdParamInput
    Cqy.rdoParameters(2).Type = rdTypeINTEGER
    Cqy.rdoParameters(2).Value = iReturnReasonCode
    Cqy.rdoParameters(3).Direction = rdParamOutput
    Cqy.rdoParameters(3).Type = rdTypeVARCHAR
    Cqy.rdoParameters(4).Direction = rdParamOutput
    Cqy.rdoParameters(4).Type = rdTypeINTEGER
    Cqy.rdoParameters(5).Direction = rdParamOutput
    Cqy.rdoParameters(5).Type = rdTypeVARCHAR

    Cqy.Execute
   
    If IsNull(Cqy.rdoParameters(3).Value) Then
        sRLOReturnType = "NULL RLOReturnType"
    Else
        sRLOReturnType = Cqy.rdoParameters(3).Value
    End If
    
    If IsNull(Cqy.rdoParameters(4).Value) Then
        iADJCode = -47
    Else
        iADJCode = Cqy.rdoParameters(4).Value
    End If
 
    If IsNull(Cqy.rdoParameters(5).Value) Then
        sErrMessage = "Error Message"
    Else
        sErrMessage = Cqy.rdoParameters(5).Value
    End If
    
    If (Cqy.rdoParameters(0) <> 0 And Cqy.rdoParameters(0) <> 1) Then
        WriteToErrorLog clsRLOAssessment.psCreateUser + " Error in function ValidateRLOItemStep2. Position 1."
        ValidateRLOItemStep2 = 0
        
        If Not gIPMCon1 Is Nothing Then
            CloseConnection gIPMCon1
        End If
        
        Set Cqy = Nothing
        Set gIPMCon1 = Nothing
        LogError "MainRLOAssessment", "ValidateRLOItemStep2", False
        err.Raise err.Number, err.Source, err.Description
        HandleErrorFatal "MainRLOAssessment", "ValidateRLOItemStep2", err.Source, err.Number, err.Description
    Else
        If IsNull(Cqy.rdoParameters(0).Value) Then
            ValidateRLOItemStep2 = -83
        Else
            ValidateRLOItemStep2 = Cqy.rdoParameters(0).Value
        End If
    End If
    
    If Not gIPMCon1 Is Nothing Then
       CloseConnection gIPMCon1
    End If

    Set Cqy = Nothing
    Set gIPMCon1 = Nothing
    
    Exit Function
ErrorHandler:
    WriteToErrorLog clsRLOAssessment.psCreateUser + " Error in function ValidateRLOItemStep2. Position 2."
    ValidateRLOItemStep2 = 0

    If Not gIPMCon1 Is Nothing Then
        CloseConnection gIPMCon1
    End If
    
    Set Cqy = Nothing
    Set gIPMCon1 = Nothing
    LogError "MainRLOAssessment", "ValidateRLOItemStep2", False
    err.Raise err.Number, err.Source, err.Description
    HandleErrorFatal "MainRLOAssessment", "ValidateRLOItemStep2", err.Source, err.Number, err.Description
End Function

Public Function ValidateRLOItemStep3(ByVal sKeycode As String, ByVal sAPN As String, ByVal sRLOReturnType As String, ByVal sItem_ReturnType As String, _
                                     ByRef iReturnReasonCode As Integer, ByVal sSSCC_ID As String, ByVal sUSR_ID As String, ByRef sErrMessage As String) As Integer
'
'   Validate the Item
'
Dim sSQL As String
Dim Cqy As rdoQuery

On Error GoTo ErrorHandler
    '
    '  Set return Code
    '
    ValidateRLOItemStep3 = -1

    Set gIPMCon1 = OpenSISDSNIPM()
    Set Cqy = New rdoQuery
    
    sSQL = "{? = call dbo.usp_CML_Validate_RLO_Item_Step3(?,?,?,?,?,?,?,?)}"

    Set Cqy = gIPMCon1.CreateQuery("", sSQL)

    Cqy.rdoParameters(0).Direction = rdParamReturnValue
    Cqy.rdoParameters(0).Type = rdTypeINTEGER
                                                                    
    Cqy.rdoParameters(1).Direction = rdParamInput
    Cqy.rdoParameters(1).Type = rdTypeINTEGER
    Cqy.rdoParameters(1).Value = CLng(sKeycode)
                                                                    
    Cqy.rdoParameters(2).Direction = rdParamInput
    Cqy.rdoParameters(2).Type = rdTypeVARCHAR
    Cqy.rdoParameters(2).Value = sAPN
                                                                    
    Cqy.rdoParameters(3).Direction = rdParamInput
    Cqy.rdoParameters(3).Type = rdTypeVARCHAR
    Cqy.rdoParameters(3).Value = sRLOReturnType

    Cqy.rdoParameters(4).Direction = rdParamInput
    Cqy.rdoParameters(4).Type = rdTypeVARCHAR
    Cqy.rdoParameters(4).Value = sItem_ReturnType

    Cqy.rdoParameters(5).Direction = rdParamInput
    Cqy.rdoParameters(5).Type = rdTypeINTEGER
    Cqy.rdoParameters(5).Value = CLng(iReturnReasonCode)

    Cqy.rdoParameters(6).Direction = rdParamInput
    Cqy.rdoParameters(6).Type = rdTypeVARCHAR
    Cqy.rdoParameters(6).Value = sSSCC_ID

    Cqy.rdoParameters(7).Direction = rdParamInput
    Cqy.rdoParameters(7).Type = rdTypeVARCHAR
    Cqy.rdoParameters(7).Value = sUSR_ID

    Cqy.rdoParameters(8).Direction = rdParamOutput
    Cqy.rdoParameters(8).Type = rdTypeVARCHAR

    Cqy.Execute
 
    If IsNull(Cqy.rdoParameters(8).Value) Then
        sErrMessage = "Error Message"
    Else
        sErrMessage = Cqy.rdoParameters(8).Value
    End If
   
    ValidateRLOItemStep3 = Cqy.rdoParameters(0).Value
    
    If Not gIPMCon1 Is Nothing Then
       CloseConnection gIPMCon1
    End If

    Set Cqy = Nothing
    Set gIPMCon1 = Nothing
    
    Exit Function
ErrorHandler:
    WriteToErrorLog clsRLOAssessment.psCreateUser + " Error in function ValidateRLOItemStep3. Position 2."
    ValidateRLOItemStep3 = 0

    If Not gIPMCon1 Is Nothing Then
        CloseConnection gIPMCon1
    End If
    
    Set Cqy = Nothing
    Set gIPMCon1 = Nothing
    LogError "MainRLOAssessment", "ValidateRLOItemStep3", False
    err.Raise err.Number, err.Source, err.Description
    HandleErrorFatal "MainRLOAssessment", "ValidateRLOItemStep3", err.Source, err.Number, err.Description
End Function


