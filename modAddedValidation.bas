Attribute VB_Name = "modAddedValidation"
Option Explicit

Public Const k_MaxLocation = 30
Private Const k_FormatItem = "0000000000000"
Private Const k_SPParam_Stockroom = "STOCKROOM"
Private Const k_SPParam_NotStockroom = "NOTSTOCKROOM"
Private Const k_SPParam_Color = "COLOUR"
Private Const k_SPParam_Size = "SIZE"
Private Const k_SPParam_BestSeller = "BESTSELLER"
Private Const k_SPParam_ALL = "ALL"

'----PDE Migration Business Function Type
Public Const k_Func_StockroomIn = "SI"
Public Const k_Func_StockroomOUT = "SO"
Public Const k_Func_GroupCapture = "GC"
Public Const k_Func_LowPresent = "LP"
Public Const k_Func_Requisition = "RN"
Public Const k_Func_RLOAssessment = "RA"
Public Const k_Func_NormalOrders = "NO"
Public Const k_Func_StoreFeedback = "SF"
Public Const k_Func_MerchMultiLocatedItem = "ML"
Public Const k_Func_EmptyPackets = "EP"

Public Const k_Func_CapacitySide = "CS"
Public Const k_Func_CapacityEnd = "CE"
Public Const k_Func_CapacityVM = "CV"
Public Const k_Func_CapacityDeleteKeycodes = "CD"
Public Const k_Func_CapacityLocationCheck = "CL"
Public Const k_Zero = "0"

'----Invalid Item for PDE Migration Function
Public Const k_NotReplenishable_SPLValue = 999999
Public Const k_SPL_Location = "0000"

Public Const k_MessageItemNotFound = "Item not found"
Public Const k_MessageItemUnReplenishable = "Not on Replenishment"
Public Const k_MessageSPLRecordNotFound = "SPL Rec Not Found!"

Public Const k_Status_QuitCycle = 3
Public Const k_Status_Clearance = 8
Public Const k_Status_Allotment = 4
Public Const k_Status_NewLine = 6

Public Const k_OutOfStockID = 2  ' 3
Public Const k_RaincheckID = 999 ' not use anymore, 2
Public Const k_EndsAndMidwayID = 1
Public Const k_CustomerOrderID = 4

Public Const k_Auckland = 1
Public Const k_Brisbane = 2
Public Const k_Direct = 4
Public Const k_NotReOrderable = 14
Public Const k_Perth = 16
Public Const k_Sydney = 19
Public Const k_Melbourne = 23
Public Const k_Max_ValidNum = 26 '----26 char in alphabet

Public Const k_WifiPrinterZebra = "ZEBRA"
Public Const k_WifiPrinterIntermec = "INTERMEC"

'----All Fields are found in Item table
Private Type udtDetails

    sBrandDesc           As String       '----Brand Descr
    sItemDesc            As String       '----Item  DESCR
    iDeptCode            As Integer      '----MDSE_XREF_NO
    iReorder             As Integer      '----STR_ORDBL
    sSourceOfSupply      As String       '----SRC
    iItemStatus          As Integer      '----CML_ITM_STAT
    sGrading             As String       'Integer      '----CORE_STK_CD
    dPackSize            As Double       '----SELL_UNT
    dSellPrice           As Double       '----RTL_PRC
    dCostPrice           As Double       '----UNT_CST
    sMFG_Styl            As String       ' Style No.
    iRainCheck           As Integer      '----Raincheck stat
    iRecall              As Integer      '----Recall stat
    iALP                 As Integer      '----ALP stat
    iCSI                 As Integer      '----CSI stat
    iKVL                 As Integer      '----KVL stat
    iQ_LBL_Reqd          As Integer      '----Label # required
    iInSIM               As Integer      '----In SIM flag
    sReplenishmentMethod As String       '----Replenishment Method
    iShelfReady          As Integer      '----Shef Ready
    
    sRLO_ReasonCode      As String
    sRLO_ReasonSubCode   As String
    sRLO_APN             As String
End Type

'----All Fields are found in SLT207 table
Public Type udtSPL
    
    dSPL_High           As Double
    dSPL_Low            As Double
    lMin_Lspl           As Long
    lMax_Lspl           As Long
    
End Type

'----All Fields are found in Item Attribute table
Private Type udtAttributes

    sColour             As String
    sSize               As String
    sBestSeller         As String
    
End Type

'----Effective Item Price are found in Price table
Private Type udtPrice

    sRegularPrice        As String
    sTempPrice           As String

End Type

Private Type udtADJ
    sADJCode            As String
    sADJCodeAbbr        As String
    lADJSOH             As Long
    lADJStoreCount      As Long
    lADJQTY             As Long
    sSOHReqd            As String
End Type


Public Type udtItem

    sItemID                             As String
    
    bItemInXrefTbl                      As Boolean
    bItemOnFile                         As Boolean
    
    bReadAPNFlag                        As Boolean
    bReadItemDetailsFlag                As Boolean
    bReadItemAttributesFlag             As Boolean
    bReadItemLocationFlag               As Boolean
    bReadItemPriceFlag                  As Boolean
    tDetails                            As udtDetails
    tAttributes                         As udtAttributes
    sLocationID(1 To k_MaxLocation)     As String
    tSPL                                As udtSPL
    
    tPrice                              As udtPrice
    sSellingFloorLocn                   As String * 4
    tADJ                                As udtADJ
    
End Type


Public gtItemData        As udtItem
Public gColApnList       As New Collection
Public gSPL              As udtSPL
Public gGStoreCon        As rdoConnection
Public gIPMCon           As rdoConnection
Public gIPSCon           As rdoConnection
Public gIPMCon1          As rdoConnection

Public Function GStore_ExecSQL(sSQL As String) As Boolean

Dim rs As rdoResultset

GStore_ExecSQL = False

On Error GoTo Error
 
    Set rs = gGStoreCon.OpenResultset(sSQL)

    GStore_ExecSQL = True

    Exit Function

Error:
    HandleErrorFatal "pdeMigrationCommon", "GStore_ExecSQL", err.Source, err.Number, err.Description
    Exit Function
End Function


Public Function IPM_ExecSQL(sSQL As String) As Boolean

Dim rs As rdoResultset

IPM_ExecSQL = False

On Error GoTo Error
 
    Set rs = gIPMCon.OpenResultset(sSQL)

    IPM_ExecSQL = True

    Exit Function

Error:
    HandleErrorFatal "pdeMigrationCommon", "IPM_ExecSQL", err.Source, err.Number, err.Description
    Exit Function
End Function


Public Sub ItemStruct_Initialise()

Dim iCount      As Integer
    
    
    gtItemData.sItemID = "0"
    gtItemData.bItemInXrefTbl = False
    gtItemData.bItemOnFile = False
    
    gtItemData.bReadAPNFlag = False
    gtItemData.bReadItemDetailsFlag = False
    gtItemData.bReadItemAttributesFlag = False
    gtItemData.bReadItemLocationFlag = False
    gtItemData.bReadItemPriceFlag = False
    
   
    With gtItemData.tAttributes
        .sBestSeller = ""
        .sColour = ""
        .sSize = ""
    End With

    With gtItemData.tDetails
        .dPackSize = 0
        .iDeptCode = 0
        .sGrading = ""
        .iItemStatus = 0
        .iReorder = 0
        .sSourceOfSupply = ""
        .sBrandDesc = ""
        .sItemDesc = ""
        .dCostPrice = 0
        .dSellPrice = 0
        .sMFG_Styl = ""
        .iRainCheck = 0
        .iRecall = 0
        .iALP = 0
        .iCSI = 0
        .iKVL = 0
        .iInSIM = 0
        .sReplenishmentMethod = ""
        .iShelfReady = 0
        
    End With

    For iCount = 1 To k_MaxLocation
        gtItemData.sLocationID(iCount) = ""
    Next
    
    With gtItemData.tSPL
        .dSPL_High = 0
        .dSPL_Low = 0
        .lMax_Lspl = 0
        .lMin_Lspl = 0
        End With
    
    
    With gtItemData.tPrice
        .sRegularPrice = "     "
        .sTempPrice = "     "
    End With
    
    With gtItemData.tADJ
        .sADJCode = ""
        .sADJCodeAbbr = ""
        .lADJSOH = 0
        .lADJStoreCount = 0
        .lADJQTY = 0
        .sSOHReqd = " "
    End With
    
End Sub

Public Function OpenSISDSNGstore() As rdoConnection
'
' Returns connection if successful in opening the connection
'  otherwise 'nothing'
'
'   ---- Allows Connection pooling
'   ---- Should be pointing to a DSN that has the Default DB  is GSTORE
'

Dim cn As rdoConnection

On Error GoTo ErrorHandler
    
    Set cn = New rdoConnection
    
    cn.CursorDriver = rdUseOdbc
    cn.Connect = "dsn=" & INVENTORY_ODBC_4
    cn.EstablishConnection rdDriverNoPrompt
        
    Set OpenSISDSNGstore = cn
        
    Set cn = Nothing
        
    Exit Function
ErrorHandler:
    LogError "modAddedValidation", "OpenSISDSNGstore", False
    err.Raise err.Number, err.Source, err.Description
    
End Function

Private Function getRsItemDetails(p_gIPMCon As rdoConnection, ByVal p_sItemID As String) As rdoResultset
'
'   Execute Stored Procedure to get the Item Details base on the Input Param
'   Return to the Calling function a rdoResultset
'
Dim sSQL            As String

On Error GoTo ErrorHandler
  
    'Delete blank
    p_sItemID = Trim(p_sItemID)
    
    sSQL = "SELECT Q_UNIT_INNER, C_MDEPT, C_KD_GRD, C_KD_STAT, F_KD_RE_ORD, C_KD_SRC, T_ABBR_DESC_KD, T_BRAND_DESC_KD, A_PR_POS,A_CHRG_OUT_COST, F_KD_RAINCHECK , F_KD_RECALL, F_KD_ALP, F_KD_CSI, F_KD_KVL, Q_LBL_REQD, F_KD_IN_SIM, C_KD_RPLMT_MTHD, F_KD_SHELF_READY "
    sSQL = sSQL & "  FROM SLT200 WHERE M_KD = " + "'" + p_sItemID + "'"
    Set getRsItemDetails = p_gIPMCon.OpenResultset(sSQL, rdOpenStatic)
    
    Exit Function
ErrorHandler:

    LogError "modAddedValidation", "getRsItemDetails", False
    err.Raise err.Number, err.Source, err.Description

End Function
Private Function getRsItemProperties(p_gIPMCon As rdoConnection, ByVal p_sItemID As String) As rdoResultset
'
'   Execute Stored Procedure to get the Item Details base on the Input Param
'   Return to the Calling function a rdoResultset
'
Dim sSQL            As String

On Error GoTo ErrorHandler
    
  
    p_sItemID = Trim(p_sItemID)
    
    sSQL = "SELECT M_STYLE "
    sSQL = sSQL & "FROM SLT931 WHERE M_KD = " + "'" + p_sItemID + "'"
    Set getRsItemProperties = p_gIPMCon.OpenResultset(sSQL, rdOpenStatic)

    Exit Function
ErrorHandler:

    LogError "modAddedValidation", "getRsItemProperties", False
    err.Raise err.Number, err.Source, err.Description

End Function

Private Function getRsStyleProperties(p_gIPMCon As rdoConnection, ByVal p_sItemID As String) As rdoResultset
'
'   Execute Stored Procedure to get the Style Properties based on the Input Parameter
'   Return to the Calling function a rdoResultset
'
Dim sSQL            As String

On Error GoTo ErrorHandler
    
  
    p_sItemID = Trim(p_sItemID)
    
    sSQL = "SELECT T_STYLE_DESC "
    sSQL = sSQL & "FROM SLT929 WHERE M_STYLE = " + p_sItemID
    Set getRsStyleProperties = p_gIPMCon.OpenResultset(sSQL, rdOpenStatic)

    Exit Function
ErrorHandler:

    LogError "modAddedValidation", "getRsStyleProperties", False
    err.Raise err.Number, err.Source, err.Description

End Function

Public Function getRsItemLocations(p_gIPMCon As rdoConnection, ByVal p_sItemID As String) As rdoResultset
'
'   Execute Stored Procedure to get the Item Location base on the Input Param
'   Return to the Calling function a rdoResultset
'
Dim sSQL            As String

On Error GoTo ErrorHandler

    'Delete blank
    p_sItemID = Trim(p_sItemID)

    sSQL = "SELECT M_STK_RM_LOCN from slt208 Where M_KD  = " + "'" + p_sItemID + "'" + " Order by M_STK_RM_LOCN"
    
    Set getRsItemLocations = p_gIPMCon.OpenResultset(sSQL, rdOpenStatic)
   
    Exit Function
ErrorHandler:

    LogError "modAddedValidation", "getRsItemLocations", False
    err.Raise err.Number, err.Source, err.Description

End Function

Public Function getRsSPL(p_gIPMCon As rdoConnection, ByVal p_sItemID As String) As rdoResultset
'
'   Execute Stored Procedure to get the Item Location base on the Input Param
'   Return to the Calling function a rdoResultset
'
Dim sSQL            As String

On Error GoTo ErrorHandler

    'Delete blank
    p_sItemID = Trim(p_sItemID)

    sSQL = "SELECT Q_UNIT_HIGH_LEVEL, Q_UNIT_LOW_LEVEL, Q_LSPL_MIN, Q_LSPL_MAX FROM SLT207 WHERE  M_KD = " + "'" + p_sItemID + "'"
    
    Set getRsSPL = p_gIPMCon.OpenResultset(sSQL, rdOpenStatic)
   
    Exit Function
ErrorHandler:

    LogError "modAddedValidation", "getRsSPL", False
    err.Raise err.Number, err.Source, err.Description

End Function

Private Function getRsItemAttributes(p_gIPMCon As rdoConnection, ByVal p_sItemID As String, ByVal p_sRequestType As String) As rdoResultset
'
'   Execute Stored Procedure to get the Item Attributes base on the Input Param
'   Return to the Calling function a rdoResultset
'
Dim sSQL            As String

On Error GoTo ErrorHandler

    'Delete blank
    p_sItemID = Trim(p_sItemID)
    
    sSQL = "SELECT * FROM SLT200 WHERE  M_KD = " + "'" + p_sItemID + "'"
    Set getRsItemAttributes = p_gIPMCon.OpenResultset(sSQL, rdOpenStatic)
    
      
    Exit Function
ErrorHandler:

    LogError "modAddedValidation", "getRsItemAttributes", False
    err.Raise err.Number, err.Source, err.Description

End Function

Private Function getRsApn(p_gIPMCon As rdoConnection, ByVal p_sItemID As String) As rdoResultset
'
'   Execute Stored Procedure to get the APN for Item base on the Input Param
'   Return to the Calling function a rdoResultset
'
Dim sSQL            As String
Dim sItem           As String


On Error GoTo ErrorHandler

    'Delete blank
    p_sItemID = Trim(p_sItemID)

    'Try APN  First
    sSQL = "SELECT * FROM SLT201 WHERE  M_APN = " + p_sItemID + " Order by M_Apn Desc"
    Set getRsApn = p_gIPMCon.OpenResultset(sSQL, rdOpenStatic)
    If (getRsApn.BOF And getRsApn.EOF) Then
        '
        ' Empty Result Set. Then As Keycode
        '
            sSQL = "SELECT * FROM SLT201 WHERE  M_KD = " + p_sItemID + " Order by M_Apn Desc"
            Set getRsApn = p_gIPMCon.OpenResultset(sSQL, rdOpenStatic)
               
     Else
      'apn found, do another search to find biggest number of apn
      sItem = getRsApn!M_KD
      sSQL = "SELECT * FROM SLT201 WHERE  M_KD = " + sItem + " Order by M_Apn Desc"
      Set getRsApn = p_gIPMCon.OpenResultset(sSQL, rdOpenStatic)
      
         
    End If
    Exit Function

ErrorHandler:

    LogError "modAddedValidation", "getRsApn", False
    err.Raise err.Number, err.Source, err.Description

End Function

Private Function getRsItemPrice(p_gIPMCon As rdoConnection, ByVal p_sItemID As String) As rdoResultset
'
'   Execute Stored Procedure to get the Price for Item base on the Input Param
'   Return to the Calling function a rdoResultset
'
Dim sSQL            As String
Dim sTodayDate      As String

On Error GoTo ErrorHandler

    'Delete blank
    p_sItemID = Trim(p_sItemID)
    sTodayDate = Format(Now(), "YYYY-MM-DD 00:00:00")
    
    sSQL = "SELECT * FROM SLT000 WHERE (M_KD = " + "'" + p_sItemID + "')" + " AND ( '" + sTodayDate + "' >= D_EFTVE) AND ( '" _
    + sTodayDate + "'<= D_END ) Order BY C_PR_TYPE DESC , D_EFTVE "
        
    Set getRsItemPrice = p_gIPMCon.OpenResultset(sSQL, rdOpenStatic)
   
    Exit Function
ErrorHandler:

    LogError "modAddedValidation ", "getRsItemPrice", False
    err.Raise err.Number, err.Source, err.Description

End Function

Private Function ValidItemDetailsForSF(ByVal p_sItemID As String, ByVal p_sPDEBusFunction As String, ByVal p_StorefeedBackSubType As Variant, ByRef p_sDisplayErrMsg As String) As Boolean

Dim sRet            As String
'Dim sAPN            As String
Dim sItemKeycode    As String


    ValidItemDetailsForSF = True
    
    If SetApnList(p_sItemID) Then       '----Translate Scan to Keycode
        '----Take the first element to the collection
        'sAPN = gColApnList.Item(1)          '---- not required in this Pde function
        sItemKeycode = Trim(gtItemData.sItemID)
        
        If SetItemDetails(sItemKeycode) Then
            
            If Not ValidRecallFlag() Then
                ValidItemDetailsForSF = False
                p_sDisplayErrMsg = "Recall Item"
                Exit Function
            End If
            
            If Not ValidItemStatus(p_sPDEBusFunction) Then
                ValidItemDetailsForSF = False
                sRet = GetItemStatusChar(gtItemData.tDetails.iItemStatus)
                p_sDisplayErrMsg = "Item Status " & sRet
                Exit Function
            End If
            
            If Not ValidReOrderFlag() Then
                If Not ( _
                    p_StorefeedBackSubType = k_RaincheckID And (gtItemData.tDetails.iItemStatus = k_Status_Allotment Or gtItemData.tDetails.iItemStatus = k_Status_NewLine)) Then
                    ValidItemDetailsForSF = True
                    p_sDisplayErrMsg = "Cannot Reorder Item"
                    Exit Function
                End If
            End If
            
            
            
            Call SetItemAttributes(gtItemData.sItemID)
        
        Else
            '----This should never happen... eg Item exist in Item_Xref by not the Item table
            ValidItemDetailsForSF = False
            p_sDisplayErrMsg = k_MessageItemNotFound
        End If
            
    Else
        p_sDisplayErrMsg = k_MessageItemNotFound
        ValidItemDetailsForSF = False
    End If

End Function

Private Function ValidItemDetailsForML(ByVal p_sItemID As String, ByVal p_sPDEBusFunction As String, ByRef p_sDisplayErrMsg As String) As Boolean

Dim sRet            As String
Dim sAPN            As String
Dim sItemKeycode    As String
    
    If SetApnList(p_sItemID) Then           '----Translate Scan to Keycode
        
        sItemKeycode = Trim(gtItemData.sItemID)
        
        If SetItemDetails(sItemKeycode) Then
            ValidItemDetailsForML = True
        Else
            '----This should never happen... eg Item exist in Item_Xref by not the Item table
            ValidItemDetailsForML = False
            p_sDisplayErrMsg = k_MessageItemNotFound
        End If
    Else
        p_sDisplayErrMsg = k_MessageItemNotFound
        ValidItemDetailsForML = False
    End If
    

End Function

Private Function ValidItemDetailsForEP(ByVal p_sItemID As String, ByVal p_sPDEBusFunction As String, ByRef p_sDisplayErrMsg As String) As Boolean

Dim sRet            As String
Dim sAPN            As String
Dim sItemKeycode    As String

ValidItemDetailsForEP = False

    If SetApnList(p_sItemID) Then           '----Translate Scan to Keycode
        
        sItemKeycode = Trim(gtItemData.sItemID)
        
        If SetItemDetails(sItemKeycode) Then
            ValidItemDetailsForEP = True
        Else
            '----This should never happen... eg Item exist in Item_Xref by not the Item table
            ValidItemDetailsForEP = False
            p_sDisplayErrMsg = k_MessageItemNotFound
        End If
    Else
        p_sDisplayErrMsg = k_MessageItemNotFound
        ValidItemDetailsForEP = False
    End If

End Function


Private Function ValidItemDetailsForGC(ByVal p_sItemID As String, ByVal p_sPDEBusFunction As String, ByRef p_sDisplayErrMsg As String) As Boolean

Dim sRet            As String
'Dim sAPN            As String
Dim sItemKeycode    As String

    p_sDisplayErrMsg = ""
    ValidItemDetailsForGC = True

    If SetApnList(p_sItemID) Then       '----Translate Scan to Keycode
        '----For Group capture don't need to extract information just check if exist
        '----Take the first element to the collection
        'sAPN = gColApnList.Item(1)          '---- not required in this Pde function
        'sItemKeycode = Trim(gtItemData.sItemID)
    Else
        p_sDisplayErrMsg = k_MessageItemNotFound
        ValidItemDetailsForGC = False
    End If

End Function

Private Function ValidItemDetailsForRN(ByVal p_sItemID As String, ByVal p_sPDEBusFunction As String, ByRef p_sDisplayErrMsg As String) As Boolean

Dim sRet            As String
'Dim sAPN            As String
Dim sItemKeycode    As String

    p_sDisplayErrMsg = ""
    ValidItemDetailsForRN = True

    If SetApnList(p_sItemID) Then       '----Translate Scan to Keycode
        '----For Requisition don't need to extract information just check if exist
        '----Take the first element to the collection
        'sAPN = gColApnList.Item(1)          '---- not required in this Pde function
        'sItemKeycode = Trim(gtItemData.sItemID)
    Else
        p_sDisplayErrMsg = k_MessageItemNotFound
        ValidItemDetailsForRN = False
    End If

End Function

Private Function ValidItemDetailsForRA(ByVal p_sItemID As String, ByVal p_sPDEBusFunction As String, ByRef p_sDisplayErrMsg As String) As Boolean

Dim sRet            As String
Dim sItemKeycode    As String

Dim sBarCodeStr     As String
Dim sReasonCode     As String
Dim sReasonSubCode  As String

ValidItemDetailsForRA = False

    p_sDisplayErrMsg = ""
    gtItemData.tDetails.sRLO_ReasonCode = ""
    gtItemData.tDetails.sRLO_ReasonSubCode = ""
    gtItemData.tDetails.sRLO_APN = ""
    sBarCodeStr = RTrim$(LTrim$(p_sItemID))
    If Len(sBarCodeStr) = 20 And Mid(sBarCodeStr, 1, 3) = "275" Then
            gtItemData.tDetails.sRLO_ReasonCode = Mid(sBarCodeStr, 4, 1)
            gtItemData.tDetails.sRLO_ReasonSubCode = Mid(sBarCodeStr, 5, 3)
            sItemKeycode = Mid(sBarCodeStr, 8, 13)
        Else
            sItemKeycode = p_sItemID
    End If
    
    If SetApnList(sItemKeycode) Then       '----Translate Scan to Keycode
       
        If Len(sItemKeycode) > 8 Then
                gtItemData.tDetails.sRLO_APN = sItemKeycode
           Else
                If gColApnList.Count > 0 Then
                    gtItemData.tDetails.sRLO_APN = gColApnList.Item(1)
                Else
                    gtItemData.tDetails.sRLO_APN = sItemKeycode
                End If
        End If
       
        sItemKeycode = Trim(gtItemData.sItemID)
        
        If SetItemDetails(sItemKeycode) Then
            Call SetItemAttributes(gtItemData.sItemID)
            Call SetItemPrice(gtItemData.sItemID)
            Call SetItemLocations(sItemKeycode)
            ValidItemDetailsForRA = True
        Else
            '----This should never happen...
            ValidItemDetailsForRA = False
            p_sDisplayErrMsg = k_MessageItemNotFound
        End If
    Else
        p_sDisplayErrMsg = k_MessageItemNotFound
        ValidItemDetailsForRA = False
    End If

End Function

Private Function ValidItemDetailsForNO(ByVal p_sItemID As String, ByVal p_sPDEBusFunction As String, ByRef p_sDisplayErrMsg As String) As Boolean

Dim sRet            As String
'Dim sAPN            As String
Dim sItemKeycode    As String

    ValidItemDetailsForNO = True

    If SetApnList(p_sItemID) Then       '----Translate Scan to Keycode
        '----Take the first element to the collection
        'sAPN = gColApnList.Item(1)          '---- not required in this Pde function
        sItemKeycode = Trim(gtItemData.sItemID)

        If SetItemDetails(sItemKeycode) Then
            
            If Not ValidRecallFlag() Then
                ValidItemDetailsForNO = False
                p_sDisplayErrMsg = "Recall Item"
                Exit Function
            End If
            
            If Not ValidReOrderFlag() Then
                ValidItemDetailsForNO = False
                p_sDisplayErrMsg = "Cannot Reorder Item"
                Exit Function
            End If
                        
            If Not ValidItemStatus(p_sPDEBusFunction) Then
                ValidItemDetailsForNO = False
                sRet = GetItemStatusChar(Trim(gtItemData.tDetails.iItemStatus))
                p_sDisplayErrMsg = "Item Status " & sRet
                Exit Function
            End If
        
        Else
            '----This should never happen... eg Item exist in Item_Xref by not the Item table
            ValidItemDetailsForNO = False
            p_sDisplayErrMsg = k_MessageItemNotFound
        End If
        
    Else
        p_sDisplayErrMsg = k_MessageItemNotFound
        ValidItemDetailsForNO = False
    End If

End Function

Private Function ValidItemDetailsForLP(ByVal p_sItemID As String, ByVal p_sPDEBusFunction As String, ByRef p_sDisplayErrMsg As String) As Boolean

Dim sRet            As String
'Dim sAPN            As String
Dim sItemKeycode    As String
Dim iLoop1          As Integer

    ValidItemDetailsForLP = True

    If SetApnList(p_sItemID) Then       '----Translate Scan to Keycode
        '----Take the first element to the collection
        'sAPN = gColApnList.Item(1)          '---- not required in this Pde function
        sItemKeycode = Trim(gtItemData.sItemID)

        If SetItemDetails(sItemKeycode) Then
            
            If Not ValidRecallFlag() Then
                ValidItemDetailsForLP = False
                p_sDisplayErrMsg = "Recall Item"
                Exit Function
            End If
            
            If Not ValidItemStatus(p_sPDEBusFunction) Then
                
                ValidItemDetailsForLP = False
                sRet = GetItemStatusChar(Trim(gtItemData.tDetails.iItemStatus))
                p_sDisplayErrMsg = "Item Status " & sRet
                Exit Function
            End If
            '
            '  Read All Stock_on_Hand details for this Keycode
            '
            gSPL.dSPL_High = k_NotReplenishable_SPLValue
            gSPL.dSPL_Low = k_NotReplenishable_SPLValue
            gSPL.lMax_Lspl = k_NotReplenishable_SPLValue
            gSPL.lMin_Lspl = k_NotReplenishable_SPLValue
            
            If SetItemSPL(gtItemData.sItemID) Then
                       
                gSPL.dSPL_Low = gtItemData.tSPL.dSPL_Low
                gSPL.dSPL_High = gtItemData.tSPL.dSPL_High
                gSPL.lMin_Lspl = gtItemData.tSPL.lMin_Lspl
                gSPL.lMax_Lspl = gtItemData.tSPL.lMax_Lspl
                
                If gSPL.lMin_Lspl = k_NotReplenishable_SPLValue Or _
                   gSPL.lMax_Lspl = k_NotReplenishable_SPLValue Or _
                   gSPL.lMin_Lspl < 0 Or _
                   gSPL.lMax_Lspl < 0 Or _
                   Abs(gSPL.dSPL_Low - k_NotReplenishable_SPLValue) < 0.01 Then
                         p_sDisplayErrMsg = k_MessageItemUnReplenishable
                         ValidItemDetailsForLP = False
                         Exit Function
                    Else
                         ValidItemDetailsForLP = True
                         Exit Function
                End If
                                            
                    
                p_sDisplayErrMsg = k_MessageSPLRecordNotFound
                ValidItemDetailsForLP = False
                Exit Function
            Else
                p_sDisplayErrMsg = k_MessageSPLRecordNotFound
                ValidItemDetailsForLP = False
                Exit Function
            End If
        Else
            ValidItemDetailsForLP = False
            p_sDisplayErrMsg = k_MessageItemNotFound
        End If
        
    Else
        p_sDisplayErrMsg = k_MessageItemNotFound
        ValidItemDetailsForLP = False
    End If

End Function

Private Function ValidItemDetailsForSO(ByVal p_sItemID As String, ByRef p_sDisplayErrMsg As String) As Boolean

Dim sRet            As String
Dim sLocation       As String
'Dim sAPN            As String
Dim sItemKeycode    As String

    '----If not in stockroom then cannot perform Stockroom Out Function
    ValidItemDetailsForSO = False
    
    '----Important ************************************************************************
    '----p_sDisplayErrMsg is used to pass the location id in ....
    '----if Item not in the Location scanned a Error Msg will be return using p_sDisplayErrMsg to stored it...
    '----Note the p_sDisplayErrMsg is a byref param
    sLocation = Trim(p_sDisplayErrMsg)
    
    If SetApnList(p_sItemID) Then       '----Translate Scan to Keycode
        '----Take the first element to the collection
        'sAPN = gColApnList.Item(1)          '---- not required in this Pde function
        sItemKeycode = Trim(gtItemData.sItemID)
        
        p_sDisplayErrMsg = ""
        'If ItemInLocation(sItemKeycode, sLocation) Then
            ValidItemDetailsForSO = True
            Exit Function
        'Else
        '    p_sDisplayErrMsg = "Item not in Loc " & sLocation & " !"
        '    ValidItemDetailsForSO = False
        'End If
        
    Else
        p_sDisplayErrMsg = k_MessageItemNotFound
        ValidItemDetailsForSO = False
    End If

End Function
Private Function ValidItemDetailsForSI(ByVal p_sItemID As String, ByRef p_sDisplayErrMsg As String) As Boolean

Dim sRet            As String
Dim sItemKeycode    As String
Dim bResult         As Boolean


ValidItemDetailsForSI = False

    If SetApnList(p_sItemID) Then               '----Translate Scan to Keycode

        sItemKeycode = Trim(gtItemData.sItemID)
        If Trim(sItemKeycode) <> "" Then
            bResult = SetItemLocations(sItemKeycode)
            If SetItemDetails(sItemKeycode) Then
                    ValidItemDetailsForSI = True
                Else
                    p_sDisplayErrMsg = "Details Not Found"
                    ValidItemDetailsForSI = False
            End If
        End If
        Exit Function
    Else
        p_sDisplayErrMsg = k_MessageItemNotFound
        ValidItemDetailsForSI = False
    End If

End Function

Public Function ValidItemStatus(ByVal p_sPDEFunction As String) As Boolean
'
'----There should be more
'

Dim iItemStatus         As Integer

On Error GoTo ErrorHandler

    ValidItemStatus = False
    iItemStatus = gtItemData.tDetails.iItemStatus
    
    Select Case p_sPDEFunction
        Case k_Func_LowPresent
            
            Select Case iItemStatus
                Case k_Status_QuitCycle, k_Status_Allotment, k_Status_Clearance
                    ValidItemStatus = False
                Case -1 '-----Default
                    ValidItemStatus = False
                Case Else
                    '----All Other status are valid
                    ValidItemStatus = True
            End Select
                
        Case k_Func_NormalOrders
            
            Select Case iItemStatus
                Case k_Status_QuitCycle, k_Status_NewLine, k_Status_Allotment, k_Status_Clearance
                    ValidItemStatus = False
                Case -1 '-----Default
                    ValidItemStatus = False
                Case Else
                    '----All Other status are valid
                    ValidItemStatus = True
            End Select
            
        Case k_Func_StoreFeedback
            
            Select Case iItemStatus
                Case k_Status_QuitCycle, k_Status_Clearance
                    ' Modified as per CR17891
                    ' Flag set to FALSE will bypass all stock check functions and
                    ' causes the ValidateItem_SF routine to exit.
                    'ValidItemStatus = False
                    ValidItemStatus = True
                Case -1 '-----Default
                    ValidItemStatus = False
                Case Else
                    '----All Other status are valid
                    ValidItemStatus = True
            End Select
            
        Case k_Func_StockroomIn
            Select Case iItemStatus
                Case k_Status_QuitCycle, k_Status_Clearance
                    ValidItemStatus = False
                Case Else
                    '----All Other status are valid
                    ValidItemStatus = True
            End Select
            
    End Select
    
    Exit Function
ErrorHandler:

    ValidItemStatus = False
    LogError "modAddedValidation ", "ValidItemStatus", False
    err.Raise err.Number, err.Source, err.Description

End Function

Private Function ValidSourceOfSupply() As Boolean
'
'   Check if Store are allowed to Reorder this Item or Not
'   Conversion due to db schema ... is int not varchar/char
'
'   Auckland D.C        (A = 1)
'   Brisbane D.C        (B = 2)
'   Direct              (D = 4)
'   Not Re-Orderable    (N = 14)
'   Perth D.C           (P = 16)
'   Sydney D.C          (S = 19)
'   Melbourne D.C       (W = 23)
'

Dim sSourceOfSupply         As String

On Error GoTo ErrorHandler

    ValidSourceOfSupply = False
    sSourceOfSupply = gtItemData.tDetails.sSourceOfSupply
    
    Select Case sSourceOfSupply
        Case "A"  'k_Auckland   ', k_NotReOrderable
            ValidSourceOfSupply = False
        Case "B", "D", "P", "S", "W", k_Zero 'k_Brisbane, k_Direct, k_Perth, k_Sydney, k_Melbourne,
            ValidSourceOfSupply = True
        Case Else
            If sSourceOfSupply < "Z" Then    'k_Max_ValidNum
                ValidSourceOfSupply = True
            End If
    End Select
    
    Exit Function
ErrorHandler:

    ValidSourceOfSupply = False
    LogError "modAddedValidation", "ValidSourceOFSupply", False
    err.Raise err.Number, err.Source, err.Description

End Function

Private Function ValidReOrderFlag() As Boolean
'
'   Check if Store are allowed to Reorder this Item or Not
'   Store cannot order b/c its temporary out of Stock
'   Conversion  : 0 = N ---> No cannot order item
'               : 1 = Y ---> Yes can order item
'
Const k_StoreCannotReorder = 0

Dim iReOrderFlag         As Integer

On Error GoTo ErrorHandler

    ValidReOrderFlag = False
    iReOrderFlag = gtItemData.tDetails.iReorder
    
    If iReOrderFlag = k_StoreCannotReorder Then
        ValidReOrderFlag = False
    Else
        ValidReOrderFlag = True
    End If
    
    Exit Function
ErrorHandler:

    ValidReOrderFlag = False
    LogError "modAddedValidation", "ValidReOrderFlag", False
    err.Raise err.Number, err.Source, err.Description

End Function
Private Function ValidRecallFlag() As Boolean
'
'   Check if this Item is RECALL
'   Conversion  : 0 = N ---> No cannot order item
'               : 1 = Y ---> Yes can order item
'

Dim iRecallFlag         As Integer

On Error GoTo ErrorHandler

    ValidRecallFlag = False
    iRecallFlag = gtItemData.tDetails.iRecall
    
    If iRecallFlag Then
        ValidRecallFlag = False
    Else
        ValidRecallFlag = True
    End If
    
    Exit Function
ErrorHandler:

    ValidRecallFlag = False
    LogError "modAddedValidation", "ValidRecallFlag", False
    err.Raise err.Number, err.Source, err.Description

End Function

Public Function SetItemLocations(ByVal p_sItemID As String) As Boolean
'
'    Get the Location of Items and Save to structure ..for referenceing
'    The return value determine whether this Item is in Stockroom
'    Input Param is Stockroom Location therefore if Result is not empty
'    Then assume know that the Item is in the stockroom
'
Const k_FormatLoc = "0000"

Dim rs      As rdoResultset
Dim iRow    As Long
Dim iLoop1 As Integer

On Error GoTo ErrorHandler

    For iLoop1 = 1 To k_MaxLocation
        gtItemData.sLocationID(iLoop1) = Space(4)
    Next iLoop1
    
    iRow = 1            'Array of tLocation starts at 1
    SetItemLocations = False
    
    Set rs = getRsItemLocations(gIPMCon, p_sItemID)
    
    If Not rs Is Nothing Then
        With rs
            If .RowCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    If Not iRow > k_MaxLocation Then
                        If Not IsNull(!M_STK_RM_LOCN) Then
                               gtItemData.sLocationID(iRow) = Format(Trim(!M_STK_RM_LOCN), k_FormatLoc)
                           Else
                               gtItemData.sLocationID(iRow) = ""
                        End If
                    End If
                    
                    iRow = iRow + 1
                    .MoveNext
                
                Loop
                
                SetItemLocations = True
                .Close
            End If
        End With
    End If

    Set rs = Nothing

    Exit Function
ErrorHandler:
    SetItemLocations = False
    LogError "modAddedValidation", "SetItemLocations", False
    err.Raise err.Number, err.Source, err.Description

End Function
Public Function SetItemSPL(ByVal p_sItemID As String) As Boolean
'
'    Get the Location of Items and Save to structure ..for referenceing
'    The return value determine whether this Item is in Stockroom
'    Input Param is Stockroom Location therefore if Result is not empty
'    Then assume know that the Item is in the stockroom
'

Dim rs      As rdoResultset
Dim iRow    As Long
Dim iLoop1 As Integer

On Error GoTo ErrorHandler

    gtItemData.tSPL.dSPL_Low = -1
    gtItemData.tSPL.dSPL_High = -1
    gtItemData.tSPL.lMin_Lspl = -1
    gtItemData.tSPL.lMax_Lspl = -1
    
    SetItemSPL = False
    
    Set rs = getRsSPL(gIPMCon, p_sItemID)
    
    If Not rs Is Nothing Then
        With rs
            If .RowCount > 0 Then
                         
                If Not IsNull(!Q_UNIT_LOW_LEVEL) Then   'MIN_LOC_QTY
                        gtItemData.tSPL.dSPL_Low = Val(!Q_UNIT_LOW_LEVEL)
                End If
                
                If Not IsNull(!Q_UNIT_HIGH_LEVEL) Then 'max_loc_qty
                        gtItemData.tSPL.dSPL_High = Val(!Q_UNIT_HIGH_LEVEL)
                End If
                
                If Not IsNull(!Q_LSPL_MIN) Then 'min_lspl
                        gtItemData.tSPL.lMin_Lspl = Val(!Q_LSPL_MIN)
                End If
                
                If Not IsNull(!Q_LSPL_MAX) Then 'max_lspl
                        gtItemData.tSPL.lMax_Lspl = Val(!Q_LSPL_MAX)
                End If
                    
                
                SetItemSPL = True
                .Close
            End If
        End With
    End If

    Set rs = Nothing

    Exit Function
ErrorHandler:
    SetItemSPL = False
    LogError "modAddedValidation", "SetItemSPL", False
    err.Raise err.Number, err.Source, err.Description

End Function


Public Function ItemInLocation(ByVal p_sItemID As String, ByVal p_sLocation As String) As Boolean
'
'    Get the Location of Items and Save to structure ..for referenceing
'    The return value determine whether this Item is in Stockroom
'    Input Param is Stockroom Location therefore if Result is not empty
'    Then assume know that the Item is in the stockroom
'
Const k_FormatLoc = "0000"

Dim rs      As rdoResultset
Dim sLoc    As String

On Error GoTo ErrorHandler

    ItemInLocation = False
    p_sLocation = Format(Trim(p_sLocation), k_FormatLoc)
    Set rs = getRsItemLocations(gIPMCon, p_sItemID)
    
    If Not rs Is Nothing Then
        With rs
            If .RowCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    If Not IsNull(!M_STK_RM_LOCN) Then
                        sLoc = Format(Trim(!M_STK_RM_LOCN), k_FormatLoc)
                    End If
                    If p_sLocation = sLoc Then
                        '----Item found in this location
                        ItemInLocation = True
                        Exit Do
                        
                    End If
                
                .MoveNext
                
                Loop
                
                .Close
            
            End If
        End With
    End If

    Set rs = Nothing

    Exit Function
ErrorHandler:
    ItemInLocation = False
    LogError "modAddedValidation", "ItemInLocation", False
    err.Raise err.Number, err.Source, err.Description

End Function

Public Function SetApnList(ByVal p_sItemID As String) As Boolean
'
'   Get a collection of Apns Related to the current ItemID
'   Currently only interested in 4 of the list
'
Const k_MaxNumApnInList = 4

Dim rs      As rdoResultset
Dim iRow    As Long
Dim sSQL    As String

On Error GoTo ErrorHandler


    SetApnList = False
    
    Set rs = getRsApn(gIPMCon, p_sItemID)
    
    '----Clear out the collection
    Set gColApnList = New Collection
    
    iRow = 0
    
    'check rs is NULL, if yes, goto slt200 to find keycode
    'check input length which should be less or equal to 8 to be keycode
    If (rs.BOF And rs.EOF) And (Len(Trim(p_sItemID)) <= 8) Then
                                       
        sSQL = "SELECT M_KD FROM SLT200 WHERE  M_KD = " + p_sItemID
        Set rs = gIPMCon.OpenResultset(sSQL, rdOpenStatic)
        
        If Not rs Is Nothing Then
           With rs
        
                If .RowCount > 0 Then
                
                    .MoveFirst
                    
                    If Not IsNull(!M_KD) Then
                        gtItemData.sItemID = Trim(!M_KD)      '----Set the ItemID
                    End If
                    
                    SetApnList = True
                 
                End If
                
                .Close
                           
            End With
    
           
        
        End If
    
      Else
        If Not rs Is Nothing Then
            With rs
                
                If .RowCount > 0 Then
                
                    .MoveFirst
                    Do While Not .EOF
                        
                        gColApnList.Add Trim(!m_apn)
                                               
                        iRow = iRow + 1
                        If iRow = 1 Then
                            If Not IsNull(!M_KD) Then
                                gtItemData.sItemID = Trim(!M_KD)      '----Set the ItemID
                            End If
                        End If
                        
                        .MoveNext
                        
                    Loop
                    
                    SetApnList = True
                 
                End If
                
                .Close
                           
            End With
            
        
        End If
        
    End If
                
                
        

    If Not rs Is Nothing Then
       Set rs = Nothing
    End If
    
    Call CheckScanItemInCol(p_sItemID)

    Exit Function
ErrorHandler:
    SetApnList = False
    LogError "modAddedValidation", "SetApnList", False
    err.Raise err.Number, err.Source, err.Description

End Function

Public Function SetItemAttributes(ByVal p_sItemID As String) As Boolean
'
'   Get the Attributes of Items and Save to structure ..for referenceing
'
Dim rs                  As rdoResultset

Dim cl As rdoColumn

On Error GoTo ErrorHandler
    
    SetItemAttributes = False
        
    Set rs = getRsItemAttributes(gIPMCon, p_sItemID, k_SPParam_ALL)
    
    If Not rs Is Nothing Then
        With rs
            If .RowCount > 0 Then
                
                .MoveFirst
                Do While Not .EOF
                        If Not IsNull(!N_ABBR_COLR) Then
                                gtItemData.tAttributes.sColour = Trim(!N_ABBR_COLR)
                           Else
                                gtItemData.tAttributes.sColour = ""
                        End If
                        
                        If Not IsNull(!N_ABBR_SIZE) Then
                                gtItemData.tAttributes.sSize = Trim(!N_ABBR_SIZE)
                           Else
                                gtItemData.tAttributes.sSize = ""
                        End If
                        
                        If Not IsNull(!F_KD_BASIC_STK) Then
                                gtItemData.tAttributes.sBestSeller = Trim(!F_KD_BASIC_STK)
                           Else
                                gtItemData.tAttributes.sBestSeller = ""
                        End If
                    
                    .MoveNext
                    
                Loop
                .Close
            End If
            SetItemAttributes = True
            
        End With
    End If
    
    Set rs = Nothing

    Exit Function
ErrorHandler:
    SetItemAttributes = False
    LogError "modAddedValidation", "SetItemAttributes", False
    err.Raise err.Number, err.Source, err.Description

End Function

Public Function SetItemDetails(ByVal p_sItemID As String) As Boolean
'
'   Get the Details of Items and Save to structure ..for referenceing
'
Dim rs    As rdoResultset

On Error GoTo ErrorHandler

    SetItemDetails = False
    Set rs = getRsItemDetails(gIPMCon, p_sItemID)
    
    If Not rs Is Nothing Then
        With rs
            
            If .RowCount > 0 Then
                
                .MoveFirst
                Do While Not .EOF
                
                    If Not IsNull(p_sItemID) Then
                            gtItemData.sItemID = Trim(p_sItemID)
                       Else
                            gtItemData.sItemID = ""
                    End If
                    
                    If Not IsNull(!Q_UNIT_INNER) Then
                            gtItemData.tDetails.dPackSize = CDbl(!Q_UNIT_INNER)
                       Else
                            gtItemData.tDetails.dPackSize = 0#
                    End If
                    
                    
                    If Not IsNull(!C_MDEPT) Then
                            gtItemData.tDetails.iDeptCode = Val(!C_MDEPT)
                       Else
                            gtItemData.tDetails.iDeptCode = 0
                    End If
                    
                    
                    If Not IsNull(!C_KD_GRD) Then
                            gtItemData.tDetails.sGrading = Trim(!C_KD_GRD)
                       Else
                            gtItemData.tDetails.sGrading = ""
                    End If
                    
                    
                    If Not IsNull(!C_KD_STAT) Then
                            gtItemData.tDetails.iItemStatus = Val(!C_KD_STAT)
                        Else
                            gtItemData.tDetails.iItemStatus = 0
                    End If
                    
                    
                    If Not IsNull(!F_KD_RE_ORD) Then
                            If (UCase(Trim(!F_KD_RE_ORD))) = "Y" Then
                                gtItemData.tDetails.iReorder = 1
                            Else
                                gtItemData.tDetails.iReorder = 0
                            End If
                                
                       Else
                            gtItemData.tDetails.iReorder = 0
                    End If
                
                    If Not IsNull(!F_KD_RAINCHECK) Then
                            If (UCase(Trim(!F_KD_RAINCHECK))) = "Y" Then
                                gtItemData.tDetails.iRainCheck = 1
                            Else
                                gtItemData.tDetails.iRainCheck = 0
                            End If
                                
                       Else
                            gtItemData.tDetails.iRainCheck = 0
                    End If
                    
                    If Not IsNull(!F_KD_RECALL) Then
                            If (UCase(Trim(!F_KD_RECALL))) = "Y" Then
                                gtItemData.tDetails.iRecall = 1
                            Else
                                gtItemData.tDetails.iRecall = 0
                            End If
                                
                       Else
                            gtItemData.tDetails.iRecall = 0
                    End If
                    
                    If Not IsNull(!F_KD_ALP) Then
                            If (UCase(Trim(!F_KD_ALP))) = "Y" Then
                                gtItemData.tDetails.iALP = 1
                            Else
                                gtItemData.tDetails.iALP = 0
                            End If
                                
                       Else
                            gtItemData.tDetails.iALP = 0
                    End If
                    
                    If Not IsNull(!F_KD_CSI) Then
                            If (UCase(Trim(!F_KD_CSI))) = "Y" Then
                                gtItemData.tDetails.iCSI = 1
                            Else
                                gtItemData.tDetails.iCSI = 0
                            End If
                                
                       Else
                            gtItemData.tDetails.iCSI = 0
                    End If
                    
                    If Not IsNull(!F_KD_KVL) Then
                            If (UCase(Trim(!F_KD_KVL))) = "Y" Then
                                gtItemData.tDetails.iKVL = 1
                            Else
                                gtItemData.tDetails.iKVL = 0
                            End If
                                
                       Else
                            gtItemData.tDetails.iKVL = 0
                    End If
                    
                    
                    If Not IsNull(!C_KD_SRC) Then
                            gtItemData.tDetails.sSourceOfSupply = Trim(!C_KD_SRC)
                       Else
                            gtItemData.tDetails.sSourceOfSupply = ""
                    End If
                    
                    
                    If Not IsNull(!T_ABBR_DESC_KD) Then
                            gtItemData.tDetails.sItemDesc = Trim(!T_ABBR_DESC_KD)
                       Else
                            gtItemData.tDetails.sItemDesc = ""
                    End If
                    
                    
                    If Not IsNull(!T_BRAND_DESC_KD) Then
                            gtItemData.tDetails.sBrandDesc = Trim(!T_BRAND_DESC_KD)
                       Else
                            gtItemData.tDetails.sBrandDesc = ""
                    End If
                    
                    
                    If Not IsNull(!A_PR_POS) Then
                           gtItemData.tDetails.dSellPrice = !A_PR_POS
                       Else
                            gtItemData.tDetails.dSellPrice = 0#
                    End If
                    
                    
                    If Not IsNull(!A_CHRG_OUT_COST) Then
                            gtItemData.tDetails.dCostPrice = !A_CHRG_OUT_COST
                       Else
                            gtItemData.tDetails.dCostPrice = 0#
                    End If
                    
                    If Not IsNull(!Q_LBL_REQD) Then
                            gtItemData.tDetails.iQ_LBL_Reqd = !Q_LBL_REQD
                       Else
                            gtItemData.tDetails.iQ_LBL_Reqd = 0
                    End If
                    
                    If Not IsNull(!F_KD_IN_SIM) Then
                            If (UCase(Trim(!F_KD_IN_SIM))) = "Y" Then
                                gtItemData.tDetails.iInSIM = 1
                            Else
                                gtItemData.tDetails.iInSIM = 0
                            End If
                                
                       Else
                            gtItemData.tDetails.iInSIM = 0
                    End If
                    
                    If Not IsNull(!C_KD_RPLMT_MTHD) Then
                            gtItemData.tDetails.sReplenishmentMethod = Trim(!C_KD_RPLMT_MTHD)
                       Else
                            gtItemData.tDetails.sReplenishmentMethod = ""
                    End If
                    
                    If Not IsNull(!F_KD_SHELF_READY) Then
                            If (UCase(Trim(!F_KD_SHELF_READY))) = "Y" Then
                                gtItemData.tDetails.iShelfReady = 1
                            Else
                                gtItemData.tDetails.iShelfReady = 0
                            End If
                                
                       Else
                            gtItemData.tDetails.iShelfReady = 0
                    End If
                                        
                    .MoveNext
                Loop
                
                SetItemDetails = True
            
            End If
            
            .Close
                    
        End With
    End If
    
    Set rs = getRsItemProperties(gIPMCon, p_sItemID)
    
    If Not rs Is Nothing Then
        With rs
            
            If .RowCount > 0 Then
                
                .MoveFirst
                Do While Not .EOF
                    If Not IsNull(!M_STYLE) Then
                        gtItemData.tDetails.sMFG_Styl = !M_STYLE
                    Else
                        gtItemData.tDetails.sMFG_Styl = ""
                    End If
                 .MoveNext
                Loop
                .Close
            End If
        End With
    End If
    
    Set rs = Nothing

    Exit Function
ErrorHandler:
    SetItemDetails = False
    LogError "modAddedValidation", "SetItemDetails", False
    err.Raise err.Number, err.Source, err.Description

End Function

Public Function SetItemPrice(ByVal p_sItemID As String) As Boolean
'
'   Get the Price of Items and Save to structure ..for referenceing
'
Const k_Temp = "TEMPORARY"
Const k_Regular = "REGULAR"


Dim sPriceType              As String
Dim rs                      As rdoResultset
Dim bReadRegularPrice       As Boolean
Dim bReadTempPrice          As Boolean

On Error GoTo ErrorHandler


    bReadRegularPrice = False
    bReadTempPrice = False
    
    SetItemPrice = False
    sPriceType = ""
   
    Set rs = getRsItemPrice(gIPMCon, p_sItemID)
    
    gtItemData.tPrice.sRegularPrice = "     "
    gtItemData.tPrice.sTempPrice = "     "
        
    If Not rs Is Nothing Then
        With rs
            
            If .RowCount > 0 Then
                
                .MoveFirst
                Do While Not .EOF
                
                    sPriceType = UCase(Trim(!C_PR_TYPE))
                    '
                    If (sPriceType <= "E") Then
                        bReadTempPrice = True
                        If Not IsNull(!A_PR) Then
                                gtItemData.tPrice.sTempPrice = Format(Trim(!A_PR), "####.00")
                            Else
                                gtItemData.tPrice.sTempPrice = "0.00"
                        End If
                        
                     Else
                    
                        bReadRegularPrice = True
                            
                        If Not IsNull(!A_PR) Then
                                gtItemData.tPrice.sRegularPrice = Format(Trim(!A_PR), "####.00")
                           Else
                                gtItemData.tPrice.sRegularPrice = "0.00"
                        End If
                        
                        
                    End If
                    
                    .MoveNext
                Loop
                
                .Close
                SetItemPrice = True
                
            End If
            
        End With
    End If
    
    Set rs = Nothing

    Exit Function
ErrorHandler:
    SetItemPrice = False
    LogError "modAddedValidation", "SetItemPrice", False
    err.Raise err.Number, err.Source, err.Description

End Function

Public Function ValidAddedCondition(ByVal p_sItemID As String, ByVal p_sPDEFunction As String, ByRef p_ErrorMessage As String, ByVal p_bValidate As String, Optional p_StorefeedBackSubType As Variant) As Boolean
'
'   This Function will serve as the Main Function to be called for
'   Items to be have additionally validation therefore
'
'
Const k_YesCheck = "Y"

On Error GoTo ErrorHandler

    '----Registry flag has indicated that we don't perform additional validation
    If UCase(Trim(p_bValidate)) <> k_YesCheck Then
        ValidAddedCondition = True
        Exit Function
    End If

    ValidAddedCondition = False
    p_sPDEFunction = UCase(Trim(p_sPDEFunction))
        
    Select Case p_sPDEFunction
        Case k_Func_StockroomIn
            '----Check Item in Stockroom
            If ValidItemDetailsForSI(p_sItemID, p_ErrorMessage) Then
                ValidAddedCondition = True
            End If
                        
        Case k_Func_Requisition
            If ValidItemDetailsForRN(p_sItemID, p_sPDEFunction, p_ErrorMessage) Then
                ValidAddedCondition = True
            End If

        Case k_Func_RLOAssessment
            If ValidItemDetailsForRA(p_sItemID, p_sPDEFunction, p_ErrorMessage) Then
                ValidAddedCondition = True
            End If
            
        Case k_Func_GroupCapture
            If ValidItemDetailsForGC(p_sItemID, p_sPDEFunction, p_ErrorMessage) Then
                ValidAddedCondition = True
            End If
        
        Case k_Func_StockroomOUT
            '----Check Item in Stockroom
            If ValidItemDetailsForSO(p_sItemID, p_ErrorMessage) Then
                ValidAddedCondition = True
            End If
        
        Case k_Func_LowPresent
            p_ErrorMessage = ""
            If ValidItemDetailsForLP(p_sItemID, p_sPDEFunction, p_ErrorMessage) Then
                ValidAddedCondition = True
            End If
        
        Case k_Func_NormalOrders
            p_ErrorMessage = ""
            If ValidItemDetailsForNO(p_sItemID, p_sPDEFunction, p_ErrorMessage) Then
                ValidAddedCondition = True
            End If
        
        Case k_Func_StoreFeedback
            p_ErrorMessage = ""
            If ValidItemDetailsForSF(p_sItemID, p_sPDEFunction, p_StorefeedBackSubType, p_ErrorMessage) Then
                ValidAddedCondition = True
            End If

        Case k_Func_MerchMultiLocatedItem
            p_ErrorMessage = ""
            If ValidItemDetailsForML(p_sItemID, p_sPDEFunction, p_ErrorMessage) Then
                ValidAddedCondition = True
            End If
        
        Case k_Func_EmptyPackets
            p_ErrorMessage = ""
            If ValidItemDetailsForEP(p_sItemID, p_sPDEFunction, p_ErrorMessage) Then
                ValidAddedCondition = True
            End If
          
                    
    End Select

    Exit Function
ErrorHandler:
    ValidAddedCondition = False
    LogError "modAddedValidation", "ValidAddedCondition", False
    err.Raise err.Number, err.Source, err.Description

End Function

Private Function CheckScanItemInCol(ByVal p_sItemID As String) As Boolean
    
Dim iCount          As Integer
Dim iPos            As Integer
Dim tmpCollection   As New Collection

iPos = 0
CheckScanItemInCol = False

If gColApnList.Count > 0 Then

    For iCount = 1 To gColApnList.Count
        If gColApnList.Item(iCount) = Trim(p_sItemID) Then
            iPos = iCount
            Exit For
        End If
    Next
           
    If iPos >= 1 Then
        
        For iCount = 1 To gColApnList.Count
    
            With tmpCollection
                If iCount = 1 Then
                    .Add Trim(p_sItemID)
                End If
                If iCount <> iPos Then
                    .Add gColApnList.Item(iCount)
                End If
            End With
            
        Next
        
        Set gColApnList = tmpCollection
        
        CheckScanItemInCol = True
        Set tmpCollection = Nothing
        
    End If
End If
    
End Function

Public Function ValidateStyle(ByVal p_sItemID As String) As Boolean
'
'   Check if the number is that of a style in the SLT929 table
'
Dim rs    As rdoResultset

On Error GoTo ErrorHandler

    ValidateStyle = False
    
    Set rs = getRsStyleProperties(gIPMCon, p_sItemID)
    
    If Not rs Is Nothing Then
        With rs
            
            If .RowCount > 0 Then
                
                .MoveFirst
                Do While Not .EOF
                    ValidateStyle = True
                 .MoveNext
                Loop
                .Close
            End If
        End With
    End If

    Set rs = Nothing

    Exit Function
ErrorHandler:
    ValidateStyle = False
    LogError "modAddedValidation", "ValidateStyle", False
    err.Raise err.Number, err.Source, err.Description

End Function
