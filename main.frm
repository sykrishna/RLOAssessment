VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "frmmain"
   ClientHeight    =   2745
   ClientLeft      =   2655
   ClientTop       =   2670
   ClientWidth     =   3600
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MouseIcon       =   "main.frx":030A
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2745
   ScaleWidth      =   3600
   Begin MSWinsockLib.Winsock TCP1 
      Index           =   0
      Left            =   3000
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'variables
Dim mbNoArg                As Boolean
Dim mvtGetData             As Variant
Dim msSeqType              As String
Dim msInputStr             As String
Dim miSeqNo                As Integer
Dim miEntryScreen          As Integer
Dim mbContinuousScan       As Boolean

Private msRawData          As String

'Consts
Const k_Max_ItemEntryRow = 2
Const k_Max_ItemEntryCol = 6
Const k_Max_MessageRow = 7
Const k_Max_MessageCol = 1

Private Sub ProcessDisplayItemPrompt(ByRef p_miState As Integer, ByRef p_StayInLoop As Integer)

On Error GoTo ErrorHandler
           
   clsMax.PrintStr 1, 1, "   RLO Assessment   "
   clsMax.PrintStr 2, 1, Space(20)
   clsMax.PrintStr 3, 1, Space(20)
   clsMax.PrintStr 4, 1, Space(20)
   clsMax.PrintStr 5, 1, Space(20)
   clsMax.PrintStr 6, 1, Space(20)
   clsMax.PrintStr 7, 1, Space(20)
   
   clsMax.PrintStr k_Max_ItemEntryRow, 1, Left("Item:" + Space(20), 20)
   
   clsMax.MsgLine "Scan Item"
   clsMax.AcceptStr k_Max_ItemEntryRow, k_Max_ItemEntryCol
   'clsMax.MaxScan_On
   clsMax.MaxNumeric_On
      
   p_miState = k_Max_GetUIItemState
   p_StayInLoop = False

   Exit Sub
ErrorHandler:
   
    HandleErrorFatal "frmMain", "ProcessDisplayItemPrompt", err.Source, err.Number, err.Description

End Sub

Private Sub ProcessItem(ByRef p_miState As Integer, ByRef p_StayInLoop As Integer, ByVal p_inputStr As Variant)
'process sell id state

Dim sItem             As String
Dim sItemDesc         As String
Dim sKeycode          As String
Dim sErrMessage       As String
Dim iRetValue         As Integer
Dim bRet              As Boolean
Dim iPos              As Integer
Dim iReturnReasonCode As Integer
Dim sRLOReturnType    As String
Dim iADJCode          As Integer
Dim lSoh              As Long
Dim dAWC              As Double
Dim dUnitCost         As Double

On Error GoTo ErrorHandler

p_StayInLoop = False
    
    Select Case msSeqType
       
      Case KEYREFRESH                                              'refreshes the screen
           clsMax.mDebug_Log "Refresh"
           clsMax.RefreshScr
           Exit Sub
          
      Case KEYEND, KEYQUIT                                          'quit returns back to aisle state
           clsMax.MsgLine "Application Close "
           p_miState = k_Max_EndState
           p_StayInLoop = True
           Exit Sub
      
      Case KEYRETURN
           
	       iPos = InStr(1, msRawData, ",")
		    
		   If iPos = 23 Then
		       'Pathfinder barcode
		       If Mid(msRawData, 1, 2) = "05" Then
		     	   sItem = Mid(msRawData, 4, 13)
		       Else
		           sItem = Left(msRawData, iPos - 1)
		       End If
           ElseIf iPos > 0 Then
               sItem = Left(msRawData, iPos - 1)
           Else
               sItem = Trim(p_inputStr)
           End If

           clsRLOAssessment.psPDTMsg = ""
           
           WriteToErrorLog "sITEM " + sItem

           If Len(sItem) = 0 Then
               clsMax.MaxWarnBell
               clsRLOAssessment.psPDTMsg = ""
               p_miState = k_Max_DisplayItemPromptState
               p_StayInLoop = True
               Exit Sub
           End If
           
           iRetValue = ValidateItem_RLOAssessment(sItem, sErrMessage)

           If iRetValue <> 0 Then
               clsMax.MaxWarnBell
               clsRLOAssessment.psPDTMsg = ""
               p_miState = k_Max_DisplayItemPromptState
               p_StayInLoop = True
               Exit Sub
           End If
                       
           sKeycode = gtItemData.sItemID
           sItemDesc = gtItemData.tDetails.sItemDesc
           clsRLOAssessment.sAPN = gtItemData.tDetails.sRLO_APN
           clsRLOAssessment.dSellPrice = gtItemData.tDetails.dSellPrice
           
           If gtItemData.tDetails.iDeptCode = 0 Then
               clsRLOAssessment.iDeptCode = k_DefaultDept
           Else
               clsRLOAssessment.iDeptCode = gtItemData.tDetails.iDeptCode
           End If
           
           If IsNumeric(gtItemData.tDetails.sRLO_ReasonCode) Then
               iReturnReasonCode = Val(gtItemData.tDetails.sRLO_ReasonCode)
           Else
               iReturnReasonCode = 0
           End If
                       
           WriteToErrorLog "Keycode " + sKeycode + " selected for RLO Assessment"

           iRetValue = ValidateRLOItemStep1(sKeycode, iReturnReasonCode, gtItemData.tDetails.sRLO_ReasonSubCode, sRLOReturnType, iADJCode, sErrMessage, dUnitCost)

           WriteToErrorLog "iRetValue = " + CStr(iRetValue) + ". Keycode = " + sKeycode + ". ReturnReasonCode = " + CStr(iReturnReasonCode) + ". Unit Cost = " + Format(dUnitCost, "0.00") 
           WriteToErrorLog "RLO_ReasonSubCode = " + gtItemData.tDetails.sRLO_ReasonSubCode + ". RLOReturnType = " + sRLOReturnType + ". ADJCode = " + CStr(iADJCode) + ". ErrMessage = " + sErrMessage + "."
       
           clsRLOAssessment.piADJCode = iADJCode
           clsRLOAssessment.psRLOReturnType = sRLOReturnType
           clsRLOAssessment.psItem_ReturnType = sRLOReturnType
           clsRLOAssessment.pdUnitCost = dUnitCost

           If iRetValue = 0 Then
               If sRLOReturnType = "Recall" Then
                   p_miState = k_Max_DisplayRecallItemPromptState
                   p_StayInLoop = True
                   Exit Sub
               Else
                   If sRLOReturnType = "Claimable" Then
                       clsRLOAssessment.psRLOReturnType = k_Claimable
                       clsRLOAssessment.psItem_ReturnType = k_Claimable
                       clsRLOAssessment.piADJCode = k_ReturnReasonCode_Claimable
                       clsRLOAssessment.psPDTMsg = k_Claimable + " Pallet"
                       p_miState = k_Max_DisplaySCMPromptState
                       p_StayInLoop = True
                       Exit Sub
                   Else
	                   If sRLOReturnType = "HoldForTesting" Then
	                       clsRLOAssessment.psRLOReturnType = k_Claimable
	                       clsRLOAssessment.psItem_ReturnType = k_Salvage
	                       clsRLOAssessment.piADJCode = k_ReturnReasonCode_Salvage
	                       clsRLOAssessment.psPDTMsg = k_Claimable + " Pallet"
	                       p_miState = k_Max_DisplaySCMPromptState
	                       p_StayInLoop = True
	                       Exit Sub
	                   Else
	                       If sRLOReturnType = "WriteOff" Then
	                           WriteToErrorLog clsRLOAssessment.psCreateUser + " Keycode = " + sKeycode + ". Item to be thrown out."
	                           WriteToErrorLog clsRLOAssessment.psCreateUser + " Keycode = " + sKeycode + ". Reason code = WO. " + _
	                                           "Quantity = 1. SOH = 0. Store count = 0."
	
					           'get SOH and AWC values from IPS
					           bRet = GetValuesFromIPS(sKeycode, lSoh, dAWC)
					
					           clsRLOAssessment.pdAWC = dAWC
	                           
	                           If clsRLOAssessment.CreateEmptyPacketsDTL(sKeycode, 1) Then
			                   End If
			                   
	                           p_miState = k_Max_DisplayWriteOffItemPromptState
	                           p_StayInLoop = True
	                           Exit Sub
	                       Else
	                           If sRLOReturnType = "Salvage" Then
	                               p_miState = k_Max_ProcessCheck_ClothingFootwearHeaterXmas
	                               p_StayInLoop = True
	                               Exit Sub
	                           Else
	                               If sRLOReturnType = "FoodItem" Or sRLOReturnType = "PersonalItem" Then
	                               	   If clsRLOAssessment.pdUnitCost < 7.0 Then
				                           If sRLOReturnType = "FoodItem" Then
				                               WriteToErrorLog clsRLOAssessment.psCreateUser + " Keycode = " + sKeycode + ". Item to be thrown out. Item is food."
				                           Else
				                               WriteToErrorLog clsRLOAssessment.psCreateUser + " Keycode = " + sKeycode + ". Item to be thrown out. Item is cosmetics or personal care item."
				                           End If
				                               
				                           WriteToErrorLog clsRLOAssessment.psCreateUser + " Keycode = " + sKeycode + ". Reason code = WO. " + _
				                                           "Quantity = 1. SOH = 0. Store count = 0."
				
								           'get SOH and AWC values from IPS
								           bRet = GetValuesFromIPS(sKeycode, lSoh, dAWC)
								
								           clsRLOAssessment.pdAWC = dAWC
				                           
				                           If clsRLOAssessment.CreateEmptyPacketsDTL(sKeycode, 1) Then
						                   End If
						                   
				                           p_miState = k_Max_DisplayWriteOffItemPromptState
				                           p_StayInLoop = True
				                           Exit Sub
				                       Else
			                               p_miState = k_Max_ProcessDisplayContainHazards
			                               p_StayInLoop = True
			                               Exit Sub
			                           End If
	                               Else
		                               p_miState = k_Max_ProcessDisplayContainHazards
		                               p_StayInLoop = True
		                               Exit Sub
		                           End If
	                           End If
	                       End If
	                   End If
                   End If
               End If
           Else
               clsMax.MaxErrBell
               clsMax.MsgLine "Invalid Item"
               clsMax.MaxNumeric_On
               p_miState = k_Max_DisplayItemPromptState
               p_StayInLoop = True
               Exit Sub
           End If

           Exit Sub
                  
      Case Else
           Exit Sub
    
    End Select
    
    Exit Sub
ErrorHandler:

    HandleErrorFatal "frmMain", "ProcessItem", err.Source, err.Number, err.Description

End Sub

Private Sub ProcessDisplayRecallItemPrompt(ByRef p_miState As Integer, ByRef p_StayInLoop As Integer)

On Error GoTo ErrorHandler
           
   clsMax.PrintStr 1, 1, "   RLO Assessment   "
   clsMax.PrintStr 2, 1, Space(20)
   clsMax.PrintStr 3, 1, Space(20)
   clsMax.PrintStr 4, 1, Space(20)
   clsMax.PrintStr 5, 1, Space(20)
   clsMax.PrintStr 6, 1, Space(20)
   clsMax.PrintStr 7, 1, Space(20)
   clsMax.PrintStr 8, 1, Space(20)
   
   clsMax.PrintStr k_Max_ItemEntryRow, 1, Left("Item:" + gtItemData.sItemID + Space(20), 20)
   clsMax.PrintStr 3, 1, "APN:" + gtItemData.tDetails.sRLO_APN
   clsMax.PrintStr 4, 1, gtItemData.tDetails.sItemDesc
   clsMax.PrintStr 5, 1, "WITHDRAWN ITEM      "
   clsMax.PrintStr 6, 1, "See manager for     "
   clsMax.PrintStr 7, 1, "instructions        "

   clsMax.MsgLine "Press Enter         "
   clsMax.MaxWarnBell

   clsMax.AcceptStr 8, 15
   'clsMax.MaxScan_On
   clsMax.MaxNumeric_On
      
   p_miState = k_Max_GetUIRecallItemState
   p_StayInLoop = False

   Exit Sub
ErrorHandler:
   
   HandleErrorFatal "frmMain", "ProcessDisplayRecallItemPrompt", err.Source, err.Number, err.Description

End Sub

Private Sub ProcessRecallItem(ByRef p_miState As Integer, ByRef p_StayInLoop As Integer, ByVal p_inputStr As Variant)
'process sell id state

Dim sItem            As String
Dim sItemDesc        As String
Dim sKeycode         As String
Dim sErrMessage      As String
Dim iRetValue        As Integer
Dim bRet             As Boolean
Dim iPos             As Integer

On Error GoTo ErrorHandler

p_StayInLoop = False
    
    Select Case msSeqType
       
      Case KEYREFRESH                                              'refreshes the screen
           clsMax.mDebug_Log "Refresh"
           clsMax.RefreshScr
           Exit Sub
          
      Case KEYEND, KEYQUIT                                          'quit returns back to aisle state

           p_miState = k_Max_DisplayItemPromptState
           p_StayInLoop = True
           Exit Sub
      
      Case KEYRETURN
           iPos = InStr(1, msRawData, ",")
            
           If iPos > 0 Then
               sItem = Left(msRawData, iPos - 1)
           Else
               sItem = Trim(p_inputStr)
           End If

           If Len(sItem) = 0 Then
               p_miState = k_Max_DisplayItemPromptState
               p_StayInLoop = True
           Else
               clsMax.MaxWarnBell
               p_miState = k_Max_DisplayRecallItemPromptState
               p_StayInLoop = True
           End If
           Exit Sub
                  
      Case Else
           Exit Sub
    
    End Select
    
    Exit Sub
ErrorHandler:

    HandleErrorFatal "frmMain", "ProcessRecallItem", err.Source, err.Number, err.Description

End Sub

Private Sub DisplaySCMError(ByRef p_miState As Integer, ByRef p_StayInLoop As Integer, ByVal p_PDTMsg As String)

On Error GoTo ErrorHandler
           
   clsMax.PrintStr 1, 1, "   RLO Assessment   "
   clsMax.PrintStr 2, 1, Space(20)
   clsMax.PrintStr 3, 1, Space(20)
   clsMax.PrintStr 4, 1, Space(20)
   clsMax.PrintStr 5, 1, Space(20)
   clsMax.PrintStr 6, 1, Space(20)
   clsMax.PrintStr 7, 1, Space(20)
   clsMax.PrintStr 8, 1, Space(20)
   
   clsMax.PrintStr 2, 1, Left("Item:" + gtItemData.sItemID + Space(20), 20)
   clsMax.PrintStr 3, 1, Left("Current " + clsRLOAssessment.psRLOReturnType + Space(20), 20)
   If Len(p_PDTMsg) >= 40 Then
      clsMax.PrintStr 4, 1, Left("Pallet SCM # in use " + Space(20), 20)
      clsMax.PrintStr 5, 1, Mid(p_PDTMsg, 21, 20)
      If Len(p_PDTMsg) >= 60 Then
         clsMax.PrintStr 6, 1, Mid(p_PDTMsg, 41, 20)
      End If
   End If
   
   clsMax.PrintStr 7, 1, Left(p_PDTMsg + Space(20), 20)
   clsMax.MsgLine Left("Press Enter" + Space(20), 20)

   clsMax.AcceptStr 8, 15
   'clsMax.MaxScan_On
   clsMax.MaxNumeric_On
      
   p_miState = k_Max_ProcessSCMErrorMsgState
   p_StayInLoop = False

   Exit Sub
ErrorHandler:
   
   HandleErrorFatal "frmMain", "DisplaySCMError", err.Source, err.Number, err.Description

End Sub

Private Sub ProcessDisplayWriteOffItemPrompt(ByRef p_miState As Integer, ByRef p_StayInLoop As Integer)

On Error GoTo ErrorHandler
           
   clsMax.PrintStr 1, 1, "   RLO Assessment   "
   clsMax.PrintStr 2, 1, Space(20)
   clsMax.PrintStr 3, 1, Space(20)
   clsMax.PrintStr 4, 1, Space(20)
   clsMax.PrintStr 5, 1, Space(20)
   clsMax.PrintStr 6, 1, Space(20)
   clsMax.PrintStr 7, 1, Space(20)
   clsMax.PrintStr 8, 1, Space(20)
   
   clsMax.PrintStr k_Max_ItemEntryRow, 1, Left("Item:" + gtItemData.sItemID + Space(20), 20)
   clsMax.PrintStr 3, 1, "APN:" + gtItemData.tDetails.sRLO_APN
   clsMax.PrintStr 4, 1, gtItemData.tDetails.sItemDesc
   clsMax.PrintStr 5, 1, "Dispose of item.    "
   clsMax.PrintStr 6, 1, "Item has been       "
   clsMax.PrintStr 7, 1, "written off.        "

   clsMax.MsgLine "Press Enter         "

   clsMax.AcceptStr 8, 15
   'clsMax.MaxScan_On
   clsMax.MaxNumeric_On
      
   p_miState = k_Max_GetUIWriteOffItemState
   p_StayInLoop = False

   Exit Sub
ErrorHandler:
   
   HandleErrorFatal "frmMain", "ProcessDisplayWriteOffItemPrompt", err.Source, err.Number, err.Description

End Sub

Private Sub ProcessDisplayClearanceItemPrompt(ByRef p_miState As Integer, ByRef p_StayInLoop As Integer)
Dim sPrice        As String
Dim sRegularPrice As String

On Error GoTo ErrorHandler
           
   clsMax.PrintStr 1, 1, "   RLO Assessment   "
   clsMax.PrintStr 2, 1, Space(20)
   clsMax.PrintStr 3, 1, Space(20)
   clsMax.PrintStr 4, 1, Space(20)
   clsMax.PrintStr 5, 1, Space(20)
   clsMax.PrintStr 6, 1, Space(20)
   clsMax.PrintStr 7, 1, Space(20)
   clsMax.PrintStr 8, 1, Space(20)
   
   clsMax.PrintStr k_Max_ItemEntryRow, 1, Left("Item:" + gtItemData.sItemID + Space(20), 20)
   clsMax.PrintStr 3, 1, "APN:" + gtItemData.tDetails.sRLO_APN
   clsMax.PrintStr 4, 1, gtItemData.tDetails.sItemDesc
   clsMax.PrintStr 5, 1, "Price item and place"
   clsMax.PrintStr 6, 1, "on Clearance Trolley"
   
   If gtItemData.tPrice.sRegularPrice = "     " Then
   	   sRegularPrice = "0.0"
   Else
       sRegularPrice = gtItemData.tPrice.sRegularPrice
   End If
   
   sPrice = Format(0.5 * CDbl(sRegularPrice), "0.00")
   
   If Len(sPrice) > 5 Then
   	   sPrice = Mid(sPrice, 1, InStr(sPrice, ".") - 1)
   End If
   
   clsMax.PrintStr 7, 1, "Suggest price $" + sPrice

   clsMax.MsgLine "Press Enter         "

   clsMax.AcceptStr 8, 15
   'clsMax.MaxScan_On
   clsMax.MaxNumeric_On
      
   p_miState = k_Max_GetUIClearanceItemState
   p_StayInLoop = False

   Exit Sub
ErrorHandler:
   
   HandleErrorFatal "frmMain", "ProcessDisplayClearanceItemPrompt", err.Source, err.Number, err.Description

End Sub

Private Sub ProcessWriteOffItem(ByRef p_miState As Integer, ByRef p_StayInLoop As Integer, ByVal p_inputStr As Variant)
'process sell id state

Dim sItem            As String
Dim sItemDesc        As String
Dim sKeycode         As String
Dim sErrMessage      As String
Dim iRetValue        As Integer
Dim bRet             As Boolean
Dim iPos             As Integer

On Error GoTo ErrorHandler

p_StayInLoop = False
    
    Select Case msSeqType
       
      Case KEYREFRESH                                              'refreshes the screen
           clsMax.mDebug_Log "Refresh"
           clsMax.RefreshScr
           Exit Sub
          
      Case KEYEND, KEYQUIT                                          'quit returns back to aisle state

           p_miState = k_Max_DisplayItemPromptState
           p_StayInLoop = True
           Exit Sub
      
      Case KEYRETURN
           iPos = InStr(1, msRawData, ",")
            
           If iPos > 0 Then
               sItem = Left(msRawData, iPos - 1)
           Else
               sItem = Trim(p_inputStr)
           End If

           If Len(sItem) = 0 Then
               p_miState = k_Max_DisplayItemPromptState
               p_StayInLoop = True
           Else
               clsMax.MaxWarnBell
               p_miState = k_Max_DisplayWriteOffItemPromptState
               p_StayInLoop = True
           End If
           Exit Sub
                  
      Case Else
           Exit Sub
    
    End Select
    
    Exit Sub
ErrorHandler:

    HandleErrorFatal "frmMain", "ProcessWriteOffItem", err.Source, err.Number, err.Description

End Sub

Private Sub ProcessDisplaySaleableOnClearanceTrolley(ByRef p_miState As Integer, ByRef p_StayInLoop As Integer)
Dim sRtn     As String
Dim bRet     As Boolean
Dim sKeycode As String
Dim lSoh     As Long
Dim dAWC     As Double

On Error GoTo ErrorHandler
    
    sRtn = clsMax.DlgBox(msInputStr, "Is this item        ", "saleable on the     ", "Clearance Trolley?  ", MSGTYP_YESNO)
   
    Select Case sRtn

       Case "N"
        	sKeycode = gtItemData.sItemID
        	
        	WriteToErrorLog clsRLOAssessment.psCreateUser + " Keycode = " + sKeycode + ". Saleable on the Clearance Trolley = N."
        	
        	If clsRLOAssessment.psRLOReturnType = "FoodItem" Or clsRLOAssessment.psRLOReturnType = "PersonalItem" Then 
                clsMax.MsgLine "                    "
                
                If clsRLOAssessment.psRLOReturnType = "FoodItem" Then
                    WriteToErrorLog clsRLOAssessment.psCreateUser + " Keycode = " + sKeycode + ". Item to be thrown out. Item is food."
                Else
                    WriteToErrorLog clsRLOAssessment.psCreateUser + " Keycode = " + sKeycode + ". Item to be thrown out. Item is cosmetics or personal care item." 
                End If
                    
                WriteToErrorLog clsRLOAssessment.psCreateUser + " Keycode = " + sKeycode + ". Reason code = WO. " + _
                                "Quantity = 1. SOH = 0. Store count = 0."

	            'get SOH and AWC values from IPS
	            bRet = GetValuesFromIPS(sKeycode, lSoh, dAWC)
	
	            clsRLOAssessment.pdAWC = dAWC
               
                If clsRLOAssessment.CreateEmptyPacketsDTL(sKeycode, 1) Then
                End If
               
                p_miState = k_Max_DisplayWriteOffItemPromptState
                p_StayInLoop = True
                Exit Sub
            Else
                p_miState = k_Max_ProcessCheck_ClothingFootwearHeaterXmas
                p_StayInLoop = True
            End If
            
       Case "Y"
            WriteToErrorLog clsRLOAssessment.psCreateUser + " Keycode = " + gtItemData.sItemID + ". Saleable on the Clearance Trolley = Y."
            WriteToErrorLog clsRLOAssessment.psCreateUser + " Keycode = " + gtItemData.sItemID + ". Item to be placed on the Clearance Trolley."'            
            p_miState = k_Max_DisplayClearanceItemPromptState
            p_StayInLoop = True
       
       Case Else
            p_StayInLoop = False
            Exit Sub

    End Select

Exit Sub
ErrorHandler:

    HandleErrorFatal "frmMain", "ProcessDisplaySaleableOnClearanceTrolley", err.Source, err.Number, err.Description

End Sub

Private Sub ProcessDisplayContainHazards(ByRef p_miState As Integer, ByRef p_StayInLoop As Integer)

On Error GoTo ErrorProcessDisplayContainHazards

    clsMax.PrintStr 1, 1, "   RLO Assessment   "
    clsMax.PrintStr 2, 1, Space(20)
    clsMax.PrintStr 3, 1, Space(20)
    clsMax.PrintStr 4, 1, Space(20)
    clsMax.PrintStr 5, 1, Space(20)
    clsMax.PrintStr 6, 1, Space(20)
    clsMax.PrintStr 7, 1, Space(20)
    clsMax.PrintStr 8, 1, Space(20)

    clsMax.PrintStr 1, 1, "   RLO Assessment   "
    clsMax.PrintStr 2, 1, "--------------------"
    clsMax.PrintStr 3, 1, "Does this item still"
    clsMax.PrintStr 4, 1, "contain Chemicals,  "
    clsMax.PrintStr 5, 1, "Sharp Edges or      "
    clsMax.PrintStr 6, 1, "Broken Glass?       "
    clsMax.PrintStr 7, 1, "Press Y or N        "
    clsMax.MsgLine "--------------------"
 

    clsMax.AcceptStr 7, 17
    'clsMax.MaxScan_On
    clsMax.MaxNumeric_Off
    
    p_miState = k_Max_ProcessContainHazards
    p_StayInLoop = False

    Exit Sub

ErrorProcessDisplayContainHazards:
    
    HandleErrorFatal "frmMain", "ProcessDisplayContainHazards", err.Source, err.Number, err.Description
End Sub

Private Sub ProcessContainHazards(ByRef p_miState As Integer, ByRef p_StayInLoop As Integer, ByVal p_inputStr As Variant)
Dim bRet     As Boolean
Dim sKeycode As String
Dim lSoh     As Long
Dim dAWC     As Double

On Error GoTo ErrorHandler

    p_StayInLoop = False
    
    Select Case msSeqType
       
        Case KEYRETURN
            If Len(p_inputStr) = 1 Then

                Select Case p_inputStr
                    Case Is = "Y"
                        clsMax.MsgLine "                    "
                        
                        sKeycode = gtItemData.sItemID
                        
                        WriteToErrorLog clsRLOAssessment.psCreateUser + " Keycode = " + sKeycode + ". Chemicals, Sharp Edges or Broken Glass = Y."
                        
                        If clsRLOAssessment.psRLOReturnType = "FoodItem" Then
                            WriteToErrorLog clsRLOAssessment.psCreateUser + " Keycode = " + sKeycode + ". Item to be thrown out. Item is food."
                        ElseIf clsRLOAssessment.psRLOReturnType = "PersonalItem" Then
                        	WriteToErrorLog clsRLOAssessment.psCreateUser + " Keycode = " + sKeycode + ". Item to be thrown out. Item is cosmetics or personal care item."
                        Else
                            WriteToErrorLog clsRLOAssessment.psCreateUser + " Keycode = " + sKeycode + ". Item to be thrown out."
                        End If
                        
                        WriteToErrorLog clsRLOAssessment.psCreateUser + " Keycode = " + sKeycode + ". Reason code = WO. " + _
                                        "Quantity = 1. SOH = 0. Store count = 0."

			            'get SOH and AWC values from IPS
			            bRet = GetValuesFromIPS(sKeycode, lSoh, dAWC)
			
			            clsRLOAssessment.pdAWC = dAWC
                       
                        If clsRLOAssessment.CreateEmptyPacketsDTL(sKeycode, 1) Then
	                    End If
	                   
                        p_miState = k_Max_DisplayWriteOffItemPromptState
                        p_StayInLoop = True
                        Exit Sub
                   
                    Case Is = "N"
                        clsMax.MsgLine "                    "
                        
                        sKeycode = gtItemData.sItemID
                        
                        WriteToErrorLog clsRLOAssessment.psCreateUser + " Keycode = " + sKeycode + ". Chemicals, Sharp Edges or Broken Glass = N."
                        
                        If clsRLOAssessment.pdUnitCost < 7.0 Then
                        	If clsRLOAssessment.psRLOReturnType = "FoodItem" Or clsRLOAssessment.psRLOReturnType = "PersonalItem" Then 
		                        clsMax.MsgLine "                    "
		                        
		                        If clsRLOAssessment.psRLOReturnType = "FoodItem" Then
		                            WriteToErrorLog clsRLOAssessment.psCreateUser + " Keycode = " + sKeycode + ". Item to be thrown out. Item is food."
		                        Else
		                            WriteToErrorLog clsRLOAssessment.psCreateUser + " Keycode = " + sKeycode + ". Item to be thrown out. Item is cosmetics or personal care item."
		                        End If
		                        
		                        WriteToErrorLog clsRLOAssessment.psCreateUser + " Keycode = " + sKeycode + ". Reason code = WO. " + _
		                                        "Quantity = 1. SOH = 0. Store count = 0."
		
					            'get SOH and AWC values from IPS
					            bRet = GetValuesFromIPS(sKeycode, lSoh, dAWC)
					
					            clsRLOAssessment.pdAWC = dAWC
		                       
		                        If clsRLOAssessment.CreateEmptyPacketsDTL(sKeycode, 1) Then
			                    End If
			                   
		                        p_miState = k_Max_DisplayWriteOffItemPromptState
		                        p_StayInLoop = True
		                        Exit Sub
		                    Else
                        	    p_miState = k_Max_ProcessCheck_ClothingFootwearHeaterXmas
                        	End If
                        Else
                            p_miState = k_Max_ProcessDisplaySaleableOnClearanceTrolley
                        End If
                            
                        p_StayInLoop = True
                        Exit Sub
                   
                    Case Else
                        clsMax.MaxErrBell
                        clsMax.MsgLine "Invalid Input"
                        clsMax.ClearAfter 7, 17
                        clsMax.AcceptStr 7, 17
                        'clsMax.MaxScan_On
                        clsMax.MaxNumeric_Off
                           
                        p_miState = k_Max_ProcessContainHazards
                        p_StayInLoop = False
                        Exit Sub
                End Select
                
            Else
                clsMax.MaxErrBell
                clsMax.MsgLine "Invalid Input"
                clsMax.ClearAfter 7, 17
                clsMax.AcceptStr 7, 17
                'clsMax.MaxScan_On
                clsMax.MaxNumeric_Off
                   
                p_miState = k_Max_ProcessContainHazards
                p_StayInLoop = False
                Exit Sub
            End If
           
            Exit Sub
                  
        Case Else
            Exit Sub
    
    End Select
    
    Exit Sub

ErrorHandler:

    HandleErrorFatal "frmMain", "ProcessContainHazards", err.Source, err.Number, err.Description

End Sub
   
Private Sub DoWork()

'function to check the states

Dim t_StayInLoop As Integer
Dim sPDTMsg      As String
Dim t_StayInLoopCount As Integer
Const k_MaxLoop = 10

On Error GoTo ErrorHandler
    
clsMax.DumpScr
clsMax.mDebug_Log "state=" & g_miState
    
t_StayInLoopCount = 0         ' To detect dead loop
t_StayInLoop = True

Do While t_StayInLoop
    
    t_StayInLoopCount = t_StayInLoopCount + 1
    If t_StayInLoopCount > k_MaxLoop Then
       g_miState = -1 * g_miState       'Force it go to Else case to log an error
    End If
    
    Select Case g_miState                                                     'what state in ?
      
      Case k_Max_DisplayItemPromptState
           Call ProcessDisplayItemPrompt(g_miState, t_StayInLoop)
           
      Case k_Max_GetUIItemState
           g_miPrevState = k_Max_GetUIItemState
           Call ProcessItem(g_miState, t_StayInLoop, msInputStr)
           
      Case k_Max_DisplayRecallItemPromptState
           Call ProcessDisplayRecallItemPrompt(g_miState, t_StayInLoop)
           
      Case k_Max_GetUIRecallItemState
           Call ProcessRecallItem(g_miState, t_StayInLoop, msInputStr)

      Case k_Max_DisplayWriteOffItemPromptState
           Call ProcessDisplayWriteOffItemPrompt(g_miState, t_StayInLoop)
           
      Case k_Max_GetUIWriteOffItemState
           Call ProcessWriteOffItem(g_miState, t_StayInLoop, msInputStr)

      Case k_Max_DisplayClearanceItemPromptState
           Call ProcessDisplayClearanceItemPrompt(g_miState, t_StayInLoop)
           
      Case k_Max_GetUIClearanceItemState
           Call ProcessClearanceItem(g_miState, t_StayInLoop, msInputStr)
      
      Case k_Max_ProcessDisplayContainHazards
           Call ProcessDisplayContainHazards(g_miState, t_StayInLoop)
           
      Case k_Max_ProcessContainHazards
           Call ProcessContainHazards(g_miState, t_StayInLoop, msInputStr)
           
      Case k_Max_ProcessDisplaySaleableOnClearanceTrolley
           Call ProcessDisplaySaleableOnClearanceTrolley(g_miState, t_StayInLoop)
      
      Case k_Max_ProcessCheck_ClothingFootwearHeaterXmas
           Call ProcessCheck_ClothingFootwearHeaterXmas(g_miState, t_StayInLoop)
      
      Case k_Max_DisplaySCMPromptState
           sPDTMsg = clsRLOAssessment.psRLOReturnType + " Pallet"
           Call DisplaySCMPrompt(g_miState, t_StayInLoop, sPDTMsg)

      Case k_Max_ProcessSCMState
           Call ProcessSCM(g_miState, t_StayInLoop, msInputStr)

      Case k_Max_DisplaySCMErrorMsgState
           sPDTMsg = clsRLOAssessment.psPDTMsg
           Call DisplaySCMError(g_miState, t_StayInLoop, sPDTMsg)
      
      Case k_Max_ProcessSCMErrorMsgState
           Call ProcessSCMError(g_miState, t_StayInLoop, msInputStr)

      Case k_Max_EndState
           DoEvents
           Sleep (1000)
           Call AbortProgram
      
      Case k_Max_ErrorState
           HandleErrorFatal "frmMain", "DoWork", err.Source, err.Number, err.Description
           t_StayInLoop = False
      
      Case Else                                                         'should not happen
           t_StayInLoop = False
           clsMax.mDebug_Log "in do work else process state=" & g_miState
           clsMax.MsgLine "LOGIC error unknown state " & g_miState
           clsMax.mDebug_Log "LOGIC error unknown state " & g_miState
           AbortProgram
   
    End Select
Loop

Exit Sub
ErrorHandler:

    HandleErrorFatal "frmMain", "DoWork", err.Source, err.Number, err.Description

End Sub

Private Sub ProcessSCMError(ByRef p_miState As Integer, ByRef p_StayInLoop As Integer, ByVal p_inputStr As Variant)
'process sell id state

Dim sItem            As String
Dim sItemDesc        As String
Dim sKeycode         As String
Dim sErrMessage      As String
Dim iRetValue        As Integer
Dim bRet             As Boolean
Dim iPos             As Integer

On Error GoTo ErrorHandler

p_StayInLoop = False
    
    Select Case msSeqType
       
      Case KEYREFRESH                                              'refreshes the screen
           clsMax.mDebug_Log "Refresh"
           clsMax.RefreshScr
           Exit Sub
          
      Case KEYEND, KEYQUIT                                          'quit returns back to aisle state
           p_miState = k_Max_DisplayItemPromptState
           p_StayInLoop = True
           Exit Sub
      
      Case KEYRETURN
           iPos = InStr(1, msRawData, ",")
            
           If iPos > 0 Then
               sItem = Left(msRawData, iPos - 1)
           Else
               sItem = Trim(p_inputStr)
           End If

           If Len(sItem) = 0 Then
               p_miState = k_Max_DisplaySCMPromptState
               p_StayInLoop = True
           Else
               clsMax.MaxWarnBell
               p_miState = k_Max_DisplaySCMErrorMsgState
               p_StayInLoop = True
           End If
           Exit Sub
                  
      Case Else
           Exit Sub
    
    End Select
    
    Exit Sub
ErrorHandler:

    HandleErrorFatal "frmMain", "ProcessSCMError", err.Source, err.Number, err.Description

End Sub

Private Sub ProcessClearanceItem(ByRef p_miState As Integer, ByRef p_StayInLoop As Integer, ByVal p_inputStr As Variant)
'process sell id state

Dim sItem            As String
Dim sItemDesc        As String
Dim sKeycode         As String
Dim sErrMessage      As String
Dim iRetValue        As Integer
Dim bRet             As Boolean
Dim iPos             As Integer

On Error GoTo ErrorHandler

p_StayInLoop = False
    
    Select Case msSeqType
       
      Case KEYREFRESH                                              'refreshes the screen
           clsMax.mDebug_Log "Refresh"
           clsMax.RefreshScr
           Exit Sub
          
      Case KEYEND, KEYQUIT                                          'quit returns back to aisle state
           p_miState = k_Max_DisplayItemPromptState
           p_StayInLoop = True
           Exit Sub
      
      Case KEYRETURN
           iPos = InStr(1, msRawData, ",")
            
           If iPos > 0 Then
               sItem = Left(msRawData, iPos - 1)
           Else
               sItem = Trim(p_inputStr)
           End If

           If Len(sItem) = 0 Then
              p_miState = k_Max_DisplayItemPromptState
              p_StayInLoop = True
           Else
               clsMax.MaxWarnBell
               p_miState = k_Max_DisplayClearanceItemPromptState
               p_StayInLoop = True
           End If
           Exit Sub
                  
      Case Else
           Exit Sub
    
    End Select
    
    Exit Sub
ErrorHandler:

    HandleErrorFatal "frmMain", "ProcessClearanceItem", err.Source, err.Number, err.Description

End Sub
   
Private Sub ProcessSCM(ByRef p_miState As Integer, ByRef p_StayInLoop As Integer, ByVal p_inputStr As Variant)
Dim iRetValue           As Integer
Dim sKeycode            As String
Dim sAPN                As String
Dim sRLOReturnType      As String
Dim sItem_ReturnType    As String
Dim iReturnReasonCode   As Integer
Dim sSCM                As String
Dim sUSR_ID             As String
Dim sErrMessage         As String
Dim iPos                As Integer

On Error GoTo ErrorHandler

p_StayInLoop = False
    
    Select Case msSeqType
       
      Case KEYREFRESH                                              'refreshes the screen
           clsMax.mDebug_Log "Refresh"
           clsMax.RefreshScr
           Exit Sub
          
      Case KEYEND, KEYQUIT                                         'quit returns back to aisle state
           p_miState = k_Max_DisplayItemPromptState
           p_StayInLoop = True
           Exit Sub
      
      Case KEYRETURN
           iPos = InStr(1, msRawData, ",")
           
           If iPos > 0 Then
               sSCM = Left(msRawData, iPos - 1)
           Else
               sSCM = Trim(p_inputStr)
           End If

           WriteToErrorLog "ProcessSCM SCM = " + sSCM

           If Len(sSCM) <> 20 Or Not IsNumeric(sSCM) Then
                 clsMax.MaxWarnBell
                 p_miState = k_Max_DisplaySCMPromptState
                 p_StayInLoop = True
                 Exit Sub
           End If
           
           If Len(sSCM) = 20 And Mid(sSCM, 1, 3) = "275" Then
                 clsMax.MaxWarnBell
                 p_miState = k_Max_DisplaySCMPromptState
                 p_StayInLoop = True
                 Exit Sub
           End If           
           
           '
           '  Process SCM Logic
            '
           sKeycode = gtItemData.sItemID
           sAPN = gtItemData.tDetails.sRLO_APN
           sRLOReturnType = clsRLOAssessment.psRLOReturnType
           sItem_ReturnType = clsRLOAssessment.psItem_ReturnType
           iReturnReasonCode = clsRLOAssessment.piReturnReasonCode
           sUSR_ID = clsRLOAssessment.psCreateUser
           
           WriteToErrorLog "Calling Step3 " + sKeycode + "," + sAPN + "," + sRLOReturnType + "," + sItem_ReturnType + "," + Str(iReturnReasonCode) + "," + sSCM + "," + sUSR_ID
           iRetValue = ValidateRLOItemStep3(sKeycode, sAPN, sRLOReturnType, sItem_ReturnType, iReturnReasonCode, sSCM, sUSR_ID, sErrMessage)
           WriteToErrorLog "ValidateRLOItemStep3 iRetValue = " + Str(iRetValue) + " " + sErrMessage
           clsRLOAssessment.psPDTMsg = sErrMessage
           
           If iRetValue = 0 Then
                 '
                 p_miState = k_Max_DisplayItemPromptState
                 p_StayInLoop = True
                 Exit Sub
              Else
                 '
                 ' Sound warning bell.... etc
                 '
                 clsMax.MaxWarnBell
                 'clsMax.MsgLine sErrMessage
                 p_miState = k_Max_DisplaySCMErrorMsgState
                 p_StayInLoop = True
                 Exit Sub
           End If
                  
      Case Else
           Exit Sub
    
    End Select
    
    Exit Sub
ErrorHandler:

    HandleErrorFatal "frmMain", "ProcessDisplaySCM", err.Source, err.Number, err.Description

End Sub

Private Sub ProcessCheck_ClothingFootwearHeaterXmas(ByRef p_miState As Integer, ByRef p_StayInLoop As Integer)

Dim sKeycode            As String
Dim sErrMessage         As String
Dim iRetValue           As Integer
Dim iReturnReasonCode   As Integer
Dim sRLOReturnType      As String
Dim iADJCode            As Integer

On Error GoTo ErrorHandler

    sKeycode = gtItemData.sItemID
    sRLOReturnType = ""
    iADJCode = 0
    sErrMessage = ""
           '
    iRetValue = ValidateRLOItemStep2(sKeycode, iReturnReasonCode, sRLOReturnType, iADJCode, sErrMessage)

    WriteToErrorLog "ValidateRLOItemStep2 iRetValue = " + CStr(iRetValue) + ". Keycode = " + sKeycode + ". sRLOReturnType = " + sRLOReturnType + "."

    If iRetValue = 0 Then
        If sRLOReturnType = k_Salvage Then
            clsRLOAssessment.psRLOReturnType = k_Salvage
            clsRLOAssessment.psItem_ReturnType = k_Salvage
            clsRLOAssessment.piADJCode = k_ReturnReasonCode_Salvage
            clsRLOAssessment.psPDTMsg = k_Salvage + " Pallet"
            p_miState = k_Max_DisplaySCMPromptState
            p_StayInLoop = True
        Else
            clsRLOAssessment.psRLOReturnType = k_Claimable
            clsRLOAssessment.psItem_ReturnType = k_Salvage
            clsRLOAssessment.piADJCode = k_ReturnReasonCode_Salvage
            clsRLOAssessment.psPDTMsg = "NonClaim RLO Returns"
            p_miState = k_Max_DisplaySCMPromptState
            p_StayInLoop = True
        End If
    Else
        '
        ' Should never come here.
        '
        ' Sound warning bell.... etc
        '
        clsMax.MaxWarnBell
        p_miState = k_Max_DisplayItemPromptState
        p_StayInLoop = True
        Exit Sub
    End If
    
    Exit Sub
ErrorHandler:

    HandleErrorFatal "frmMain", "ProcessCheck_ClothingFootwearHeaterXmas", err.Source, err.Number, err.Description

End Sub

Private Sub DisplaySCMPrompt(ByRef p_miState As Integer, ByRef p_StayInLoop As Integer, ByVal p_Msg As String)

On Error GoTo ErrorHandler
           
   clsMax.PrintStr 1, 1, "   RLO Assessment   "
   clsMax.PrintStr 2, 1, Space(20)
   clsMax.PrintStr 3, 1, Space(20)
   clsMax.PrintStr 4, 1, Space(20)
   clsMax.PrintStr 5, 1, Space(20)
   clsMax.PrintStr 6, 1, Space(20)
   clsMax.PrintStr 7, 1, Space(20)
   clsMax.PrintStr 8, 1, Space(20)
   
   clsMax.PrintStr 2, 1, Left("Item:" + gtItemData.sItemID + Space(20), 20)
   clsMax.PrintStr 3, 1, "APN:" + gtItemData.tDetails.sRLO_APN
   clsMax.PrintStr 4, 1, gtItemData.tDetails.sItemDesc
   clsMax.PrintStr 5, 1, "Place on            "
   clsMax.PrintStr 6, 1, Left$(p_Msg + Space(20), 20)
   clsMax.PrintStr 7, 1, "SCM:                "
   clsMax.MsgLine "Scan SCM Now        "

   clsMax.AcceptStr 7, 5
   'clsMax.MaxScan_On
   clsMax.MaxNumeric_On
      
   p_miState = k_Max_ProcessSCMState
   p_StayInLoop = False

   Exit Sub
ErrorHandler:
   
   HandleErrorFatal "frmMain", "DisplaySCMPrompt", err.Source, err.Number, err.Description

End Sub

Private Sub DisplayMainMenu(Optional intOption)
'
'displays main menu on the screen after Capacity is selected
'
On Error GoTo ErrorHandler
  
    clsMax.CurrentScreen = miEntryScreen
    
    If IsMissing(intOption) Then
        'Can only create this new screen once
        miEntryScreen = clsMax.NewScr
        clsMax.PrintTitle ("   RLO Assessment   ")
    End If
    
    clsMax.MaxNumeric_On
    '
    '  The following two statement is doing the same function !

    Call clsMax.MaxEan_On
    
    Exit Sub
ErrorHandler:
    
    HandleErrorFatal "frmMain", "DisplayMainMenu", err.Source, err.Number, err.Description

End Sub

Public Sub Form_Term_Load()
    On Error GoTo ErrorHandler
    
    clsMax.TermKeyLoad                                             'Load the terminal escape sequence
    clsMax.mDebug_Log "in form load after TermKeyLoad"
                                   
    If Command = "" Then                                           'local test station
        clsMax.mDebug_Log "in form load after command"
        mbNoArg = True
        clsMax.mDebug_Log ("listening")
        TCP1(0).LocalPort = 2005
        TCP1(0).Listen
    Else
        mbNoArg = False
        clsMax.mDebug_Log "in form load in else"
        giSockIndex = 0
        TCP1(0).RemoteHost = "LOCALHOST"
        TCP1(0).RemotePort = 2005
        clsMax.mDebug_Log ("connecting")
        clsMax.mDebug_Log "in form load after connecting"
        TCP1(0).Connect
    End If
 
    Exit Sub
ErrorHandler:

    HandleErrorFatal "frmMain", "Form_Term_Load", err.Source, err.Number, err.Description

End Sub

Private Sub TCP1_Close(iIndex As Integer)
'closes the connection
    On Error GoTo ErrorHandler

    TCP1(iIndex).Close
    Unload TCP1(iIndex)
    giSockIndex = giSockIndex - 1
    AbortProgram

    Exit Sub
ErrorHandler:

    LogError "frmMain", "TCP1_Close", False
    giSockIndex = giSockIndex - 1
    AbortProgram

End Sub

Private Sub TCP1_ConnectionRequest(iIndex As Integer, ByVal RequestID As Long)
'function to connect to max and display the menu and set the state to plu
'Last Modified date: 26/08/97

Dim t_dummy As Integer
    
On Error GoTo ErrorHandler
        
    g_max_RequestID = RequestID
    
    With clsMax
        .mDebug_Log ("connectionrequest")
        giSockIndex = giSockIndex + 1
        Load TCP1(giSockIndex)                                       'create a new instance to accept request
        TCP1(giSockIndex).Accept RequestID
        .mDebug_Log "in tpc connection request state=" & g_miState
        .ClearScr
    End With
    
    g_MaxBlocking = False
    
    Call DisplayMainMenu
    Call ProcessDisplayItemPrompt(g_miState, t_dummy)
    
    clsStd.ShowGUI = False
    With clsRLOAssessment
        .psCreateUser = clsStd.UserID
        .piStoreID = clsStd.sToreID
        .psStoreName = "Store " + Str(.piStoreID)
        .psDescription = "PDE Migration RA"
    End With
    clsMax.mDebug_Log ("ConnectionRequest Finished")
    
    Exit Sub
ErrorHandler:

    HandleErrorFatal "frmMain", "TCP1_ConnectionRequest", err.Source, err.Number, err.Description

End Sub

Private Sub TCP1_Connect(index As Integer)
'Last Modified: <Fariba Mokarram> 2/06/97
Dim t_dummy As Integer

On Error GoTo ErrorHandler

    With clsMax
        .mDebug_Log ("Connected")
        TCP1(index).SendData Command & vbCrLf                       'send command arg back
        .ClearScr
    End With
    
    Call DisplayMainMenu
    Call ProcessDisplayItemPrompt(g_miState, t_dummy)

    clsStd.ShowGUI = False
    With clsRLOAssessment
        .psCreateUser = clsStd.UserID
        .piStoreID = clsStd.sToreID
        .psStoreName = "Store " + Str(.piStoreID)
        .psDescription = "PDE Migration RA"
    End With
    clsMax.mDebug_Log ("Connection Finished")
    
    Exit Sub
ErrorHandler:
    
    HandleErrorFatal "frmMain", "TCP1_Connect", err.Source, err.Number, err.Description

End Sub

Private Sub TCP1_DataArrival(index As Integer, ByVal bytesTotal As Long)
Dim bFuncKey As Boolean
Dim bcontinue As Boolean
    
On Error GoTo ErrorHandler
    
    bcontinue = True
    
    Do While bcontinue = True
       
       bcontinue = False
       
       If g_MaxBlocking = True Then
           clsMax.MsgLine "Please Wait " + Format(Now, "hh:mm:ss")
           Exit Sub
       End If
       
       Call Max_DataArrival(index, bcontinue)
    
    Loop

    Exit Sub
ErrorHandler:

    HandleErrorFatal "frmMain", "TCP1_DataArrival", err.Source, err.Number, err.Description

End Sub

Private Sub Max_DataArrival(ByRef index As Integer, ByRef bcontinue As Boolean)
   
Dim bFuncKey As Boolean
    
On Error GoTo ErrorHandler
    
    frmMain.TCP1(index).GetData mvtGetData, vbString               'invoke getdata method
    msRawData = mvtGetData
    bFuncKey = clsMax.FuncKeyPressCheck(mvtGetData, msSeqType)

    If mvtGetData <> "" Then                                       'not escape sequence &CRLF
'       clsMax.ClearMsgLine                                         'clear msgline on first hit of input
       clsMax.EchoToClient index, mvtGetData                       'echo input back
    End If
   
    If bFuncKey Then                                               'func key and CR LF entered
       g_MaxBlocking = True
       
       msInputStr = clsMax.InputStr                                'saved up for later use

       '---VK--- Stopped executing Dowork in case of Continuous Scan
       If Not mbContinuousScan Then
            Call DoWork                                             'bulk of the work done here
       End If

       clsMax.InputStr = ""                                        'reset after CR or func keys
       
       g_MaxBlocking = False
       If frmMain.TCP1(index).BytesReceived > 0 Then
          bcontinue = True
       End If
       
       mbContinuousScan = False
    
    End If

    Exit Sub
ErrorHandler:

    HandleErrorFatal "frmMain", "MAX_DataArrival", err.Source, err.Number, err.Description

End Sub

Public Function GetValuesFromIPS(ByVal sKeyCode As String, ByRef lSoh As Long, ByRef dAWC As Double) As Boolean
'
'   Get SOH and AWC values from the IPS system
'
Dim sSQL As String
Dim Cqy As rdoQuery

On Error GoTo ErrorHandler
    '
    '  Set return Code
    '
    GetValuesFromIPS = False
    lSoh = 0
    dAWC = 0#
    
    Set gIPSCon = OpenIPSDSN()
    Set Cqy = New rdoQuery
    
    sSQL = "{? = call dbo.usp_IPS_GetSOHSOOSITLastReceiptDateAWC(?,?,?,?,?,?,?)}"
    
    Set Cqy = gIPSCon.CreateQuery("", sSQL)
    
    Cqy.rdoParameters(0).Direction = rdParamReturnValue
    Cqy.rdoParameters(0).Type = rdTypeINTEGER
    Cqy.rdoParameters(1).Direction = rdParamInput
    Cqy.rdoParameters(1).Type = rdTypeINTEGER
    Cqy.rdoParameters(1).Value = clsStd.StoreID
    Cqy.rdoParameters(2).Direction = rdParamInput
    Cqy.rdoParameters(2).Type = rdTypeINTEGER
    Cqy.rdoParameters(2).Value = CLng(sKeyCode)
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
        lSoh = 0
    Else
        lSoh = Cqy.rdoParameters(3).Value
    End If
    
    If IsNull(Cqy.rdoParameters(7).Value) Then
        dAWC = 0#
    Else
        dAWC = Cqy.rdoParameters(7).Value
    End If
    
    If Cqy.rdoParameters(0) <> 0 Then
        GetValuesFromIPS = False
        
        If Not gIPSCon Is Nothing Then
            CloseConnection gIPSCon
        End If
        
        Set Cqy = Nothing
        Set gIPSCon = Nothing
        LogError "frmMain", "GetValuesFromIPS", False
        err.Raise err.Number, err.Source, err.Description
        HandleErrorFatal "frmMain", "GetValuesFromIPS", err.Source, err.Number, err.Description
    Else
        GetValuesFromIPS = True
    End If
    
    If Not gIPSCon Is Nothing Then
        CloseConnection gIPSCon
    End If
    
    Set Cqy = Nothing
    Set gIPSCon = Nothing
    Exit Function
ErrorHandler:
    GetValuesFromIPS = False
    
    If Not gIPSCon Is Nothing Then
        CloseConnection gIPSCon
    End If
    
    Set Cqy = Nothing
    Set gIPSCon = Nothing
    LogError "frmMain", "GetValuesFromIPS", False
    err.Raise err.Number, err.Source, err.Description
    HandleErrorFatal "frmMain", "GetValuesFromIPS", err.Source, err.Number, err.Description
End Function
