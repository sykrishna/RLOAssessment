VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRLOAssessment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Kmart PDE Migration Project
'
'   Name:       <RLOAssessment> : clsRLOAssessment
'
Option Explicit

'Properties
Public psBatchType             As String
Public piBatchNo               As Long
Public piStoreID               As Integer
Public psStoreName             As String
Public psCreateUser            As String
Public psDescription           As String
Public pdCreateDateTime        As Date
Public pdCommitDateTime        As Date
Public psStatus                As String
Public piADJCode               As Integer
Public psRLOReturnType         As String
Public psItem_ReturnType       As String
Public piReturnReasonCode      As Integer
Public psPDTMsg                As String
Public lBatchNo                As Long
Public lSeqNo                  As Long
Public iDeptCode               As Integer
Public dSellPrice              As Double
Public sAPN                    As String
Public pdAWC                   As Double
Public pdUnitCost              As Double
Const k_BatchCompleted = "1"

'Connection variables
Private mRS                    As rdoResultset
Private piOS2FileHandle        As Integer
Private psSourceFile           As String
Private psSourceFileProcessed  As String

Private Sub Class_Initialize()

   psBatchType = "RA"
   piBatchNo = 0
   '
   'Call SQL to retrieve Store ID / Name
   '
   piStoreID = 0
   psStoreName = ""
   
   psCreateUser = ""
   psDescription = ""
   pdCreateDateTime = Now
   pdCommitDateTime = #1/1/1800#
   psStatus = "0"
   piADJCode = 0
   
   psPDTMsg = ""

End Sub

Public Function CreateEmptyPacketsBatch() As Integer

Dim sSQL            As String
Dim iRetryCount     As Integer
Dim rs              As rdoResultset
Dim sUserName       As String
Dim sUserID         As String

On Error GoTo ErrorCreateEmptyPacketsBatch

    CreateEmptyPacketsBatch = False

    sUserID = Trim(psCreateUser)
    If Len(sUserID) > 12 Then
        sUserName = Mid(sUserID, 1, 12)
    Else
        sUserName = sUserID
    End If

    sSQL = "Set Nocount On "
    sSQL = sSQL & "Insert into CML_EmptyPacket_HDR "
    sSQL = sSQL & "(Usr_ID,Creation_TimeStamp,Commit_TimeStamp,Stat)"
    sSQL = sSQL & " Values(" & SQLText(sUserName) + ","
    sSQL = sSQL & SQLText(Format(Now(), "dd mmm yyyy hh:mm:ss")) + ","
   	sSQL = sSQL & SQLText(Format(Now(), "dd mmm yyyy hh:mm:ss")) + ","
   	sSQL = sSQL & k_BatchCompleted + " ) "
    sSQL = sSQL & " Select @@Identity as BatchHDR_ID "
    sSQL = sSQL & " Set Nocount Off"

    Set rs = gIPMCon.OpenResultset(sSQL)
    
    If Not rs.EOF Then
        If IsNull(rs!Batchhdr_id) Then
            lBatchNo = 0
            CreateEmptyPacketsBatch = False
        Else
            lBatchNo = CLng(rs!Batchhdr_id)
            CreateEmptyPacketsBatch = True
        End If
    Else
        lBatchNo = 0
        CreateEmptyPacketsBatch = False
    End If

    Exit Function

ErrorCreateEmptyPacketsBatch:
    HandleErrorFatal "clsRLOAssessment", "CreateEmptyPacketsBatch", err.Source, err.Number, err.Description
End Function

Public Function CreateEmptyPacketsDTL(ByVal psKeyCode As String, ByVal plQty As Long) As Integer
'
' function to insert details of Write Off into the table
'
Dim sSQL        As String
Dim rs          As rdoResultset

On Error GoTo Error

CreateEmptyPacketsDTL = False

    lSeqNo = 0
    
    If lSeqNo = 0 Then
       '
       '    The First Time !
       '
       Call CreateEmptyPacketsBatch
    End If
    
    lSeqNo = lSeqNo + 1
    sSQL = "Insert Into CML_EmptyPacket_DTL "
    sSQL = sSQL & "(Batch_no, Seq_No, Capture_TimeStamp, Keycode, APN, Capture_QTY, Reason_Code, Dept_No, Sell_Prc, Cost_Prc, SOH, StoreCount) "
    sSQL = sSQL & " Values( "
    sSQL = sSQL & Str(lBatchNo) & ","          'Batch_no
    sSQL = sSQL & Str(lSeqNo) & ","            'Seq_No
    sSQL = sSQL & SQLText(Format(Now(), "dd mmm yyyy hh:mm:ss")) + ","       'Capture_TimeStamp
    sSQL = sSQL & psKeyCode + ","
    sSQL = sSQL & Str(sAPN) & ","                       'APN
    sSQL = sSQL & Str(plQty) & ","                      'Capture_QTY
    sSQL = sSQL & SQLText("WO") & ","            		'Reason_Code
    sSQL = sSQL & Str(iDeptCode) & ","                  'Dept_No
    sSQL = sSQL & Str(dSellPrice) & ","                 'Sell_Prc
    sSQL = sSQL & Str(pdAWC) & ","                      'Average Weighted Cost
    sSQL = sSQL & Str(0) & ","                       	'SOH
    sSQL = sSQL & Str(0) & ")"                			'StoreCount
    
    Set rs = gIPMCon.OpenResultset(sSQL)
    
    CreateEmptyPacketsDTL = True

    Exit Function
Error:
    HandleErrorFatal "clsRLOAssessment", "CreateEmptyPacketsDtl", err.Source, err.Number, err.Description
    Debug.Print sSQL
    Exit Function

End Function