VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSOH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' This class contains routines which are used to connect SOH DB
'
Option Explicit

'public
Public rdoWpdtimeVmCon         As rdoConnection

'private
Private iVMConnectFlag        As Integer  '0: initial try; 1: connected vm01 last time; 2: connected vm02 last time
Public Sub CloseVmConnect()

    If Not rdoWpdtimeVmCon Is Nothing Then
        'close connection
        rdoWpdtimeVmCon.Close
        Set rdoWpdtimeVmCon = Nothing
    End If

End Sub

Public Function OpenVmConnect() As Boolean

'
' flag as :
'           0: if store number is even, then try wpdtime02 first
'               if store number is odd, then try wpdtime01 first
'           1: try wpdtime01 first
'           2: try wpdtime02 first

 OpenVmConnect = False
 
    
 Select Case iVMConnectFlag
    
    Case 0
          
        If ((clsStd.sToreID Mod 2) = 0) Then
            Set rdoWpdtimeVmCon = OpenVm02
        Else
            Set rdoWpdtimeVmCon = OpenVm01
        
        End If
        
        'Connect by first try? if not try second
        If rdoWpdtimeVmCon Is Nothing Then
            
            If ((clsStd.sToreID Mod 2) = 0) Then
                    Set rdoWpdtimeVmCon = OpenVm01
            Else
                    Set rdoWpdtimeVmCon = OpenVm02
            End If
            
            'Second try
            If rdoWpdtimeVmCon Is Nothing Then
                iVMConnectFlag = 0
            Else
                If ((clsStd.sToreID Mod 2) = 0) Then
                        iVMConnectFlag = 1
                Else
                        iVMConnectFlag = 2
                End If
            End If            
            
        Else
            If ((clsStd.sToreID Mod 2) = 0) Then
                    iVMConnectFlag = 2
            Else
                    iVMConnectFlag = 1
            End If
            
        End If
        
    Case 1   ' last time, openvm01 connected, so try VM01 first
        
        Set rdoWpdtimeVmCon = OpenVm01
           
        If rdoWpdtimeVmCon Is Nothing Then
            
            Set rdoWpdtimeVmCon = OpenVm02
            
            If rdoWpdtimeVmCon Is Nothing Then
                iVMConnectFlag = 0
            Else
                iVMConnectFlag = 2
            
            End If
        Else
            iVMConnectFlag = 1
            
        End If
        
        
        
    Case 2   ' last time, openvm02 connected, so try VM02 first
        
        Set rdoWpdtimeVmCon = OpenVm02
           
        If rdoWpdtimeVmCon Is Nothing Then
        
            Set rdoWpdtimeVmCon = OpenVm01
            
            If rdoWpdtimeVmCon Is Nothing Then
                iVMConnectFlag = 0
            Else
                iVMConnectFlag = 1
                
            End If
        
        Else
            iVMConnectFlag = 2
            
        End If
        
        
End Select

    If (iVMConnectFlag = 0) Then
        OpenVmConnect = False
    Else
        OpenVmConnect = True
    End If
    
    
    
    Exit Function
    
End Function

Private Sub Class_Initialize()
 
 
 iVMConnectFlag = 0


End Sub

Private Function OpenVm01() As rdoConnection
'
' Returns connection if successful in opening the connection
'  otherwise 'nothing'
'
'   ---- Allows Connection pooling
'   ---- Should be pointing to a DSN for wpdtimevm01
'

Dim cn As rdoConnection

On Error GoTo ErrorHandler
    
    Set cn = New rdoConnection
    
    cn.CursorDriver = rdUseOdbc
    
    
    cn.Connect = "dsn=" & "SISWPDTIMEVM01"
    'cn.Connect = "DSN=SISWPDTIMEVM01;UID=KmtSOH;PWD=KmtSOH_usr;"
    
    cn.EstablishConnection rdDriverNoPrompt
        
    Set OpenVm01 = cn
        
    Set cn = Nothing
        
    Exit Function
ErrorHandler:
    
    WriteToErrorLog ("clsSOH - OpenVm01 can not connect to DB: SISWPDTIMEVM01")
    Set OpenVm01 = Nothing
    Set cn = Nothing
    Exit Function
   
    
End Function

Private Function OpenVm02() As rdoConnection
'
' Returns connection if successful in opening the connection
'  otherwise 'nothing'
'
'   ---- Allows Connection pooling
'   ---- Should be pointing to a DSN for wpdtimevm02
'

Dim cn As rdoConnection

On Error GoTo ErrorHandler
    
    Set cn = New rdoConnection
    
    cn.CursorDriver = rdUseOdbc
    
    cn.Connect = "dsn=" & "SISWPDTIMEVM02"
    'cn.Connect = "DSN=SISWPDTIMEVM02;UID=KmtSOH;PWD=KmtSOH_usr;"
    
    cn.EstablishConnection rdDriverNoPrompt
        
    Set OpenVm02 = cn
        
    Set cn = Nothing
        
    Exit Function

ErrorHandler:

    WriteToErrorLog ("clsSOH - OpenVm02 can not connect to DB: SISWPDTIMEVM02")
    Set OpenVm02 = Nothing
    Set cn = Nothing
    Exit Function
    

End Function

Public Function getSOH(ByVal sToreID As String, ByVal sItemCode As String, _
                       ByRef lSOH As Long) As Boolean
'
'The function return true if DB and table found even no record
'The function return false if DB or table not found and log error
'

Dim rs                      As rdoResultset
Dim sSQL                    As String

On Error GoTo ErrorHandler

    '
    ' Clean up buffer
    '
    lSOH = 0
          
    getSOH = True
    
    sSQL = "SELECT SOH FROM SOH WHERE Str_ID = " + sToreID + " and M_KD = " + sItemCode
           
    Set rs = rdoWpdtimeVmCon.OpenResultset(sSQL, rdOpenStatic)
    
    If Not rs Is Nothing Then
        With rs
            If .RowCount > 0 Then
                .MoveFirst
                    
                If Not IsNull(!SOH) Then
                     lSOH = !SOH
                    
                End If
                
                .Close
            End If
        End With

AbandonCn:

    End If

   
    Set rs = Nothing

    Exit Function
    
ErrorHandler:
    
    Dim er As rdoError
    Dim msg As String
    msg = " clsSOH: getSOH - An error occurred :" & err & " - " & Error & vbCr
    For Each er In rdoErrors
        
        msg = msg & er.Description & ":" & er.Number & vbCr
        
        WriteToErrorLog (msg)
         
    Next er
    
    getSOH = False
    
    Resume AbandonCn
   
End Function