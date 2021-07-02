Attribute VB_Name = "SIMConstants"

' Module : SIMConstants
' Creator : Pauline Tran 24/01/1997
'
' Overall Function:
' Module to store all Constants required within our VB Programs
'
'Last Modified:     Ian Frazer 21/11/2000 Added MSG_STOCKTAKE_PROBLEMS
'                   Ian Frazer 2/10/2000 Added TRAINING_FILE constant
'                   Ian Frazer 6/7/2000 Added NEW_ZEALAND constant
'                   Ian Frazer 3/2/2000 Added SR_PROBLEM_FILE constant
'                   Ian Frazer 9/11/1999 Add MSG_ORPHAN_KC_REPROCESSING constant
'                                        Add MSG_ORPHAN_KC constant
'                   Ian Frazer 26/10/1999 Add MSG_RECEIPTS_MISSING constant
'                   Pauline Tran 24/05/1997
'                   Pauline Tran 25/06/1997 Added constants for messages
'                   Bob Byrne 7/7/97 - ODBC connections, etc.
'                   Pauline Tran 10/07/1997 Added the Query constants
'                   Dean Lane 10/9/97   Added OrientStructure definition for clsFileIO
'                   Dean Lane 15/9/97   Add PROCESS_SUCCESS return code set by ExitProcess
'                   Dean Lane 26/9/1997 Add Host Maintenance Application and Function Id's
'                   Dean Lane 30/9/1997 Add MSG_REPORTUNAVAILABLE for ordering
'                   Dean Lane 2/10/1997 Rename SMSHostMaintApplicationId to DRSApplicationId
'                   Dean Lane 8/10/1997 Add MSG_TIMEOUT for ordering
'                   Leanne    9/10/1997 Add CCF constants
'                   Dean Lane 16/10/1997 Add MSG_FILE_NOT_FOUND
'                   Dean Lane 20/10/1997 Add MSG_TRANSMIT_FAILED;Define PasswordCode
'                   Dean Lane 27/10/1997 Add MSG_NEGATIVEVALUE and MSG_ABORTNOW
'                   Dean Lane 31/10/1997 Added Batch Maintenance messages
'                   Pauline Tran 05/11/97 Taken out CCF Constants - moved to form level
'                   Dean Lane 15/12/1997 Change SIS DSNs to be 0001 DSN 01,02,03
'                   Dean Lane 19/12/1997 Setup DSNs for existing business codes and added
'                                        a DSN 04 to be used for Gstore datbase
'                   Cindy Emery 08/01/1998 Added message constant for Conflicting functions
'                                         and General Capture Function ID Constant
'                   Weiyi Xu    09/01/1998 Added function id constant for conflicting functions
'                   Dean Lane   14/1/1998   Added ordering msg constants 529 and 530
'                   Dean Lane   15/11998    Add STOCKTAKE_PC constant
'                   Pauline Tran 21/1/1998  Put 3 CCF constants back. (Made these CCF constants Public. --WX)
'                   Dean Lane    04-2-1998  Add AWS_Calc function id
'                   Wayne Albury 5/2/1998   Added MSG_RECBEFOREDESPATCHDATE (Receival)
'                   Leanne Lim   6/2/1998   Changed Max order function id to 5
'                   Pauline Tran 09/02/1998 Added Status of Orders
'                   Nicholas Trigg 23/02/1998  added MSG_FIXTURE_TOO_LONG,
'                       MSG_KEYCODENUMERIC, MSG_KEYCODETOOLARGE,
'                       MSG_KEYCODE_INVALID , MSG_NEGQTY, MSG_QTYTOOLARGE
'                       MSG_TO_OR_FROM
'                   Wayne Albury 04/03/1998  Added Housekeeping (Function ID)
'                   Nicholas Trigg 06/03/1998 Added MSG_DOUBLE_DIGIT
'                       AND MSG_NO_STORE_DETAILS
'                   Nicholas Trigg 13/03/1998 Changed message
'                       number 722 to MSG_NO_QUOTE_SPACE from
'                       MSG_TO_OR_FROM
'                   Nicholas Trigg 16/03/1998 Added messages:
'                       MSG_ENTER_KEYCODE, MSG_ENTER_FROM_FIXTURE,
'                       MSG_ENTER_TO_FIXTURE, MSG_ENTER_QUANTITY
'                   Nicholas Trigg 16/03/1998 Changed messages:
'                       Removed message 735,
'                       Added MSG_ENTER_FROM_OR_TO_FIXTURE,
'                       Changed value of MSG_ENTER_QTY to 734
'                   Nicholas Trigg 16/03/1998 Fixed Message 733
'                   Nicholas Trigg 17/03/1998 Changed:
'                       MSG_NO_QUOTE_SPACE to MSG_NO_QUOTE
'                   Nicholas Trigg 17/03/1998 Changed:
'                       MSG_FIXTURE_TOO_LONG to
'                       MSG_COLES_FIXTURE_TOO_LONG AND Added:
'                       MSG_KMART_FIXTURE_TOO_LONG
'                   Nicholas Trigg 18/03/1998 Changed:
'                       MSG_COLES_FIXTURE_TOO_LONG to
'                       MSG_FIXTURE_TOO_LONG AND Changed:
'                       MSG_KMART_FIXTURE_TOO_LONG to
'                       MSG_DELETE_ROW
'                   Nicholas Trigg 25/03/1998 Added:
'                       MSG_QTY_NOT_NUMERIC AND
'                       MSG_QTY_GREATER_THAN_100
'                   Weiyi Xu    03/04/1998 Added Type MaxScreenItemDetail
'                   Gordon Staines 07/04/1998 Added:
'                       stMyerGraceBros to enuBusinessType
'                   Wayne Albury 08/03/1998   Added MSG_CANNOT_TRANSMIT_INPROG,MSG_CANNOT_TRANSMIT_INCOMP
'                       MSG_CAO_CARTONQTYRANGE and MSG_CAO_CARTONQTYRANGE_STK
'                   Nicholas Trigg 08/04/98 Added MSG_DUPLICATE_CAT_NO
'                   Nicholas Trigg 25/04/98 Added:
'                       MSG_NO_BATCH_MODIFY, MSG_NO_KEYCODE_MODIFY,
'                       MSG_NO_SELLPRICE_MODIFY
'                   Nicholas Trigg 25/04/98 Added:
'                       MSG_NO_RECORDS
'                   Nicholas Trigg 12/05/98 Added:
'                       MSG_FINALISED_STOCKCOUNT(777)
'                   Michael Xuereb 14/05/98 Added: ORDERINGTRANSMISSION
'                   Wayne Albury 07/07/1998 Added: MSG_TRANSMIT_ONE_ORDER (537)
'                   Wayne Albury 21/07/1998 Added: MSG_DECIMALPOINTINVALID (538)
'                   Michael Xuereb 28/07/98 Added: MSG_RECAFTERDESPATCHBYYR (257)
'                   Dean Lane   2/10/1998   Moved enuColesMerchLevel and
'                                           enuKmartMerchLevel merchandise
'                                           definitions from modStkTkGlobals.bas to this module
'                   Wayne Albury 12/10/1998 Added: 540 and 541
'                   Wayne Albury 13/10/1998 Added Housekeeping constants
'                   Wayne Albury 05/11/1998 Modified Housekeeping constants
'                   Shreekanth Sindia 10/12/1998   Waste EOW CCF Node name change
'                   Nicholas Trigg 23/03/1999 Added STOCKTAKE_ALLOWS_SIGNOFF_FINALISE_ACCESS
'                                             Added STOCKTAKE_MULTIMEDIA_TRAINING_PC = 62 'Configure and control the multimedia training application
'                   Nicholas Trigg 30/03/1999 Added MSG_APPLICATION_MULTIFUNCTION_ACCESS = 1219
'                   Weiyi Xu       06/05/1999 Changed MSG_CANNOTUPDATEDATE from 257 to 300.
'                                             Reason: 257 has been used for MSG_RECAFTERDESPATCHBYYR
'                   Damien Schaeche 17/05/1999 MSG_ORDER_WILL_BE_UPDATED = 542
'                                              MSG_NOT_ORDER_CREATOR = 543
'                                              MSG_INCORRECT_STATUS_OFORDER = 544
'                                              MSG_ORDER_HAS_BEEN_UPDATED = 545
'
'                   Alex S          06/10/1999 Added "Multipacking" Factor to StkDtlCapture
'                   V Cheung    14/01/2000  Add FUNCTION_ID as SIM_KMART_BATCH = 64
'                   Ian Frazer  22/01/2000  Added KMART_HOUSEKEEPING constant 66
'                   Jo Gault    12 May 2000 Add new MSG section for SAO (System Aided Ordering)




'************************************************************************************
'************************************************************************************
'CONSTANTS
'************************************************************************************

'************************************************************************************
'Application constants
'************************************************************************************
Public Const SIMApplicationId = 2
Public Const DRSApplicationId = 4  'DL
Public Const SMSApplicationId = 6
Public Const SharedComponentsAppID = 8

'************************************************************************************
'Function ID Constants
'************************************************************************************
Public Const MAX_WASTE = 1      'WX
Public Const HOST_MAINTENANCE = 2 'DL

Public Const MAXORDERING = 5    'LL

Public Const STOCKTAKE_COUNT_MAX = 7 'MaxStocktake-Count
Public Const STOCKTAKE_MVMNT_MAX = 8 'MaxStocktake-MVMNT
Public Const STOCKTAKE_MVMNT_PC = 9 'PC Stocktake-MVMNT
Public Const DCRECEIVAL = 10
Public Const AWS_CALC = 11 'DL

Public Const STOCKTAKE_SCOPE_PC = 50 'PC Stocktake-SCOPE
Public Const STOCKTAKE_STATUS_PC = 51 'PC Stocktake-STATUS
Public Const STOCKTAKE_ADD_BATCH_PC = 52 'PC Stocktake-ADDBATCH
Public Const STOCKTAKE_ADD_MODIFY_ITEMS_PC = 53 'PC Stocktake-MODIFY ITEMS
Public Const STOCKTAKE_EXCEPTIONS_PC = 54 'PC Stocktake-EXCEPTIONS
Public Const STOCKTAKE_FINALISATION_PC = 55 'PC Stocktake-FINALISATION
Public Const STOCKTAKE_BATCH_PROCESSING_PC = 56 'PC Stocktake-BATCH PROCESSING
Public Const STOCKTAKE_MAX_OFFLINE = 57 'PC Stocktake-Loading offline max file to server
Public Const STOCKTAKE_GET_SALES_RETURNS = 58 'Batch process to read SalesReturns file from OS/2
Public Const STOCKTAKE_NOTIFICATION_PC = 59 'Notify screen
Public Const STOCKTAKE_BACKUP_PC = 60 'Backup the stocktake tables in an adhoc manner
Public Const STOCKTAKE_TRAINING_PC = 61 'Configure and control the training application
Public Const STOCKTAKE_MULTIMEDIA_TRAINING_PC = 62 'Configure and control the multimedia training application
Public Const STOCKTAKE_ALLOWS_SIGNOFF_FINALISE_ACCESS = 63 'Part of the finalisation applications security functionality

Public Const SIM_KMART_BATCH = 64 'Kmart batch processing
Public Const KMART_HOUSEKEEPING = 66
Public Const STOCKTAKE_CHECKER = 67

Public Const ORDERTRANSMISSION = 500    'TP problably should all be changed to 502 XXX
Public Const ORDERINGTRANSMISSION = 502 'MX
Public Const SAO = 503  'this is for the main SAO ordering application:  not actually
                        'used in the code, but is used in the html page & security files.
Public Const SAO_REPORTS = 504
'Public Const SAO_ITEM_LOC_MAINT = 505  ---NO, USE SMS FuncID 161
Public Const SAO_CALC_ORDER = 506  'this is the actual ordering calculation PART of the
                                    'main SAO application, not the whole application
Public Const SAO_FCST_ORDER = 507

Public Const STOCKCOUNT_APPLY = 610  'CE
'Public Const STOCKTAKE_PC = 700 'dl

Public Const WasteEOWFunctionID = 1100   ' *** Waste EOW CCF Node name change  ** SSindia

Public Const SISHOUSEKEEPING = 1400 'WA

Public Const PROMO_GNDLA_PLAN_MAIN = 1800 'AH 25/9/98

Public Const Offline_Client_Download = 1900 'AH 31/3/99
Public Const STARTED_DOWNLOAD = 1901 'AH 31/3/99
Public Const SUCCESSFUL_DOWNLOAD = 1902 'AH 31/3/99
Public Const UNSUCCESSFUL_DOWNLOAD = 1903 'AH 31/3/99

Public Const Offline_Server_Download = 2000 'AH 31/3/99


'************************************************************************************
'VARIOUS STRING CONSTANTS
'************************************************************************************

'CCF Constants
Public Const CCF_SUCCEED = "S"
Public Const CCF_RESEND = "V"  'Resend, because fail to transmit to destination host
Public Const CCF_FAIL = "F"

'Data Source constant
Public Const INVENTORY_ODBC_1 = "SIS0001DSN01"
Public Const INVENTORY_ODBC_2 = "SIS0001DSN02"     'dl 15/12/1997
Public Const INVENTORY_ODBC_3 = "SIS0001DSN03"     'dl 15/12/1997
Public Const INVENTORY_ODBC_4 = "SIS0001DSN04"
Public Const INVENTORY_ODBC_5 = "SIS0001DSN05"
'Business Code 0002
Public Const INVENTORY_ODBC_11 = "SIS0002DSN01"
Public Const INVENTORY_ODBC_12 = "SIS0002DSN02"
Public Const INVENTORY_ODBC_13 = "SIS0002DSN03"
Public Const INVENTORY_ODBC_14 = "SIS0002DSN04"

'Business Code 0003
Public Const INVENTORY_ODBC_21 = "SIS0003DSN01"
Public Const INVENTORY_ODBC_22 = "SIS0003DSN02"
Public Const INVENTORY_ODBC_23 = "SIS0003DSN03"
Public Const INVENTORY_ODBC_24 = "SIS0003DSN04"

Public Const INVENTORY_DRS_ODBC_1 = "DRS0001DSN02"

'Functions constants
Public Const SIM_FAIL = 0
Public Const SIM_SUCCESS = 1
Public Const SIM_MsgCode = 2

'Process Return Codes
Public Const PROCESS_SUCCESS = 100           'DL (Cobol standard success value)

'GUI Constants
Public Const GRIDBACKCOLOUR = &HD2E9FF          'peach
Public Const GRIDBEVELFACECOLOUR = 8388608      'blue
Public Const GRIDROWCOLOUR_NOEDIT = &HE0E0E0    'grey
Public Const GRIDROWCOLOUR_EDIT = &HFFFFFF      'white

' API Constants used for the width of the drop down list box
Public Const CB_SETDROPPEDWIDTH = &H160
Public Const CB_FINDSTRING = &H14C

'Filter Query Specific Constants
Public Const QUERY_NO = 0
Public Const QUERY_SPECIFIC = 1
Public Const QUERY_ALL = 2

' Coles Business constants
Public Const AdvertiseSection = 1       'Section constants
Public Const BookItemSection = 2
Public Const FixturesSection = 3
Public Const StockRoomSection = 4
Public Const ExpenseSection = 5
Public Const StoreUseSection = 6

'TPC Constants
Public Const GroceryTPC = "GR"
Public Const VarietyTPC = "VA"
Public Const DairyTPC = "DY"
Public Const FreshProduceTPC = "FV"
Public Const MeatTPC = "MT"
Public Const DeliTPC = "DL"
Public Const UnknownTPC = "XX"
Public Const LiquorTPC = "LY"
Public Const BakehouseTPC = "BH"
Public Const GeneralMerchTPC = "GM"
Public Const ApparelTPC = "AP"

'Supplier Constants
Public Const XdockSupplier = "X"
Public Const DirectSupplier = "D"
Public Const WareHouseSupplier = "W"

'Status of Orders       PT 09/02/1998
Public Const INPROGRESS_STATUS = "I"    'Moved from MAXConstants
Public Const UNCOMPLETE_STATUS = "U"
Public Const COMPLETE_STATUS = "C"      'Added
Public Const TRANSMIT_STATUS = "T"
Public Const DELETE_STATUS = "D"
Public Const MODIFY_STATUS = "M"
Public Const CANCEL_STATUS = "X"

'Order status descriptions
Public Const INPROGRESS_DESC = "In Progress"
Public Const UNCOMPLETE_DESC = "Incomplete"
Public Const COMPLETE_DESC = "Untransmitted"
Public Const TRANSMIT_DESC = "Transmitted"
Public Const DELETE_DESC = "Deleted"
Public Const MODIFY_DESC = "Modifying"
Public Const CANCEL_DESC = "Cancelled"

'Various Order Constants
Public Const RosterRunoutDays = 14
Public Const NONEXPENSECODE = 9
Public Const WrappingCategory = 970     'Category Constants
Public Const StoreExpenseCategory = 987
Public Const MAX_TRAN_ID = 9999         'DL  maximum Transaction Id for ordering

'***** Order Item exception codes
Public Const INVLD_ITEM = 2             '-- Item not found
Public Const INVLD_ORDDETAILS = 3       '-- Item not found in ord_itm table
Public Const INVLD_CATEXPCD = 6         '-- Invalid category for expense code
Public Const INVLD_EXPCD = 7            '-- Invalid expense code
Public Const ITM_NOTINSTORE = 8         '-- Item was not found in stor_sell_itm table and is not valid
Public Const ITM_NOTXDOCK = 9           '-- Merch type of item is not X-Dock orderable item
Public Const ITM_NOTORDERABLE = 10      '-- Item is not orderable
Public Const OUTOFRSTR_AUTODELETE = 11  '-- Out of roster auto deleted item
Public Const OUTOFRSTR_USERDELETE = 12  '-- Out of roster user deleted item
Public Const OUTOFRSTR_CANCELORDER = 13 '-- Out of roster order cancelled
Public Const PASSWORDCODE = 20          '-- In roster - password was entered
Public Const VALIDATE_PASSWORD = 21     '-- Validation process - password required

'***** Constants that refer to what table a sell item was reference from *****
Public Const g_cHostMaint = "H"
Public Const g_cStoreSell = "S"
Public Const g_cNonStoreSell = "N"

'***** Out Of Roster Constants *****
Public Const ROSTER_PASSWORD = "P"
Public Const ROSTER_DELETE = "D"
Public Const ROSTER_CONTINUE = "A"
Public Const ROSTER_NORMAL = "N"
Public Const ROSTER_SPECIAL = "S"

'Ddatabase constants
Public Const giFIRST As Integer = 0
Public Const giNEXT As Integer = 1
Public Const giPREV As Integer = 2
Public Const giLAST As Integer = 3

'Error constants
Public Const SIS_TIMEOUT_ERR = 38

'Constants required for the hold status of items
Public Const DELETE_ACTN = "D"
Public Const MODIFY_ACTN = "M"
Public Const ADD_ACTN = "A"
Public Const CANCEL_ACTN = "C"
Public Const NOTHING_ACTN = ""

' Constants used by Print Engine window
Public Const CW_USEDEFAULT = &H80000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Public Const WS_BORDER = &H800000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000

'declaration below required for report selection window
Public gReportSelection As Integer
Public gReportSelectionCaption As String
Public Const ACTIVATEREPORT = 1
Public Const ACTIVATEDTLREPORT = 2
Public Const PRINTSCREEN = 3
Public Const STORE_CHARGE_SUMMARY_REPORT = 4   'WX

'String constants for combo box
Public Const GREATER_THAN = "greater than"
Public Const LESS_THAN = "less than"
Public Const EQUAL_TO = "equal to"

'Indicates a New Zealand Store
Public Const NEW_ZEALAND = "KZ"

'Various String Constants
Public Const TRAINING_MODE = "IN TRAINING"
Public Const SIMFrmCaption = "Inventory Management"
Public Const NO_MATCH_FOUND = " - NO MATCH FOUND"
Public Const g_cCancel = "Cancel"

'***** System Aided Ordering Process types *****
Public Const g_cMAX = "MAX"
Public Const g_cSAO = "SAO"
Public Const G_I_ADO_TIMEOUT = -2147217871

'***** Database constants *****
Public Const g_cCMSDB = "CMSDB"
Public Const g_cGSTORE = "GSTORE"

'Housekeeping constants
Public Const SISAPPS_HK = "SISHK"   'this could be altered if the stor_cd_lu table is changed
Public Const SISHKFilePath = "sis\files\backup"

'Name of flag file used to indicate that the sales and return exe (RFSU120.EXE) failed to
'complete processing of the sales and return file.
Public Const SR_PROBLEM_FILE = "RFSU120Problem"

'Name of flag file to indicate that a training stocktake is being created. When stocktake
'batch finds this file it will process training keycode and APN files
Public Const TRAINING_STOCKTAKE = "TrainingStarted"

'***** Null conversion constant(s) - these are here because most variable types
'       can't handle NULLs *****
'**NB:  THESE VALUES ARE USED THROUGHOUT SAO PROGRAMS, IN STORED PROCS & TABLES:
'       DO NOT CHANGE.
Public Const g_cNumNULL = -666
Public Const g_cDateNULL = "01 Jan 1800"


'************************************************************************************
'************************************************************************************
'USER DATA TYPES
'************************************************************************************
'Merchandise Table Level Numbers (Coles only)
Public Enum enuColesMerchLevel
    stTPC = 1
    stCatNo = 2
    stSubCatNo = 3
    stItmCls = 4
    stSubItmCls = 5
End Enum

Public Enum enuKmartMerchLevel
    stDept = 1
    stMinDept = 2
    stFamily = 3
    stClass = 4
    stSubCls = 5
    stSubSubCls = 6
End Enum

'User datatypes as required in Order
Public Type udtStorOrdHdr
    iStorOrdId          As Integer
    iMdseTypeCd         As Integer
    sDescription        As String   ' 80
    dtCreationDateTime  As Date
    sOrderStatus        As String   ' 1
    sCreationUser       As String
    dtCommitDateTime    As Date
    sCommitUser         As String
    sMerchType          As String   ' 20
End Type

Public Type udtStorOrdDtl
    lStorOrdId          As Long
    iSectionNo          As Integer
    lSellItemId         As Long
    iCartonQty          As Integer
    iSeqNo              As Integer
    iItemExceptionCode  As Integer
    sTPCTypeCode        As String   ' 2
    sStorOrdItmSrceCd   As String   ' 2
    sOrdItmSuplyCd      As String   ' 1
    sOrdItmVelcd        As String   ' 1
    iStorOrdStatCd      As Integer
    sAisleId            As String   ' 2
    sStorCdDesc         As String   ' 60
    sItmNmb             As String   ' 40
    sItmNmXb            As String   ' 40 'actually 39?
    sItemDesc           As String
    sSectionDesc        As String
End Type

'User datatypes as required in GeneralCapture
Public Type HdrCapture
    psAisleId           As String
    psAisleDesc         As String
    psCreateUser        As String
    pdCreateDateTime    As Date
    pdCommitDateTime    As Date
    piBatchNo           As Integer
    psStatus            As String
    piStockTakeNo       As Integer
    psCaptureType       As String
    psCaptureDesc       As String
End Type

Public Type DtlCapture
    plSellId           As Long
    piSeqNo            As Integer
    psName             As String
    psSizeQty          As String
    psSizeUm           As String
    psCode             As String
    pdQty              As Double
    pdModifyQty        As Double
    pdSOH              As Double
    pReason            As String              '*** Alex S, feb 2000
End Type
'User datatypes as required in Stocktake Capture
Public Type StkHdrCapture
    psLocn              As String
    psLocnDesc          As String
    psFromLocn          As String
    psToLocn            As String
    psCreateUser        As String
    pdCreateDateTime    As Date
    pdCommitDateTime    As Date
    plBatchNo           As Long
    piStatus            As Integer
    psStatus            As String
    piStockTakeNo       As Integer
    psCaptureType       As String
    psCaptureDesc       As String
End Type

Public Type StkDtlCapture
    psSellId           As String
    piSeqNo            As Integer
    piReCountNo        As Integer
    psBrand            As String
    psName             As String
    psSizeQty          As String
    psSizeUm           As String
    psCode             As String
    pdQty              As Double
    pdModifyQty        As Double
    pdSOH              As Double
    pdUnitCost         As Double
    pdSellPrice        As Double
    iScanType          As Integer
    pdItemsinPack      As Double   '*** alex S, added oct 1999
End Type

'WXu
'Type to hold a stocktake item's lastest information
Type MaxScreenItemDetail
    BatchNo     As Integer
    Status      As Integer
    LocnDescr   As String      'For field Desc on Max screen
    Locn        As String      'For field Locn on Max screen
    SellId      As String      'For field Item on Max screen
    ItemName    As String      'For displaying item name
    RecountNo   As Integer
    SellPrice   As Double
    Qty         As Double      'For field Qty on Max screen
    SeqNo       As Integer
End Type


'User datatypes as required in Waste Entry
Public Type WasteEntry
    psTPC               As String
    plSellId            As Long
    psPluId             As String
    psName              As String
    psSizeQty           As String
    psSizeUm            As String
    psReason            As String
    pdQty               As Double
End Type

'DL
'User datatypes as required by clsFileIO
Type OrientStructure
      Orientation As Long
      Pad As String * 32
End Type

Public Enum enuBusiness
    stUnknown = -1
    stColes = 1
    stLiquorland = 2
    stKMart = 3
    stbilo = 4
    stKaties = 5
    stMyerGraceBros = 6
    stOfficeworks = 7
    stRedRooster = 8
    stTarget = 9
    stFosseys = 10
    stPickNPay = 11
End Enum

Public Enum enuDBOption
    stBulkCopy = 1
    stNullByDefault = 2
    stCheckPoint = 3
    stTruncateLog = 4
    stSingleUser = 5
    stDBOOnly = 6
    stReadOnly = 7
End Enum


'************************************************************************************
'GSTORE..CML_MESSAGES
'************************************************************************************
'************************************************************************************
'MESSAGE CONSTANTS - SIMApplicationID Constants (APPLICATIONID = 2)
'************************************************************************************
'------------------------------------------------------------------------------------
'General messages - (APPLICATIONID = 2; MessageID 1 -> 100)
'------------------------------------------------------------------------------------
Public Const MSG_INVALIDSELLID = 1              'WASTE
Public Const MSG_SELLBLANK = 2                   'WASTE
Public Const MSG_MUSTBENUMERIC = 3              'WASTE,StockTake
Public Const MSG_MUSTBEALPHABETIC = 4           'WASTE, STOCKTAKE CAPTURE
Public Const MSG_INCORRECTDATEFORMAT = 5        'Receiving
Public Const MSG_STARTBEFOREENDDATE = 6         'Receiving,Waste,SOH Report,GeneralCapture
Public Const MSG_PLUBLANK = 7                    'WASTE
Public Const MSG_PROGRAMMALFUNCTION = 8         'Ordering,
Public Const MSG_INCOMPLETEINFO = 9
Public Const MSG_CANTLOADSECTION = 10
Public Const MSG_INCOMPLETEDISCARD = 11
Public Const MSG_SELLIDXNOTEXISTS = 12           'WASTE
Public Const MSG_NOITEMFOUND = 13                'WASTE
Public Const MSG_INVALIDREPORT = 14
Public Const MSG_NODISPSEQFOUND = 15
Public Const MSG_NODATES = 16
Public Const MSG_NUMERICSUPID = 17              'Receiving
Public Const MSG_NUMERICPICKNO = 18             'Receiving
Public Const MSG_NUMERICSELLID = 19             'Receiving,Waste
Public Const MSG_NUMERICPLU = 20                'Waste
Public Const MSG_NOTIMPLEMENTED = 21           'WASTE,STOCKTAKE,GENERAL CAPTURE
Public Const MSG_INVALIDSEARCHPARAMETER = 22
Public Const MSG_INVALIDPASSWORD = 23
Public Const MSG_NOPASSWORDENTERD = 24
Public Const MSG_INVALIDPLU = 25                'WASTE
Public Const MSG_INVALIDMERCH = 26
Public Const MSG_INVALIDSECTION = 27
Public Const MSG_INVALIDCARTONQTY = 28
Public Const MSG_INVALIDPRINTSEQ = 29
Public Const MSG_INVALIDISSUETYPE = 30
Public Const MSG_NOTPCFOUND = 31
Public Const MSG_NOCATEGORYFOUND = 32
Public Const MSG_PRINTERROR = 33
Public Const MSG_CCF_TRANSFER_FAIL = 34
Public Const MSG_CCF_TRANSMIT_FAIL = 35
Public Const MSG_FILE_NOT_FOUND = 36
Public Const MSG_NEGATIVEVALUE = 37
Public Const MSG_ABORTNOW = 38
Public Const MSG_CONFIRM_DELETE = 39
Public Const MSG_CONFLICT_FUNCTION = 40 'CE

Public Const MSG_QTYBLANK = 66                      'WASTE
Public Const MSG_PLUXNOTEXISTS = 67                 'WASTE
Public Const MSG_ROWSELECT = 68
Public Const MSG_NEGATIVEQTY = 69
Public Const MSG_INVALID_RANGE = 70
Public Const MSG_NODATATOREPORT = 71                'PromAdv
Public Const MSG_CONFIRM_SAVECHANGES = 72           'Save Changes permanently?
Public Const MSG_INFO_NONCRITICAL_MALFUNC = 73      'A non-critical program malfunction has occurred. \n Please report this problem to your support person.
Public Const MSG_INFO_DUPLICATE_SELLID = 74         'Duplicate Sell ID found
Public Const MSG_INFO_SAVECOMPLETE = 75             'Changes have been saved successfully
Public Const MSG_INFO_INVALID_LISTITEM = 76         'Invalid List Item Entered.'
Public Const MSG_INFO_SAVE_BEFOREREPORT = 77        'Save changes?  Unsaved changes will not be shown on the report.
Public Const MSG_INFO_CONFIRM_DELETE = 78           'Are you sure you want to delete ?
Public Const MSG_INVALID_VALUE = 79                 'Please enter a numeric value for Percentage Factor

'-----------------------------------------------------------------------------------
'Waste messages - (APPLICATIONID = 2; MessageID 101 -> 200)
'-----------------------------------------------------------------------------------
Public Const MSG_BLANKCODE = 101
Public Const MSG_SAVECHANGES = 102
Public Const MSG_QTYAMOUNT = 103
Public Const MSG_INVALIDCODE = 104
Public Const MSG_INVALIDQTY = 105
Public Const MSG_CONFIRMQTY = 106

'-------------------------------------------------------------------------------------------------------------------------
'Receival messages - (APPLICATIONID = 2; MessageID 201 -> 300)
'-------------------------------------------------------------------------------------------------------------------------
'Email messages
Public Const MSG_UPDATEERROR = 200
Public Const MSG_NOTIFYUPDATE = 201
Public Const MSG_WSC_UPDATEERROR = 202
Public Const MSG_WSC_UPDATENOTIFY = 203
Public Const MSG_RIS_UPDATENOTIFY = 204
Public Const MSG_SOH_UPDATE_FAIL = 205  'WX

'Receival messages
Public Const MSG_CANNOTREPORT = 250
Public Const MSG_NODELIVERYDETAILS = 251
Public Const MSG_NUMERICDCORDERID = 252
Public Const MSG_SUPIDNOTDECIMAL = 253
Public Const MSG_PICKNONOTDECIMAL = 254
Public Const MSG_DCORDERIDNOTDECIMAL = 255
Public Const MSG_RECBEFOREDESPATCHDATE = 256    'WA
Public Const MSG_CHARGEDATECANNOTBESMALLER = 256
Public Const MSG_RECAFTERDESPATCHBYYR = 257     'MX708
Public Const MSG_NODATEFOUND = 258

Public Const MSG_RECGREATERTHANTODAY = 260      'AH

Public Const MSG_CANNOTUPDATEDATE = 300

'-------------------------------------------------------------------------------------------------------------------------
'???? messages - (APPLICATIONID = 2; MessageID 301 -> 400)
'-------------------------------------------------------------------------------------------------------------------------


'-------------------------------------------------------------------------------------------------------------------------
'Roster messages - (APPLICATIONID = 2; MessageID 401 -> 500)
'-------------------------------------------------------------------------------------------------------------------------
Public Const MSG_ROSTERFILE = 401
Public Const MSG_HSEKEEPING_SUCCESSFUL = 402
Public Const MSG_HSEKEEPING_FAILED = 403
Public Const MSG_RUNOUT_CHECKED = 404
Public Const MSG_RUNOUT_FOUND = 405
Public Const MSG_SUCCESSFUL = 406
Public Const MSG_FAILED = 407
Public Const MSG_DELETE_FAILED = 408
Public Const MSG_DELETE_FAILED_NORECORD = 409
Public Const MSG_FAILED_NORECORD = 410
Public Const MSG_INVALID_TRANSACTIONCD = 411
Public Const MSG_INVALID_ACTIONCD = 412
Public Const MSG_MAIL_FAILED_ROSTER_RECS = 413
Public Const MSG_MAIL_RUNOUT_WARNING = 414

'-------------------------------------------------------------------------------------------------------------------------
'Order messages - (APPLICATIONID = 2; MessageID 501 -> 600)
'-------------------------------------------------------------------------------------------------------------------------
Public Const MSG_TRANSMITORDER = 501
Public Const MSG_TRANSMITOK = 502
Public Const MSG_DELETEORDER = 503
Public Const MSG_SELECTRANGE = 504              'Please select %s to %s
Public Const MSG_CANTFINDORDER = 505
Public Const MSG_SELECTSTATUSTOSEARCH = 506
Public Const MSG_ACTIONITEM = 507
Public Const MSG_NOITEMFORORDER = 508
Public Const MSG_NOTORDERABLE = 509
Public Const MSG_STOCKROOMRANGE = 510
Public Const MSG_CARTONQTYRANGE = 511
Public Const MSG_UNABLEADDITEM = 512
Public Const MSG_UNABLEUPDATEITEM = 513
Public Const MSG_QTYPOSITIVE = 514
Public Const MSG_QTYNEGATIVE = 515
Public Const MSG_SELECTITEMTODEL = 517
Public Const MSG_SELECTORDERTOMOD = 518
Public Const MSG_CREATEDUSERNUMERIC = 519
Public Const MSG_COMMITEDUSERNUMERIC = 520
Public Const MSG_ORDERITEMIDNUMERIC = 521
Public Const MSG_QUANTITYNUMERIC = 522
Public Const MSG_TRANSLOGFILENOTFOUND = 523
Public Const MSG_ITEMIDNOTDECIMAL = 524     'Used in DC Delivery as well
Public Const MSG_QTYNOTDECIMAL = 525
Public Const MSG_REPORTUNAVAILABLE = 526
Public Const MSG_TIMEOUT = 527
Public Const MSG_TRANSMIT_FAILED = 528
Public Const MSG_CANNOT_UPDATE = 529
Public Const MSG_CANNOT_TRANSMIT_MOD = 530
Public Const MSG_CANNOT_TRANSMIT_INPROG = 531
Public Const MSG_CANNOT_TRANSMIT_INCOMP = 532
Public Const MSG_CAO_CARTONQTYRANGE = 533
Public Const MSG_CAO_CARTONQTYRANGE_STK = 534
Public Const MSG_MEAT_CARTONQTYRANGE = 535
Public Const MSG_MEAT_CARTONQTYRANGE_STK = 536
Public Const MSG_TRANSMIT_ONE_ORDER = 537
Public Const MSG_DECIMALPOINTINVALID = 538
Public Const MSG_ORDTRANS_CANCELLEDBYUSER = 539
Public Const MSG_ORDER_ALREADYDELETED = 540
Public Const MSG_DELETE_ONEORDER = 541
Public Const MSG_ORDER_WILL_BE_UPDATED = 542
Public Const MSG_NOT_ORDER_CREATOR = 543
Public Const MSG_INCORRECT_STATUS_OFORDER = 544
Public Const MSG_ORDER_HAS_BEEN_UPDATED = 545
Public Const MSG_INFO_EXCL_SELLIDS = 546            'These Sell IDs will be excluded from the current and any future order for TPC - %s
Public Const MSG_CONFIRM_DELETE_ITEMS = 547         'Are you sure you want to delete the selected %s?
Public Const MSG_INFO_UPDATENEG = 548               'Negative limit has been reached
Public Const MSG_INFO_UPDATEPOS = 549               'Positive limit has been reached
Public Const MSG_CONFIRM_PROCESSEDWASTE = 550       'Have you processed all Waste Items?
Public Const MSG_CONFIRM_CONT_CALCORDER = 551       'This process will take approximately 30 minutes to complete.....
Public Const MSG_CONFIRM_CANCEL_ORDER = 552         'Are you sure you want to cancel the selected order(s)?
Public Const MSG_INFO_CANCEL_COMPORDER = 553        'Completed orders cannot be cancelled.
Public Const MSG_INFO_NOTORDERABLE = 554            'Sell ID %s is not orderable.
Public Const MSG_CONFIRM_RUNFORECAST = 555          'Confirm whether user wants to continue order calc even though forecast algorithm needs to be run
Public Const MSG_INFO_CANNOTEXIT = 556              'Cannot close SAO when orders are in progress
Public Const MSG_INFO_NOEXCL_SELLIDS = 557          'There are no excluded Sell IDs for this TPC
Public Const MSG_INFO_NO_ROSTER = 558               'No roster has been found for the selected TPC
Public Const MSG_INFO_TPCREQUIRED = 559             'Please select a TPC.
Public Const MSG_INFO_SAOORDERIDREQUIRED = 560      'You must select a SAO Order ID for the report.
Public Const MSG_INFO_MAXORDERREQUIRED = 561        'Please select at least one MAX order to compare.
Public Const MSG_INFO_NO_MAXORDER = 562             'No MAX order to compare with.
Public Const MSG_INFO_STARTDT_GREAT_ENDDT = 563     'Start date (From) can not be greater than End date (To).
Public Const MSG_INFO_CATEGORYREQUIRED = 564        'Please select a category.
Public Const MSG_INFO_TPC_INVALIDSELLID = 565       'Selected Sell ID is invalid for the current TPC.
Public Const MSG_INFO_TPC_INVALID = 566             'SAO Reports: TPC not in list.
Public Const MSG_INFO_ORDER_INVALID = 567           'SAO Reports: SAO Order not in list.
Public Const MSG_INFO_CATEGORY_INVALID = 568        'SAO Reports: Category not in list.
Public Const MSG_INFO_CATEGORY_SELECT = 569         'Please restrict your filter to less than thirty categories.
Public Const MSG_INFO_MAXORDER_EXCEEDS = 570        'Please restrict your filter to less than forty max orders.
Public Const MSG_INFO_MISSING_DATA = 571            'Missing required details (%s)
Public Const MSG_INFO_MAXSELLID_EXCEEDS = 572       'Please restrict your filter to less than twenty Sell IDs.
Public Const MSG_CONFIRM_RUNORDER = 573             'Confirm whether user wants to continue order calc although the forecast details are out of date
Public Const MSG_CANT_CALC_ORDER = 574              'Inform user they can't calculate an order due to missing forecast information
Public Const MSG_CONFIRM_TRANS_ORDER = 575          'Confirm Transmit of a single order.
Public Const MSG_CONFIRM_TRANS_ORDERS = 576         'Confirm Transmit of multiple orders.

'-------------------------------------------------------------------------------------------------------------------------
'General Capture messages - (APPLICATIONID = 2; MessageID 601 -> 700)
'-------------------------------------------------------------------------------------------------------------------------
Public Const MSG_FIVEDIGITSQTY = 601 ' FOR BOTH STOCKTAKE AND GENERAL CAPTURE
Public Const MSG_TWODIGITSQTY = 602 ' FOR BOTH STOCKTAKE AND GENERAL CAPTURE

'-------------------------------------------------------------------------------------------------------------------------
'Stocktake messages - (APPLICATIONID = 2; MessageID 701 -> 800) ***More Stocktake codes at MessageID 1200 -> 1300***
'-------------------------------------------------------------------------------------------------------------------------
Public Const MSG_AISLEID = 701
Public Const MSG_VIEW_NOTIFY_REPORT = 702
Public Const MSG_BLANKAISLE = 703
Public Const MSG_BLANKSTATUS = 704
'StockTake Init Messages
Public Const MSG_HOST_MAINT_REQD = 705
Public Const MSG_INVALID_PERIOD = 706
Public Const MSG_REINITIATE = 707
'StockTake Location Discrepancy Messages
Public Const MSG_LOC_DISC_LOAD_ERR = 708
Public Const MSG_RETRYCANCEL_TIMEOUT = 709
Public Const MSG_RETRYCANCEL_DUPBATCH = 710
Public Const MSG_LOCATION_VALID = 711               '**Double reference
Public Const MSG_INVALID_AISLEID = 711
Public Const MSG_INVALID_QTY = 712
Public Const MSG_ITEM_NOT_FOUND = 713
Public Const MSG_LOCDISC_CANCELLED = 714
Public Const MSG_LOCDISC_NOTFOUND = 715
'Stocktake Stockroom Movement
Public Const MSG_FIXTURE_TOO_LONG = 716
Public Const MSG_KEYCODENUMERIC = 717
Public Const MSG_KEYCODETOOLARGE = 718
Public Const MSG_KEYCODE_INVALID = 719
Public Const MSG_NEGQTY = 720
Public Const MSG_QTYTOOLARGE = 721
Public Const MSG_NO_QUOTE = 722
Public Const MSG_DOUBLE_DIGIT = 723
Public Const MSG_NO_STORE_DETAILS = 724
Public Const MSG_SEND_COMPLETE = 725
Public Const MSG_SALES_FILE_NA = 726
Public Const MSG_SELECT_ITEM = 727
Public Const MSG_NOTHING_TO_DO = 728
Public Const MSG_NO_DETAILS = 729
Public Const MSG_CONFIRM_ACTION = 730
Public Const MSG_ACTION_COMPLETE = 731
Public Const MSG_ENTER_KEYCODE = 732
Public Const MSG_ENTER_FROM_OR_TO_FIXTURE = 733
Public Const MSG_ENTER_QUANTITY = 734
Public Const MSG_DELETE_ROW = 735
Public Const MSG_MUST_BE_NUMERIC = 736
Public Const MSG_CONFIRM_DEL = 737
Public Const MSG_ENTER_VALUE = 738
Public Const MSG_QTY_NOT_NUMERIC = 739
Public Const MSG_QTY_GREATER_THAN_100 = 740
Public Const MSG_SIGNED_OFF = 741
Public Const MSG_MUST_BE_FINALISED = 742
Public Const MSG_FINALISE_NOT_READY = 743
Public Const MSG_DONE_OVERNIGHT = 744
Public Const MSG_NO_INVENTORY = 745
Public Const MSG_APPL_PRFL_CHG = 746
Public Const MSG_DPT_ALRDY_INCLUDED = 747
Public Const MSG_NO_BUS_DTL = 748
Public Const MSG_NO_STORE_DTL = 749
Public Const MSG_FILE_NOT_CREATED = 750
Public Const ERR_SEND_FILE_FAILED = 751
Public Const MSG_NO_PARAMS = 752
Public Const MSG_NO_EXCPTN = 753
Public Const MSG_STILL_COUNTING = 754
Public Const MSG_DUPLICATE_CAT_NO = 755
Public Const MSG_SELECT_LIKE_LOCN = 756
Public Const MSG_NO_PERMISSION = 757
Public Const MSG_MAIL_ERROR_FILE = 758
Public Const MSG_NO_BATCH_MODIFY = 759
Public Const MSG_NO_KEYCODE_MODIFY = 760
Public Const MSG_NO_SELLPRICE_MODIFY = 761
'Stocktake Reports
Public Const MSG_NO_RECORDS = 762
Public Const MSG_MISSING_FILE = 763
Public Const MSG_DECIMAL_PLACES = 764
Public Const MSG_ENTER_SELL_PRICE = 765
Public Const MSG_NOTIFY_ARRIVED = 766
Public Const MSG_QUOTES_NOT_ALLOWED = 767
Public Const MSG_MAY_NOT_FINALISE = 768
Public Const MSG_NEGSELL = 769
Public Const MSG_SELLTOOLARGE = 770
Public Const MSG_SELL_DECIMAL_PLACES = 771
Public Const MSG_CANNOTBE_MORE_FOUR_DIGITS = 772
Public Const MSG_CANNOTBEBLANK = 773
Public Const MSG_CANNOTBE_MORE_FOURTEEN_LETTERS = 774
Public Const MSG_CANNOTBE_MORE_TWO_LETTERS = 775
Public Const MSG_INCOMPLETEENTRY = 776
Public Const MSG_FINALISED_STOCKCOUNT = 777
Public Const MSG_NO_QUOTES = 778
Public Const MSG_NO_EDIT = 779
Public Const MSG_FIXTURE_NOT_NUMERIC = 780
Public Const MSG_BACKUP_OK = 781
Public Const MSG_BACKUP_FAILED = 782
Public Const MSG_TRAINING_CLEARED = 783
Public Const MSG_BATCH_SCHEDULED = 784
Public Const MSG_OVERNIGHT_COMPLETE = 785
Public Const MSG_FILE_UNAVAILABLE = 786
Public Const MSG_DO_TRAINING = 787
Public Const MSG_CLEAR_REFERENCE = 788
Public Const MSG_CHARGEFOR_ITEMS = 789
Public Const MSG_STOCKTAKE_COMPARE = 790
Public Const MSG_NEW_LOCATION = 791
Public Const MSG_LONG_DELAY = 792
Public Const MSG_REPORT_PREFACE = 793
Public Const MSG_FIRST_CHECK_LIST = 794
Public Const MSG_VERIFY_TOP_VALUES = 795            'Not Used
Public Const MSG_VERIFY_EXCPTN_ITEMS = 796          'Not Used
Public Const MSG_VERIFY_LOCATIONS = 797             'Not Used
Public Const MSG_CONFIRM_MANUAL_ADJUSTMENTS = 798   'Not Used
Public Const MSG_STOCKTAKE_PROBLEMS = 2033

'All messages numbers between these constants have been
'reserved for use by the Finalisation Project to
'create the check list for store managers prior to file sending
Public Const MSG_LAST_CHECK_LIST = 800              'Not Used

'-------------------------------------------------------------------------------------------------------------------------
'Transmission messages - (APPLICATIONID = 2; MessageID 801 -> 900)
'-------------------------------------------------------------------------------------------------------------------------
Public Const MSG_NUMERICTRANSID = 801     'Transmission Log messages
Public Const MSG_CANNOTBEDECIMAL = 802

'-------------------------------------------------------------------------------------------------------------------------
'BackStock messages - (APPLICATIONID = 2; MessageID 901 -> 1000)
'-------------------------------------------------------------------------------------------------------------------------
Public Const MSG_SELECTMERCHTYPE = 901

'-------------------------------------------------------------------------------------------------------------------------
'Promotion Orders messages - (APPLICATIONID = 2; MessageID 1001 -> 1100)
'-------------------------------------------------------------------------------------------------------------------------
Public Const MSG_RANGEINBATCH = 1001
Public Const MSG_NOSTANDARDPRICEFORITEM = 1002
Public Const MSG_NOPROMPRICEFORITEM = 1003
Public Const MSG_NOCOLUMNSELL = 1004

'-------------------------------------------------------------------------------------------------------------------------
'Batch Maintenance messages - (APPLICATIONID = 2; MessageID 1101 -> 1199)
'-------------------------------------------------------------------------------------------------------------------------
Public Const MSG_NODELETEEVENT = 1101
Public Const MSG_INVALIDID = 1102
Public Const MSG_NUMERICID = 1103
Public Const MSG_ITEMEXISTS = 1104
Public Const MSG_NODECIMAL = 1105
Public Const MSG_NONEGATIVE = 1106
Public Const MSG_NUMERICEVENTNO = 1107
Public Const MSG_CLOSESCREEN = 1108
Public Const MSG_NOITEMSTOADD = 1109
Public Const MSG_ENTERBATCHID = 1110
Public Const MSG_BATCHIDINUSE = 1111
Public Const MSG_INVALID = 1112
Public Const MSG_CANNOTFIND = 1113

'-------------------------------------------------------------------------------------------------------------------------
'Stocktake messages - (APPLICATIONID = 2; MessageID 1200 -> 1300) ***More Stocktake codes at MessageID 701 -> 800***
'-------------------------------------------------------------------------------------------------------------------------
Public Const MSG_MISSING_VALUE = 1200
Public Const MSG_RERUN_FINALISE = 1201
Public Const MSG_MISSING_ADJUSTMENT = 1202
Public Const MSG_REPORTS_WILL_BE_SENT = 1203
Public Const MSG_ACTION_REQ = 1204
Public Const MSG_SELECT_COMPARISON = 1205
Public Const MSG_MAY_NOT_REUSE = 1206
'Scope messages
Public Const MSG_INVALID_STARTDATE = 1207
Public Const MSG_INVALID_ENDDATE = 1208
Public Const MSG_DATE_RANGE_INUSE = 1209
Public Const MSG_DATE_NOT_IN_PERIOD = 1210
Public Const MSG_STOCKTAKE_CREATED = 1211
Public Const MSG_STOCKTAKE_UPDATED = 1212
Public Const MSG_PROFILE_NOTSAVED = 1213
Public Const MSG_STOCKTAKE_ELAPSED = 1214
Public Const MSG_STOCKTAKE_NO_MODIFY = 1215
Public Const MSG_ACTIVE_STOCKTAKE = 1216
Public Const MSG_NO_ITEMS_TO_COUNT = 1217
Public Const MSG_NO_SCHEDULED_DET = 1218
Public Const MSG_APPLICATION_MULTIFUNCTION_ACCESS = 1219    'The logged on user (%s) does not have permission to run the %s function.
Public Const MSG_RECEIPTS_MISSING = 1220
Public Const MSG_ERROR = 1221
Public Const MSG_ORPHAN_KC = 1222
Public Const MSG_ORPHAN_KC_REPROCESSING = 1223

'-------------------------------------------------------------------------------------------------------------------------
'Gndola End Maintenance messages - (APPLICATIONID = 2; MessageID 1801 -> 1900)
'-------------------------------------------------------------------------------------------------------------------------
Public Const MSG_DELETE_ENDNUMBER = 1801        'Not Used
Public Const MSG_COMPLETE_ALLFIELDS = 1802      'Not Used
Public Const MSG_ENTER_ENDNUMBER = 1803         'Not Used
Public Const MSG_INVALID_ENDNUMBER = 1804       'Not Used
Public Const MSG_CONSECUTIVE_ENDNUMBER = 1805   'Not Used
Public Const MSG_TIMING_LASTITEM_DELETE = 1806  'Not Used
Public Const MSG_NO_GONDOLA_ENDS = 1807         'Not Used

'-------------------------------------------------------------------------------------------------------------------------
'PDE Ofline Client Download messages - (APPLICATIONID = 2; MessageID 1901 -> 2000)
'-------------------------------------------------------------------------------------------------------------------------
Public Const MSG_PDEDOWNLOAD_START = 1901       'Not Used
Public Const MSG_PDEDOWNLOAD_FINISHED = 1902    'Not Used
Public Const MSG_PDEDOWNLOAD_ERROR = 1903       'Not Used

'-------------------------------------------------------------------------------------------------------------------------
'Merchandise Trend Report messages - (APPLICATIONID = 2; MessageID 2001 -> 2100)
'-------------------------------------------------------------------------------------------------------------------------
Public Const MSG_NOTONSTOREFILE = 2020
Public Const MSG_MATCHINGSELLIDNOTFOUND = 2021
Public Const MSG_MATCHINGPLUNOTFOUND = 2022






'************************************************************************************
'MESSAGE CONSTANTS - SMSApplicationId Constants (APPLICATIONID = 6)
'************************************************************************************
'------------------------------------------------------------------------------------
'General messages - (APPLICATIONID = 6; MessageID 1 -> 10,000)
'------------------------------------------------------------------------------------
Public Const MSG_BATCH_RUN1 = 2025
Public Const MSG_BATCH_RUN2 = 2026

Public Const MSG_FINALCOUNTQTYTOOLARGE = 2023
Public Const MSG_MISSINGREQUIREDDATA = 2024








'************************************************************************************
'SHARED COMPONENT MESSAGE CONSTANTS - SharedComponentsAppID Constants (APPLICATIONID = 8)
'************************************************************************************
'------------------------------------------------------------------------------------
'Shared Component messages - (APPLICATIONID = 8; MessageID 3000 -> 3999)
'------------------------------------------------------------------------------------
'functions missing crucial info for load or update
Public Const MSG_SHR_MISSING_PLUID = 3000
Public Const MSG_SHR_MISSING_SELLITEMID = 3001
Public Const MSG_SHR_MISSING_CATEGORYID = 3002
Public Const MSG_SHR_MISSING_SUBCATEGORY = 3003
Public Const MSG_SHR_MISSING_ITEMCLASS = 3004
Public Const MSG_SHR_MISSING_CATEGORYDESC = 3005
Public Const MSG_SHR_MISSING_REQUIREDFIELDS = 3006
'NOTE:  requires message: "Please enter a numeric value for Current (Units)"
Public Const MSG_SHR_MISSING_PRESUNITS = 3007
Public Const MSG_SHR_LOC_TYPE_REQD = 3008
Public Const MSG_SHR_LOC_CODE_REQD = 3009
Public Const MSG_SHR_SHLF_CODE_REQD = 3010
Public Const MSG_SHR_MISSING_PRES_PERCENT = 3011
'failed search codes
Public Const MSG_SHR_SELLITEMID_NOT_FOUND = 3050
Public Const MSG_SHR_RANGED_OUT = 3051
Public Const MSG_SHR_RANGED_IN = 3052
Public Const MSG_SHR_NODATAEXISTS = 3053


