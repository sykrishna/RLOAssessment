Attribute VB_Name = "modError"
' modError
'
' Module to provide extensive error log handling.
'
' Developed by: Bob Byrne - June 1997.
' Modified by:  Bob Byrne 07/07/1997 - InitializeErrorHandler.
'               Dean Lane 11/9/1997  - Write log files to ..\logs\ directory
'               Dean Lane 15/9/1997  - Add optional ShowGUI parm to LogError to hide display
'               Bob Byrne 29/9/1997  - Add support for a parent program to set the error log
'                                       for a child OCX contol.
'               Fmokarra  17/10/1997 - Changed Logerror to return error log path.
'               Dean Lane 20/10/1997 - Convert Logerror back to a sub. OpenParentErrorLogEarly
'                                       already performs last required change.
'               Bob Byrne 16/12/1997 - Alter design so that the logfile is always opened for append.
'                                       When the log grows to > 1MB in size it will be copied to a
'                                       file of filename '"ExistingLogfileName".BUP'.  Thus at worst
'                                       there would be 2MB of logfile available.  Additionally at the
'                                       first I/O request, a session seperator is written to the file;
'                                       This is to more easily define the seperate sessions.
'                                       No code changes are necessary to support any of this new
'                                       functionality.
'               Bob Byrne 17/12/1997 - LogError/CurrentEventClear:
'                                       Optional variable bShowGUI now datatype boolean (was variant)
'                                        and renamed to bShowGUIEventMsgbox to better define its purpose.
'                                       MsgBox is now AppModal with an Exclamation icon and the
'                                        AppExename as the title.
'                                      RDOErrors - SQLState severity now shown - Warnings can be ignored.
'               Bob Byrne 7/1/98     - Remove dependency on clsStd from InitializeErrorHandler.
'               Damien Turudic 15/1/98 - Added conditional statements for applications being used.
'               Bob Byrne 2/2/1998   - Ensure error log directory exists.
'                                    - Add sub 'WriteToErrorLog' to allow messages within the log as well as errors.
'                                    - Consolidate setting up log and any maintenance functionality.
'               Bob Byrne 25/2/1999  - Enhance logged message to descriminate between those error numbers where
'                                       'vbObjectError' has been used and those where it has not been used.
'                                       'vbObjectError' is used when the programmer wishes to raise a user-generated error message.
'               Bob Byrne 13/7/1999  - Add application exename & version to 'Log Opened' message.
'
' Usage:
'       To use the modError module follow these guidelines.
'
'   Normal application:
'       Call 'InitializeErrorHandler' within the 'Main' sub of modMain.
'       Call 'LogError' where needed. (Procs and Events)(will log rdo errors automatically)
'       Call 'WriteToErrorLog' where needed. (Allows info messages within the log)
'
'   Application with child OCX that includes modError within it's code too:
'       App:
'           Call 'InitializeErrorHandler' within the 'Main' sub of modMain.
'           Call 'OpenParentErrorLogEarly' to establish the error log/path.
'           Call a control's 'Property Let' method to store the ErrorLogpath.
'       Control:
'           The control's 'Property Let' method should call the modError
'               'SetControlErrorLogFile' function to establish the
'               same path for the control.  It is assumed that doing this is as
'               a result of the log file being opened.
'           Call 'LogError' where needed. (Procs and Events)(will log rdo errors automatically)
'           Call 'WriteToErrorLog' where needed. (Allows info messages within the log)
'
'
' $History: modError.bas $
'
'*****************  Version 116  *****************
'User: Tgiannon     Date: 1/03/01    Time: 11:04
'Updated in $/Applications/Data Reporting System [DRS]/Software/Planning/Source/CatComparison
'Converted back to VB5.
'
'*****************  Version 115  *****************
'User: Ntrigg       Date: 23/10/00   Time: 15:24
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/PDE_MIGRATION/DataMigration/DMSyncRequest
'For DataMigration Kmart 001.001.060 build compilation
'
'*****************  Version 114  *****************
'User: Ntrigg       Date: 23/10/00   Time: 11:40
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/PDE_MIGRATION/DataMigration/PMDMLogSvr
'For DataMigration Kmart 001.001.060 build compilation
'
'*****************  Version 113  *****************
'User: Ntrigg       Date: 22/10/00   Time: 15:58
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/PDE_MIGRATION/DataMigration/DMCleanProcess
'version 2.2.60 - check in and out to get latest version
'
'*****************  Version 112  *****************
'User: Ifrazer      Date: 8/09/00    Time: 11:45
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/CMLStocktake/StkTkFinalise
'
'*****************  Version 111  *****************
'User: Mbowden      Date: 15/08/00   Time: 10:38
'Updated in $/Applications/Financial Data Management [FDM]/Software/Source/Error Logging
'Changed  #FIN conditional compile options in WriteToErrorLog
'
'*****************  Version 110  *****************
'User: Sjabbour     Date: 18/07/00   Time: 14:57
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/waste/Waste
'Fixef speed and GST quantity calculation
'
'*****************  Version 109  *****************
'User: Sjabbour     Date: 18/07/00   Time: 14:23
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/waste/MaxWaste
'Fixed speed and GST calculation
'
'*****************  Version 108  *****************
'User: Ifrazer      Date: 12/07/00   Time: 15:59
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/CMLStocktake/StkTkBackup
'
'*****************  Version 107  *****************
'User: Mparker      Date: 7/07/00    Time: 14:54
'Updated in $/Applications/Store Merchandise System [SMS]/Software/Current/PriceChange/Source
'
'*****************  Version 106  *****************
'User: Tgiannon     Date: 6/26/00    Time: 13:30
'Updated in $/Applications/Store Planning System [SPS]/Software/Source/Forecast Manager/SPS0005 - User Forecast Sales
'Default Casual Percentage added to Forecast Form
'
'*****************  Version 105  *****************
'User: Tgiannon     Date: 6/22/00    Time: 17:24
'Updated in $/Applications/Store Planning System [SPS]/Software/Source/Common Files
'Release 4 Enhancements
'
'*****************  Version 104  *****************
'User: Tgiannon     Date: 6/19/00    Time: 14:13
'Updated in $/Applications/Store Planning System [SPS]/Software/Source/Forecast Manager/SPS0023 - Comparison Integration
'Corrected Date
'
'*****************  Version 103  *****************
'User: Tgiannon     Date: 6/19/00    Time: 14:12
'Updated in $/Applications/Store Planning System [SPS]/Software/Source/Forecast Manager/SPS0023 - Comparison Integration
'
'*****************  Version 102  *****************
'User: Tgiannon     Date: 3/19/00    Time: 13:42
'Updated in $/Applications/Store Planning System [SPS]/Software/Source/Forecast Manager/SPS0023 - Comparison Integration
'Batch Run Comparisons for all TPC.
'
'*****************  Version 101  *****************
'User: Mscales      Date: 17/05/00   Time: 16:50
'Updated in $/Applications/Store Merchandise System [SMS]/Software/Current/SMSFindOCX/Source
'Changed End to Stop. End not valid in DLL. Stop has same functionality
'in runtime
'
'*****************  Version 100  *****************
'User: Tgiannon     Date: 9/05/00    Time: 14:13
'Updated in $/Applications/Store Planning System [SPS]/Software/Source/Common Files
'Changed the program references from the "old" VB5 OCX files:
'ComCtl32.OCX and ComCt232.OCX
'To the new VB6 refernces:
'MSComCtl.OCX and MSComCt2.OCX
'
'*****************  Version 98  *****************
'User: Cpayne       Date: 1/05/00    Time: 10:27
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/GeneralCapture
'Version 6
'
'*****************  Version 96  *****************
'User: Kng          Date: 14/04/00   Time: 10:33
'Updated in $/Applications/Store Merchandise System [SMS]/Software/LabelEngine/Source
'Fixed Label Print problem due to SQL page blocking cause by large
'transaction batch (Commits every 100). Workaround by reducing commit
'batch to 1
'
'*****************  Version 95  *****************
'User: Kng          Date: 7/04/00    Time: 11:43
'Updated in $/Applications/Store Merchandise System [SMS]/Software/HostMaintenanceApply(HMApply)/Source
'
'*****************  Version 94  *****************
'User: Mxatkins     Date: 6/03/00    Time: 17:21
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/imdc0000 SIS Batch/SISBatch
'
'*****************  Version 93  *****************
'User: Mxatkins     Date: 6/03/00    Time: 17:19
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/imdc0000 SIS Batch/SISBatch
'
'*****************  Version 92  *****************
'User: Mxatkins     Date: 6/03/00    Time: 17:09
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/imdc0000 SIS Batch/SISBatch
'
'*****************  Version 91  *****************
'User: Kng          Date: 2/02/00    Time: 16:16
'Updated in $/Applications/Store Merchandise System [SMS]/Software/PriceChange/Source
'
'*****************  Version 90  *****************
'User: Mxatkins     Date: 12/01/00   Time: 11:05
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/imdc0000 SIS Batch/SisBatch2000/WSCPdSta
'
'*****************  Version 89  *****************
'User: Mxatkins     Date: 12/01/00   Time: 9:22
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/imdc0000 SIS Batch/SisBatch2000/WSCWkChk
'
'*****************  Version 88  *****************
'User: Mxatkins     Date: 12/01/00   Time: 9:14
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/imdc0000 SIS Batch/SisBatch2000/WSCPdSta
'
'*****************  Version 87  *****************
'User: Ifrazer      Date: 1/07/00    Time: 5:37p
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/CMLStocktake/StockTakeMonitor
'
'*****************  Version 86  *****************
'User: Rcuthill     Date: 7/01/00    Time: 10:32
'Updated in $/Applications/Store Merchandise System [SMS]/Software/PLUDailySales/Source/PLU Daily Sales
'
'*****************  Version 85  *****************
'User: Rcuthill     Date: 31/12/99   Time: 13:54
'Updated in $/Applications/Store Merchandise System [SMS]/Software/PLUDailySales/Source/PLU Daily Sales
'
'*****************  Version 84  *****************
'User: Rcuthill     Date: 29/11/99   Time: 14:06
'Updated in $/Applications/Store Merchandise System [SMS]/Software/ProfitRetentionSystem[PRS]/Source
'GST Amendments
'
'*****************  Version 82  *****************
'User: Ifrazer      Date: 11/30/99   Time: 2:37p
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/CMLStocktake/StkTkDeSIF
'
'*****************  Version 81  *****************
'User: Ifrazer      Date: 11/30/99   Time: 1:37p
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/CMLStocktake/MaxOffLine
'
'*****************  Version 80  *****************
'User: Askvorts     Date: 25/11/99   Time: 11:20
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/CMLStocktake/StkTkStatus
'Added DoEvents
'
'*****************  Version 79  *****************
'User: Ifrazer      Date: 11/24/99   Time: 1:05p
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/CMLStocktake/StkTkBatch
'
'*****************  Version 78  *****************
'User: Kng          Date: 23/11/99   Time: 15:41
'Updated in $/Applications/Store Merchandise System [SMS]/Software/Range/Source
'
'*****************  Version 77  *****************
'User: Kng          Date: 17/11/99   Time: 16:59
'Updated in $/Applications/Store Merchandise System [SMS]/Software/ScheduleApply/Source
'
'*****************  Version 76  *****************
'User: Ifrazer      Date: 11/09/99   Time: 3:23p
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/CMLStocktake/StockTakeMonitor
'
'*****************  Version 75  *****************
'User: Akorneyk     Date: 11/10/99   Time: 15:25
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/CMLStocktake/MaxStkTk
'
'*****************  Version 74  *****************
'User: Rdas         Date: 5/10/99    Time: 9:37
'Updated in $/Applications/Store Merchandise System [SMS]/Software/Grid Presenter/Source
'
'*****************  Version 73  *****************
'User: Jjgault      Date: 29/09/99   Time: 14:47
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/CMLStocktake/StkTkFinalise
'Added time to string printed to finalise.log
'
'*****************  Version 72  *****************
'User: Askvorts     Date: 16/09/99   Time: 14:25
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/CMLStocktake/StktkScope
'Fixed Err =40002 and more comments
'
'*****************  Version 71  *****************
'User: Askvorts     Date: 2/09/99    Time: 14:08
'Updated in $/Applications/Store Inventory System [SIS]/Software/Source/SAMIM/CMLStocktake/StktkScope
'Work arounf for Err = 40002,
'Create dummy BufferObject property to fix
'
'*****************  Version 70  *****************
'User: Kng          Date: 26/08/99   Time: 12:00
'Updated in $/STORE MERCHANDISE SYSTEMS/src/In Store Maintenance/DissipationBusinessObject/CurrentBuild
'
'*****************  Version 69  *****************
'User: Bbyrne       Date: 13-07-99   Time: 8:31
'Updated in $/Store Merchandise Systems/src/Utilities/DataDiff
'Add application exename & version to 'Log Opened' message.
'
'This will help to identify which version of the program encountered the
'error.
'
'*****************  Version 68  *****************
'User: Askvorts     Date: 23/06/99   Time: 10:52
'Updated in $/STORE INVENTORY SYSTEMS/SAMIM/CMLStocktake/StktkScope
'
'   Mistaken checkout/checkin by user
'*****************  Version 67  *****************
'User: Bbyrne       Date: 27-04-99   Time: 14:10
'Updated in $/Store Merchandise Systems/src/In Store Maintenance/Range Application
'Fixed function 'OpenParentLogEarly' to stop superfluous message being
'written to log file.
'
'*****************  Version 66  *****************
'User: Bbyrne       Date: 1-03-99    Time: 8:53
'Updated in $/Store Merchandise Systems/src/In Store Maintenance/PRSParameter
'Enhance 'LogError' to provide the user-defined message number if the
'error number contains vbObjectError and the user-supplied number
'exceeds 512.
'
'*****************  Version 64  *****************
'User: Jiliadis     Date: 3-09-98    Time: 19:20
'Updated in $/Store Inventory Systems/Inventory/src/SIMProtoType
'   Mistaken checkout/checkin by user
'*****************  Version 63  *****************
'User: Bbyrne       Date: 24-08-98   Time: 9:11
'Updated in $/Store Merchandise Systems/src/Utilities/SQLCComp
'The Date/Time stamps within modError assumed that the Regional Settings
'for the AM & PM indicators were set.
'It turns out that these are not set during the software installation,
'so the Date/Time stamps were not showing the AM/PM indicator on the
'times.
'This has now been changed to be independent of the Regional Settings.
'
'*****************  Version 62  *****************
'User: Jatabak      Date: 5-26-98    Time: 4:38p
'Updated in $/Store Administration Systems/FDM/src/DocketMaintenance
'New method called "GetErrorLogPath" is added , it returns the
'ErrorlogPath to caller.
'Due to FDM restructure the following change have been made :
'       ModError no longer needs to maintain the log file       Ignore BB
'      "InitializeErrorHandler" has been deleted                Ignore BB
'      "MaintLogFile" has been deleted
'they were called by Conditional Compliatation  Argument # fin
'               NOTE - this is outside of modError. BB
'
'*****************  Version 61  *****************
'User: Mbowden      Date: 5/26/98    Time: 10:39a
'Updated in $/Store Administration Systems/DRS/src/FinInt
''Array Dimension' Error handling included in GetVBErrType
'
'*****************  Version 60  *****************
'User: Lromagna     Date: 22/05/98   Time: 16:47
'Updated in $/Store Administration Systems/FDM/src/Float
'Added "Exit function" to prevent the code from going to the error
'handle             NOTE - this is outside of modError. BB
'
'*****************  Version 59  *****************
'User: Jatabak      Date: 5-21-98    Time: 12:17p
'Updated in $/Store Administration Systems/DRS/src/AccountReport
'a Method called GetErrorLogfilename is added
'
'*****************  Version 58  *****************
'User: Jatabak      Date: 5-21-98    Time: 10:06a
'Updated in $/Store Administration Systems/FDM/src/OfficeCount
'A method called deleteLogFile is added
'
'*****************  Version 57  *****************
'User: Mbowden      Date: 5/20/98    Time: 4:25p
'Updated in $/Store Administration Systems/DRS/src/FinInt
'   Mistaken checkout/checkin by user
'
'*****************  Version 56  *****************
'User: Joliver      Date: 19/05/98   Time: 11:50
'Updated in $/Store Administration Systems/DRS/src/FinancialAuditReport
'Added more Lost Connection RDO Error constants for use in
'GetRDOErrType()
'
'*****************  Version 55  *****************
'User: Lromagna     Date: 5/14/98    Time: 4:15p
'Updated in $/Store Administration Systems/FDM/src/Clearance
'   Mistaken checkout/checkin by user
'
'*****************  Version 54  *****************
'User: Lromagna     Date: 5/11/98    Time: 5:05p
'Updated in $/Store Administration Systems/FDM/src/Clearance
'Added a function (GetRDOErrType) to check rdo error type.
'This function will be used in all FDM and DRS applications.
'*****************  Version 53  *****************
'User: Lromagna     Date: 5/11/98    Time: 5:00p
'Updated in $/Store Administration Systems/FDM/src/Clearance
'   Mistaken checkout/checkin by user
'
'*****************  Version 52  *****************
'User: Lromagna     Date: 5/11/98    Time: 11:59a
'Updated in $/Store Administration Systems/FDM/src/Clearance
'   Mistaken checkout/checkin by user
'
'*****************  Version 51  *****************
'User: Pbryson      Date: 1-05-98    Time: 10:11
'Updated in $/Store Planning Systems/src/General  Reports/SPSREP
'Fixed a problem with the images directory path name.
'Added a description to the operator column i.e. 'Operator & Id'
'
'*****************  Version 50  *****************
'User: Gstasis      Date: 4/27/98    Time: 11:49a
'Updated in $/Store Administration Systems/DRS/src/EFTReconciliation
'Added a line in WriteToErrorLog to insert a line before logging the
'date/time to a file
'
'*****************  Version 49  *****************
'User: Gstasis      Date: 4/21/98    Time: 12:24p
'Updated in $/Store Administration Systems/FDM/src/MonitorFinancialHealth
'   Mistaken checkout/checkin by user
'
'*****************  Version 48  *****************
'User: Bbyrne       Date: 16-04-98   Time: 8:49
'Updated in $/Store Merchandise Systems/src/Host Maintenence/Apply/HMApply
'New function 'ReplaceErrorLogFileName' provided so that an application
'can have a log of its own called 'Application.log' and an errorlog
'whose filename is provided by the programmer e.g.
'ApplicationErrors.log'.
'Also the following directive was inserted: 'Option Compare Text' to
'ensure that all text comparisions within the module were
'case-insensitive unless otherwise specified.
'
'*****************  Version 47  *****************
'User: Lromagna     Date: 8-04-98    Time: 1:56p
'Updated in $/Store Administration Systems/FDM/src/FinancialCycle
'   Mistaken checkout/checkin by user
'
'*****************  Version 46  *****************
'User: Lromagna     Date: 6-03-98    Time: 9:49a
'Updated in $/Store Administration Systems/FDM/src/Clearance
'   Mistaken checkout/checkin by user
'
'*****************  Version 45  *****************
'User: Lromagna     Date: 5-03-98    Time: 5:01p
'Updated in $/Store Administration Systems/FDM/src/Clearance
'   Mistaken checkout/checkin by user
'
'*****************  Version 44  *****************
'User: Lromagna     Date: 5-03-98    Time: 4:02p
'Updated in $/Store Administration Systems/FDM/src/Clearance
'   Mistaken checkout/checkin by user
'
'*****************  Version 43  *****************
'User: Walbury      Date: 4-03-98    Time: 3:39p
'Updated in $/Store Inventory Systems/SAMIM/imdc0450
'   Mistaken checkout/checkin by user
'
'*****************  Version 42  *****************
'User: Gstasis      Date: 4-02-98    Time: 10:50a
'Updated in $/Store Administration Systems/FDM/Utility/NewStoreUtility
'   Mistaken checkout/checkin by user
'
'*****************  Version 41  *****************
'User: Gstasis      Date: 3-02-98    Time: 4:59p
'Updated in $/Store Administration Systems/FDM/Utility/NewStoreUtility
'   Mistaken checkout/checkin by user
'
'*****************  Version 40  *****************
'User: Lromagna     Date: 2-03-98    Time: 10:03a
'Updated in $/Store Administration Systems/FDM/src/Clearance
'   Mistaken checkout/checkin by user
'
'*****************  Version 39  *****************
'User: Lromagna     Date: 2-03-98    Time: 9:55a
'Updated in $/Store Administration Systems/FDM/src/Clearance
'   Mistaken checkout/checkin by user
'
'*****************  Version 38  *****************
'User: Bbyrne       Date: 2-02-98    Time: 14:12
'Updated in $/Store Merchandise Systems/src/In Store Maintenance/PRSParameter
'Added additional functionality - 'WriteToErrorLog' function to allow
'any string to be written to the error log file.  Allows program
'specific info to be written at the same time as an error.
'Log directory is now created if it does not exist.
'Also general cleanup performed.
'
'*****************  Version 37  *****************
'User: Dturudic     Date: 1-23-98    Time: 2:05p
'Updated in $/Store Administration Systems/FDM/src/Float
'   Mistaken checkout/checkin by user
'
'*****************  Version 36  *****************
'User: Dturudic     Date: 1-22-98    Time: 9:45a
'Updated in $/Store Administration Systems/FDM/src/Float
'   Mistaken checkout/checkin by user
'
'*****************  Version 35  *****************
'User: Gstasis      Date: 20-01-98   Time: 2:26p
'Updated in $/Store Administration Systems/FDM/src/Department Transfer
'   Mistaken checkout/checkin by user
'
'*****************  Version 34  *****************
'User: Dturudic     Date: 1-19-98    Time: 2:39p
'Updated in $/Store Administration Systems/FDM/src/Float
'   Mistaken checkout/checkin by user
'
'*****************  Version 33  *****************
'User: Dturudic     Date: 1-19-98    Time: 1:38p
'Updated in $/Store Administration Systems/FDM/src/Float
'   Mistaken checkout/checkin by user
'
'*****************  Version 32  *****************
'User: Dturudic     Date: 1-19-98    Time: 10:08a
'Updated in $/Store Administration Systems/FDM/src/Float
'   Mistaken checkout/checkin by user
'
'*****************  Version 31  *****************
'User: Bbyrne       Date: 7-01-98    Time: 14:34
'Updated in $/Store Merchandise Systems/src/In Store Maintenance/Multi Located
'Removed dependency on clsStd from initializeErrorHandler.
'
'*****************  Version 30  *****************
'User: Bbyrne       Date: 17-12-97   Time: 10:44
'Updated in $/Store Merchandise Systems/src/In Store Maintenance/Multi Located
' LogError/CurrentEventClear: Optional variable bShowGUI now datatype
'boolean (was variant) and renamed to bShowGUIEventMsgbox to better
'define its purpose.
'MsgBox is now AppModal with an Exclamation icon and the ApppExename as
'the title.
'RDOErrors - SQLState severity now shown - Warnings can be ignored.
'
'*****************  Version 29  *****************
'User: Bbyrne       Date: 16-12-97   Time: 14:17
'Updated in $/Store Merchandise Systems/src/In Store Maintenance/Multi Located
'Log now always opened in append mode.
'When log exceeds 1Mb in size it is copied to logfile.BUP
'A seperator line is inserted between sessions for easy identification.
'
'*****************  Version 28  *****************
'User: Gstaines     Date: 11-14-97   Time: 3:42p
'Updated in $/Store Planning Systems/src/SPS0012
'   Mistaken checkout/checkin by user
'
'*****************  Version 27  *****************
'User: Gstaines     Date: 10/22/97   Time: 10:11a
'Updated in $/Store Planning Systems/src/SPS0012
'   Mistaken checkout/checkin by user
'
'*****************  Version 26  *****************
'User: Dlane        Date: 10/20/97   Time: 11:02a
'Updated in $/Store Inventory Systems/SAMIM/mxif8100
'Change Logerror from function back to a sub type.
'
'*****************  Version 25  *****************
'User: Fmokarra     Date: 13/10/97   Time: 9:31
'Updated in $/Store Inventory Systems/SAMIM/BackStock
'   Mistaken checkout/checkin by user
'
'*****************  Version 24  *****************
'User: Bbyrne       Date: 10/10/97   Time: 7:32
'Updated in $/Store Merchandise Systems/src/In Store Maintenance/Multi Located
'Version 1.02
'
'*****************  Version 23  *****************
'User: Slie         Date: 3/10/97    Time: 16:38
'Updated in $/Store Planning Systems/src/SPSREP
'   Mistaken checkout/checkin by user
'
'*****************  Version 22  *****************
'User: Slie         Date: 30/09/97   Time: 19:39
'Updated in $/Store Planning Systems/src/Reports/Sam's Reports
'   Mistaken checkout/checkin by user
'
'*****************  Version 21  *****************
'User: Slie         Date: 30/09/97   Time: 19:30
'Updated in $/Store Planning Systems/src/Reports/Sam's Reports
'   Mistaken checkout/checkin by user
'
'*****************  Version 20  *****************
'User: Bbyrne       Date: 30/09/97   Time: 12:02
'Updated in $/Store Merchandise Systems/src/In Store Maintenance/Multi Located
'Implement source safe version.
'
'
Option Explicit
Option Compare Text     ' Text case-insensitive comparisions

Private msErrorLogFile As String    ' Storage for path to error log file
Private mbLogOpened As Boolean      ' Indicator that log has been opened at least once already
Private msCurrentEvent As String    ' Fields/flags for event handling
Private msErrorEvent As String
Private msComputerName As String    ' Storage for this computer's name
Private msDomainName As String      ' Storage for domain name
Private msUserID As String          ' Storage for user id on pc
Private mbLogRDOWarning As Boolean  ' Used for switch to write warning rdo errors to log file
Private mctlSTAMsge As Control      ' Storage of control used for mailing purposes
'Private mbIsSTAObjectSet As Boolean

'******Alex s, sep 99, storage for obj/var passed to this Mod workaround err= 40002.
'******Beware if you delete this ref then you may get err= 40002
Private mBufferObject As Variant

'Comparison Batch Run
' True = "All TPCs" was Selected
' False = "All TPCs" not selected
Public bRunBatchTPC As Boolean

' RDO error constants
Public Const RDO_CONNECTION_LOST = 59   ' Connection lost RDO error number
Public Const RDO_SERVER_LOST = 53       ' Server Not Found RDO error number.
Public Const RDO_CONNECTION_BROKEN = 4  ' Another Lost Connection Type RDO error number.
Public Const RDO_DEAD_LOCK = 1205       ' Dead lock RDO error number
Public Const RDO_TABLE_LOCK = 0         ' Table lock RDO error number
Public Const RDO_DUPLICATE_ENTRY = 2627 ' Duplicate entry RDO error number

Public Const sTABLELOCKED = "S1T00"     ' Table lock RDO error description (sqlstate).
Public Const sCOMMLINKFAIL = "08S01"    ' The network link has been broken (sqlstate).
'Visual Basic Error constants
Public Const PERMISSION_DENIED = 70     ' The file cannot be opened, the file is locked
Public Const MISS_FILE_ERR = 53         ' The file does not exist
Public Const ARRAY_DIMENSION_ERR = 9   ' Referencing a nonexistent array element.

Public Sub InitializeErrorHandler(Optional bLogRDOWarning As Boolean = True)
    ' This routine fetches the desired environment variables.
    '
    ' Revised 7/1/98 to remove a dependency on clsStd. BB.
    
    msComputerName = Environ$("COMPUTERNAME")
    msDomainName = Environ$("USERDOMAIN")
    msUserID = Environ$("USERNAME")
    
    mbLogRDOWarning = bLogRDOWarning
    
End Sub

'Public Sub InitializeSTAControl(Optional ctlStaMsg As Control)
'
'    'DT - initialise variables used within module (modError).
'    Set mctlSTAMsge = Nothing
'
'     If (ctlStaMsg Is Nothing) Then
'        mbIsSTAObjectSet = False
'        ' If control argument is missing, return a Null. ie. initialised value.
'    Else
'        ' If control argument is present, control value passed.
'        Set mctlSTAMsge = ctlStaMsg
'        mbIsSTAObjectSet = True
'    End If
'
'End Sub

Public Sub SetErrorLogFile(sFilename As String)
    ' Save the supplied error log filename/path
    ' Used if the programmer wishes to override the default logfile name.
    '      Should be called before the first attempt to open the log file.
    ' **WARNING** - use of this function overrides any testing that the path
    '               actually exists.  Logging will not occur if the path
    '               does not exist.
    
    msErrorLogFile = sFilename
    
End Sub

Public Function OpenParentErrorLogEarly() As String             ' BB 29/9/97
    ' Function to allow a parent application to open the error log early,
    '  thus also allowing it to advise any child controls of the error log path.
    ' Returns the logfile path.
    '          Bob Byrne 16/12/97   Change open operation to always open in Append mode
    '                               (Create if not exists)
    Dim iFileNumber As Integer
    Dim sLogDir As String
    
    iFileNumber = FreeFile
    SetupLog
'''    Open msErrorLogFile For Append As #iFileNumber
'''
'''    ' Put something easily reconisable between sessions
'''    If Not mbLogOpened Then
'''        Print #iFileNumber, "--------------- Log Opened at: " & Format(Now(), "dd-mmm-yyyy hh:mm:ss AM/PM") & " ---------------"
'''        mbLogOpened = True
'''    End If
'''    Close #iFileNumber
    
    OpenParentErrorLogEarly = msErrorLogFile
    
End Function

Public Sub SetControlErrorLogFile(sFilename As String)
    ' Save the supplied error log filename/path from the parent.
    
    msErrorLogFile = sFilename
    
End Sub

Public Function sGetParentDir(sCurr As String) As String
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

Public Sub LogError(sModule As String, sRoutine As String, _
                    bIsEventSubroutine As Boolean, _
                    Optional bShowGUIEventMsgbox As Boolean = True)
    '
    ' Routine which gets called when we get an unexpected error in a routine
    '  and we take the defined error-handler.
    '
    ' This routine will only displays a messagebox if the error occurs
    '  in the top level event of the call stack AND bShowEventMsgbox is true.
    '
    'Modified: Dean Lane 11/9/97    Write log files to ..\logs\ directory
    '          Dean Lane 15/9/97    Optional showing of error message - allow for batch processing
    '          Bob Byrne 29/9/97    If rdoErrors exist - log them.
    '          Bob Byrne 16/12/97   Change open operation to always open in Append mode
    '                               (Create if not exists)
    '          Bob Byrne 17/12/97   Changed vShowGUI from a variant datatype to a boolean
    '                               and renamed the variable for clarity.
    '                               Gave the messagebox an icon and made it AppModal; title is AppExeName
    '          Bob Byrne 2/2/1998   Ensure error log directory exists.
    '
    Dim sMsg As String
    Dim sLogDir As String
    Dim iFileNumber As Integer
    
    ' Build message without cr/lf for logging
    sMsg = Format(Now(), "dd-mmm-yyyy hh:mm:ss AM/PM") & " "
    #If FIN Then
        'Display Information ie. UserID and Domain in part of Log.
        sMsg = sMsg & SetUpAddtnlMsg
    #End If
    sMsg = sMsg & msComputerName & ", Module: " & sModule
    sMsg = sMsg & ", Routine: " & sRoutine
    sMsg = sMsg & ", ErrorSource: " & err.Source
    ' Customize the output message text for user-defined error numbers.         BB 24/2/99
    ' The first 512 message numbers that can be added to vbObjectError
    '  are reserved for OLE error messages.
    If (vbObjectError And err.Number) And ((err.Number - vbObjectError) > 512) Then
        sMsg = sMsg & ", ErrorNumber(USER): " & err.Number & "(" & (err.Number - vbObjectError) & ")"
        sMsg = sMsg & ", USER ErrorDescription: " & err.Description
    Else
        sMsg = sMsg & ", ErrorNumber: " & err.Number
        sMsg = sMsg & ", ErrorDescription: " & err.Description
    End If
    
    iFileNumber = FreeFile
    SetupLog
    Open msErrorLogFile For Append As #iFileNumber
    ' Put something easily reconisable between sessions (if necessary)
    If Not mbLogOpened Then
        Print #iFileNumber, "--------------- Log Opened at: " & Format(Now(), "dd-mmm-yyyy hh:mm:ss AM/PM") & " - App.Exename=" & App.EXEName & " - Version:" & App.Major & "." & App.Minor & "." & App.Revision & " ---------------"
        mbLogOpened = True
    End If
    Print #iFileNumber, sMsg
    Close #iFileNumber
    
    If bIsEventSubroutine Then
        If UCase$(msCurrentEvent) <> UCase$(sRoutine) Then
            '----- This is an error in a cascaded
            '      event.  Save the event subroutine's
            '      name for later
            msErrorEvent = sRoutine
        Else
            '----- This is an error that has occurred
            '      in the top-level event subroutine
            If bShowGUIEventMsgbox Then
                Call SupportMSG
            End If
            If msErrorEvent <> "" Then
                '----- The event occurred as a result of
                '      a cascaded event, log the error
                LogEventError sModule, msErrorEvent, sRoutine
            End If
            msCurrentEvent = ""
            msErrorEvent = ""
        End If
    End If
    
    ' See if any rdo Error is outstanding & log it .                   'BB 29/9/97
    LogRdoErrors sModule, sRoutine
    
    'Implement Mail -conditional on FIN
    #If FIN Then
        If bIsEventSubroutine Then
            If (msCurrentEvent <> "") And (msCurrentEvent <> sRoutine) And bShowGUIEventMsgbox Then
                Call SupportMSG
            End If
        End If
    #End If
    
End Sub

Sub LogEventError(sModuleName As String, sErrorEvent As String, _
                    sPassedEvent As String)
    Dim sMsg As String
    Dim iFileNumber As Integer
    
    '----- Write error to log file
    iFileNumber = FreeFile

    ' Timing dictates that the file will have already been opened.
    '  'LogError' has already opened the file
    Open msErrorLogFile For Append As #iFileNumber
    
    sMsg = Format(Now(), "dd-mmm-yyyy hh:mm:ss AM/PM") & " "
    #If FIN Then
        'Display Information ie. UserID and Domain in part of Log.
        sMsg = sMsg & SetUpAddtnlMsg
    #End If
    sMsg = sMsg & msComputerName & " An instruction in event '" & sModuleName
    sMsg = sMsg & "." & sPassedEvent & " caused a cascaded error in event subroutine "
    sMsg = sMsg & sErrorEvent & "."
    Print #iFileNumber, sMsg
    Close #iFileNumber
    
End Sub

Public Sub CurrentEventSet(sPassedEvent As String)

    If msCurrentEvent = "" Then
       msCurrentEvent = sPassedEvent
    End If
    
End Sub

Sub CurrentEventClear(sPassedEvent As String, Optional bShowGUIEventMsgbox As Boolean = True)
    '------ Clear the current event, only if
    '       the current event is not a cascaded
    '       event
    If UCase$(sPassedEvent) = UCase$(msCurrentEvent) Then
        msCurrentEvent = ""
    End If
    
    '-------If an error is outstanding then log it
    If msErrorEvent <> "" Then
        LogEventError "", msErrorEvent, sPassedEvent
        '------- If the error was not the result of a cascade,
        '        then a function request in response to an event
        '        has failed.  Inform the user and clear the storage.
        If msCurrentEvent = "" Then
            If bShowGUIEventMsgbox Then
                Call SupportMSG
            End If
            msErrorEvent = ""
        End If
    End If
    
End Sub

Public Sub LogRdoErrors(sModule As String, sRoutine As String)      'BB 29/9/97
    ' Routine  called to log all rdo errors currently outstanding.
    '
    ' rdoErrors are not shown to the user - there is no display.
    '
    Dim sMsg As String
    Dim sLogDir As String
    Dim iFileNumber As Integer
    Dim Iindex As Integer
    
    If rdoErrors.Count > 0 Then
        iFileNumber = FreeFile
        
        Open msErrorLogFile For Append As #iFileNumber
        
        For Iindex = 0 To rdoErrors.Count - 1
            If Left$(rdoErrors(Iindex).SQLState, 2) <> "01" Or mbLogRDOWarning = True Then
                ' Build message without cr/lf for logging
                sMsg = Format(Now(), "dd-mmm-yyyy hh:mm:ss AM/PM") & " "
                #If FIN Then
                    'Display Information ie. UserID and Domain in part of Log.
                    sMsg = sMsg & SetUpAddtnlMsg
                #End If
                sMsg = sMsg & msComputerName & ", Module: " & sModule
                sMsg = sMsg & ", Routine: " & sRoutine
                ' Enhance RDO Errors info.
                If Left$(rdoErrors(Iindex).SQLState, 2) = "01" Then
                    sMsg = sMsg & ", *Warning* "
                ElseIf UCase$(Left$(rdoErrors(Iindex).SQLState, 2)) = "IM" Then
                    sMsg = sMsg & ", *ODBC Implementation Error* "
                Else
                    sMsg = sMsg & ", *Error* "
                End If
                sMsg = sMsg & ", RdoErrorSource: " & rdoErrors(Iindex).Source
                sMsg = sMsg & ", RdoErrorNumber: " & rdoErrors(Iindex).Number
                sMsg = sMsg & ", RdoErrorDescription: " & rdoErrors(Iindex).Description
    
                Print #iFileNumber, sMsg
            End If
        Next
        
        Close #iFileNumber
    End If
    
End Sub

Private Sub CheckLogSize(sLogPath As String)        ' Implemented 16/12/97 BB
    Dim lSize As Long
    Dim sBUP As String
    
    If Dir(sLogPath) <> "" Then                     ' If logfile found
        lSize = FileLen(sLogPath)                   ' Get filesize
        If lSize > 1048576 Then                     ' If log is larger than 1 Mb copy to backup
            If InStr(sLogPath, ".") Then            ' Concat BUP behind "."
                sBUP = Mid(sLogPath, 1, InStr(sLogPath, ".") - 1)
            End If
            sBUP = sBUP & ".BUP"
            FileCopy sLogPath, sBUP                 ' Copy log file to .BUP
            Kill sLogPath                           ' Remove existing log file (so it can be started afresh)
        End If
    End If
    
End Sub

Public Sub TerminateMSG()
    
    'DT 15/1/98 - implemented and called on unexpected errors when system will cease operation.
    MsgBox "Due to the previous error this program will now abnormally terminate.", _
           vbCritical + vbApplicationModal
    
End Sub

Public Sub SupportMSG()

    MsgBox "A program malfunction has occurred and the " & _
            "action you requested cannot be reliably completed.  " & vbCrLf & _
            "Please report this to your support person.", _
            vbApplicationModal + vbExclamation + vbOKOnly, App.EXEName
            
End Sub

Public Sub WriteToErrorLog(sWhatToWrite As String)          ' BB 2/2/98
    Dim iFileNumber As Integer
    
    '----- Write to log file
    iFileNumber = FreeFile

    SetupLog
    
    ' Put something easily reconisable between sessions (if necessary)
    Open msErrorLogFile For Append As #iFileNumber
    If Not mbLogOpened Then
        Print #iFileNumber, ""
        #If FIN Then
            'Display Information ie. UserID and Domain in part of Log.
            Print #iFileNumber, "--------------- Log Opened at: " & Format(Now(), "dd-mmm-yyyy hh:mm:ss AM/PM") & " ---------------"
            Print #iFileNumber, SetUpAddtnlMsg & msComputerName
        #Else
            Print #iFileNumber, "--------------- Log Opened at: " & Format(Now(), "dd-mmm-yyyy hh:mm:ss AM/PM") & " - App.Exename=" & App.EXEName & " - Version:" & App.Major & "." & App.Minor & "." & App.Revision & " ---------------"
        #End If
        mbLogOpened = True
    End If
    
    #If FIN Then
        Print #iFileNumber, Format(Now(), "dd-mmm-yyyy hh:mm:ss AM/PM") & " " & sWhatToWrite
    #Else
        'jg 24/09/1999 put time in print string
        Print #iFileNumber, Format(Now(), "hh:mm:ss") & " " & sWhatToWrite
    #End If
    Close #iFileNumber
    
End Sub

Private Sub SetupLog()                                      ' BB 2/2/98
' Ensure log path exists - create if necessary.
' Build actual full path name.
' If "FIN" group do your own log maintenance.
' Rollover log if greater than 1Mb.
'
    Dim sLogDir As String
    
    If msErrorLogFile = "" Then
        sLogDir = sGetParentDir(App.Path) 'DL
        ' Ensure directory exists                           ' BB 2/2/98
        If Dir(sLogDir & "\logs", vbDirectory) = "" Then    ' Doesn't exist
            MkDir (sLogDir & "\logs")
        End If
        msErrorLogFile = sLogDir & "\logs\" & App.EXEName & ".log"
    End If
   
    CheckLogSize (msErrorLogFile)

End Sub

Private Function SetUpAddtnlMsg() As String

    SetUpAddtnlMsg = "User ID: " & msUserID & ", Domain: " & msDomainName & ", Computer: "

End Function

Public Function ReplaceErrorLogFileName(sNewName As String) As Boolean
    ' This routine provides the ability to override the default logfile name.
    ' It is not a replacement for "SetErrorLogFile" which assumes the full path
    '  of the log file is being provided.
    ' This routine assumes the normal log directory structure, but allows
    '  the logfile name to be specified.
    ' It checks to ensure that the logfilename contains a ".log' extension because
    '  all the roll-over logic assumes that extension.
    '
    ' A common use for this function is for the case where the program itself writes
    '  debugging/trace information to a logfile that conforms to the standard naming
    '  convention of "App.EXEName & ".log".  The error log file in this case could be
    '  called "App.EXEName & "Errors.log"
    '
    Dim sLogDir As String
    
    ReplaceErrorLogFileName = False
    If InStr(sNewName, ".log") = 0 Then     ' If file does not end in .log abort
        Exit Function
    End If
    sLogDir = sGetParentDir(App.Path)
    If Dir(sLogDir & "\logs", vbDirectory) = "" Then    ' Ensure directory exists
        MkDir (sLogDir & "\logs")                       ' Doesn't exist
    End If
    SetErrorLogFile (sLogDir & "\logs\" & sNewName)
    ReplaceErrorLogFileName = True
End Function

Public Function GetRDOErrType() As String
        
    '***********************
    'This function will return an rdo error number
    '***********************
    
    Dim err As rdoError
    
    For Each err In rdoErrors
        Select Case rdoErrors(err).Number
            
            Case RDO_DEAD_LOCK
                GetRDOErrType = RDO_DEAD_LOCK
                
            Case RDO_TABLE_LOCK     ' = 0 (zero)
                If err.SQLState = sTABLELOCKED Then
                    GetRDOErrType = sTABLELOCKED
                ElseIf err.SQLState = sCOMMLINKFAIL Then
                    GetRDOErrType = RDO_CONNECTION_LOST
                End If
                
            Case RDO_CONNECTION_LOST
                GetRDOErrType = RDO_CONNECTION_LOST
                
            Case RDO_SERVER_LOST
                GetRDOErrType = RDO_CONNECTION_LOST
                
            Case RDO_CONNECTION_BROKEN
                GetRDOErrType = RDO_CONNECTION_LOST
            
            Case RDO_DUPLICATE_ENTRY
                GetRDOErrType = RDO_DUPLICATE_ENTRY
        End Select
    Next err

End Function

Public Function GetVBErrType() As String

'This function will return an VB error number

        Select Case err.Number
            Case PERMISSION_DENIED
                GetVBErrType = PERMISSION_DENIED

            Case MISS_FILE_ERR
                GetVBErrType = MISS_FILE_ERR
                
            Case ARRAY_DIMENSION_ERR
                GetVBErrType = ARRAY_DIMENSION_ERR
                
        End Select

End Function
Public Function DeleteLogFile() As String
' This function deletes the log file
' it should be called in the begining of the application
' created at 12 may 1998 by John Atabak
On Error GoTo errHandle

    Call SetupLog
    ' if the log filename is been set
    If msErrorLogFile <> "" Then
        ' if the log file exist
        If Dir(msErrorLogFile) <> "" Then
            Kill (msErrorLogFile)
        End If
    End If

    Exit Function

errHandle:


End Function
    

Public Function GetErrorLogfilename() As String
    ' Function to return the logfile name and checks if we can open it
    ' created by john Atabak at 21 may 1998
On Error GoTo errHandle
        
    Dim iFileNumber As Integer
    
    iFileNumber = FreeFile
    SetupLog
    Open msErrorLogFile For Append As #iFileNumber
    Close #iFileNumber
    
    GetErrorLogfilename = msErrorLogFile
    
    Exit Function
errHandle:
   GetErrorLogfilename = ""

End Function
Public Function GetErrorLogPath() As String
    ' Function to return the logfile path
    ' created by john Atabak at 26 may 1998
On Error GoTo errHandle
        
    Dim sFileNumber As String
    Dim sFilePath As String
    
    sFileNumber = GetErrorLogfilename
    sFilePath = sGetParentDir(sFileNumber)
    
    
    Exit Function
errHandle:
   GetErrorLogPath = ""

End Function


Public Sub BufferObject(vOBJtoBuffer As Variant)

' Alex S New sub to pass object to this module
' sep 99, workaround err = 40002

Dim iFileNumber As Integer


On Error GoTo Err_hand

If IsObject(vOBJtoBuffer) Then
    Set mBufferObject = vOBJtoBuffer
Else
    mBufferObject = vOBJtoBuffer
End If


 
'------Let OS to do smth to org some sort ofdelay
        
    DoEvents
    iFileNumber = FreeFile
    SetupLog
    ' Open and Colse Log
    Open msErrorLogFile For Append As #iFileNumber
    Close #iFileNumber
'--------

Exit Sub

Err_hand:
    LogError "modError", "BufferObject", False, False
    err.Raise err.Number
End Sub



