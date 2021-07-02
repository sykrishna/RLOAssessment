Attribute VB_Name = "MAXConstants"
' Module : SIMConstants
' Creator : Pauline Tran 08/07/1997
'
' Last Modified : Pauline Tran 08/07/1997
'                 Dean Lane 19/9/1997 DefId:100 Add MAX_AISLE_LEN
'                 Dean Lane 14/1/1998 DefId:409 Add Uncomplete/InProgress status
'
' Overall Function:
' Module to store all MAX Constants required within our VB MAX Programs

'MAX constants
Public Const ENDKEYSEQ = 9
Public Const MAXKEYSEQ = 35
Public Const MAXMAXKEYSEQ = 13
Public Const MAXCOL = 20
Public Const MAXROW = 8
Public Const MSGROW = 8
Public Const MSGCOL = 1
Public Const MAXSCR = 19

Public Const MAX_AISLE_LEN = 2 'DL

'MAX keys constants
Public Const KEYEND = "KEYEND"
Public Const KEYDELREC = "KEYREMOVE"
Public Const KEYCLRFLD = "KEYCLRFLD"
Public Const KEYPRVFLD = "FUNC-12"
Public Const KEYHELP = "FUNC-14"
Public Const KEYSEARCH = "KEYFIND"
Public Const KEYSEARCHNEXT = "PAGEDOWN"
Public Const KEYLEFT = "KEYLEFT"
Public Const KEYRIGHT = "KEYRIGHT"
Public Const KEYLIST = "FUNC-13"
Public Const KEYREFRESH = "FUNC-08"
Public Const KEYQUIT = "FUNC-15"
Public Const KEYRETURN = "CRLF"
Public Const KEYMENU = "KEYSELECT"
Public Const KEYUP = "KEYUP"
Public Const KEYDOWN = "KEYDOWN"
Public Const KEYBACKSPACE = "BS"
Public Const MSGTYP_YESNO = 0
Public Const MSGTYP_INFO = 1
Public Const KEYCLEAR_P = "CLEAR_P"
Public Const KEYCLEAR_S = "CLEAR_S"
Public Const KEYCLEAR_N = "CLEAR_N"
Public Const KEYCLEAR = vbKeyEscape
Public Const KEYCTRL_P = "CTRL-P"
Public Const KEYCTRL_A = "CTRL-A"
Public Const KEYCTRL_D = "CTRL-D"
Public Const KEYCTRL_O = "CTRL-O"
'Max Escape sequence
Public Const KEYMAXNUMERIC_ON = "MAXNUMERIC_ON"
Public Const KEYMAXNUMERIC_OFF = "MAXNUMERIC_OFF"
Public Const KEYMAXSCAN_ON = "MAXSCAN_ON"
Public Const KEYMAXSCAN_OFF = "MAXSCAN_OFF"
Public Const KEYMAXEAN_ON = "MAXEAN_ON"
Public Const KEYMAXEAN_OFF = "MAXEAN_OFF"

'Type required for the MAX Screen
Public Type WINSCR
       Start_x                As Integer
       Start_y                As Integer
       MsgCnt                 As Integer
       Cur_x                  As Integer
       Cur_y                  As Integer
       Scr(MAXROW)            As String * MAXCOL
End Type

Public StdScr(MAXSCR) As WINSCR

'Other MAX constants
Public Const WORD_DELETE = "DEL"
Public Const WORD_DONE = "*DONE*"
