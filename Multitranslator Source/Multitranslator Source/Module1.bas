Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
'-----------------------------------------------------------------------------------------------------------------------
'File Copy ----> BackupDatabase
'-----------------------------------------------------------------------------------------------------------------------

Public Const FO_MOVE As Long = &H1
Public Const FO_COPY As Long = &H2
Public Const FO_DELETE As Long = &H3
Public Const FO_RENAME As Long = &H4
Public Const FOF_MULTIDESTFILES As Long = &H1
Public Const FOF_CONFIRMMOUSE As Long = &H2
Public Const FOF_SILENT As Long = &H4
Public Const FOF_RENAMEONCOLLISION As Long = &H8
Public Const FOF_NOCONFIRMATION As Long = &H10
Public Const FOF_WANTMAPPINGHANDLE As Long = &H20
Public Const FOF_CREATEPROGRESSDLG As Long = &H0
Public Const FOF_ALLOWUNDO As Long = &H40
Public Const FOF_FILESONLY As Long = &H80
Public Const FOF_SIMPLEPROGRESS As Long = &H100
Public Const FOF_NOCONFIRMMKDIR As Long = &H200

Public Type SHFILEOPSTRUCT
     hwnd As Long
     wFunc As Long
     pFrom As String
     pTo As String
     fFlags As Long
     fAnyOperationsAborted As Long
     hNameMappings As Long
     lpszProgressTitle As String
End Type

Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
'-----------------------------------------------------------------------------------------------------------------------
'Help File
'-----------------------------------------------------------------------------------------------------------------------
Public Declare Function HTMLHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Private Const HH_DISPLAY_TOC = &H1
'-----------------------------------------------------------------------------------------------------------------------

Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_EX_STATICEDGE = &H20000
Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000
Public Const WS_CHILD = &H40000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DLGFRAME = &H400000
Public Const WS_GROUP = &H20000
Public Const WS_HSCROLL = &H100000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_SYSMENU = &H80000
Public Const WS_POPUP = &H80000000
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_TABSTOP = &H10000
Public Const WS_THICKFRAME = &H40000
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000
Public Const WM_CLOSE = &H10
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const HWND_TOPMOST& = -1
Public Const HWND_NOTOPMOST = -2
Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
'--------------------------------------------------------------
'Database
'--------------------------------------------------------------
Dim i As Integer
Dim oc As ADODB.Connection
Dim rs As ADODB.Recordset
Dim fld As ADODB.Field
Dim sql As String



Sub Load_Data_List(Tabel As String, lst As ListBox, Index As Integer)
Dim X As Integer
Set oc = New ADODB.Connection
Set rs = New ADODB.Recordset
 oc.Provider = "Microsoft.Jet.OLEDB.4.0"
Dim pp As String

On Error GoTo err

pp = App.path & "\Data.mdb"
oc.Open pp
sql = "select * from " & Tabel
rs.Open sql, oc, adOpenKeyset

If Not rs.EOF Then

For i = 0 To rs.RecordCount - 1
            lst.AddItem rs.Fields(Index)
            rs.MoveNext
Next i
End If
Set oc = Nothing
Set rs = Nothing
Exit Sub

err:
MsgBox "Data.mdb pada directory Program Files\Monsoft\MultiTranslator\ tidak ditemukan....", vbInformation, "MultiTranslator 1.0"
End
End Sub

Sub Backup_Database(frm As Form)

Dim result As Long, fileop As SHFILEOPSTRUCT
Dim path As String
Dim tanggal As String
Dim path1 As String
tanggal = Format(Date, "dd-mm-yyyy")

If Dir(App.path & "\Data.mdb") = "" Then
MsgBox "Database tidak ditemukan...", vbInformation, App.Title
Exit Sub
End If

On Error Resume Next
MkDir App.path & "\Backup Data"
Kill App.path & "\Backup Data" & "\" & tanggal & ".BAK"
path1 = App.path & "\Data.mdb"
path = App.path & "\Backup Data" & "\" & tanggal & ".BAK"

With fileop
        .hwnd = frm.hwnd
        .wFunc = FO_COPY
        .pFrom = path1 & vbNullChar & vbNullChar
        .pTo = path & vbNullChar & vbNullChar
        .fFlags = FOF_SIMPLEPROGRESS Or FOF_FILESONLY
End With
result = SHFileOperation(fileop)
If result <> 0 Then
        MsgBox err.LastDllError
        Exit Sub
Else
        If fileop.fAnyOperationsAborted <> 0 Then
                      MsgBox "Backup Data Gagal...", vbExclamation, App.Title
        Exit Sub
        End If
SaveSetting "MultiTrans", "Setting", "Backup", tanggal
End If
End Sub

Public Sub RemoveBorder(lhWnd As Long)
  
  Dim lStyle As Long

  lStyle = GetWindowLong(lhWnd, GWL_STYLE)

  lStyle = lStyle And Not (WS_BORDER Or WS_DLGFRAME Or WS_CAPTION Or WS_BORDER Or WS_SIZEBOX Or WS_THICKFRAME)

  SetWindowLong lhWnd, GWL_STYLE, lStyle

  SetWindowPos lhWnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE

End Sub
Public Sub MakeFlat(lhWnd As Long)
  
  Dim lStyle As Long

  lStyle = GetWindowLong(lhWnd, GWL_EXSTYLE)

  lStyle = lStyle And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE

  SetWindowLong lhWnd, GWL_EXSTYLE, lStyle
  RemoveBorder lhWnd
  
End Sub
Sub Jumlah_Kata()
Dim i As Integer

If Main.List1.ListIndex <= 0 Then
Main.st.Panels.Item(3).Text = "Jumlah Kata : 0"
End If

For i = 0 To Main.List1.ListCount - 1
Main.st.Panels.Item(3).Text = "Jumlah Kata : " & i + 1
Main.st.Panels.Item(1).Text = "Loading..."
Next i

Main.st.Panels.Item(1).Text = "indra_g28@yahoo.co.id"

End Sub
Public Sub HHShowContents(lhWnd As Long)
    HTMLHelp lhWnd, App.path & "\Help.chm" & "", HH_DISPLAY_TOC, 0
End Sub
