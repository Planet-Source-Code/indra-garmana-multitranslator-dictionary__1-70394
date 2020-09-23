VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MultiTranslator"
   ClientHeight    =   5430
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8535
   Begin VB.CommandButton Command1 
      Caption         =   "Terjemah"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Translate"
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton Dummy 
      Height          =   195
      Left            =   9120
      TabIndex        =   11
      Top             =   1440
      Width           =   195
   End
   Begin MSComctlLib.StatusBar st 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   5055
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Arah Terjemah :"
            TextSave        =   "Arah Terjemah :"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Jumlah Kata :"
            TextSave        =   "Jumlah Kata :"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6840
      Top             =   8520
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   6720
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame2 
      Caption         =   "Output"
      Height          =   2055
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   8310
      Begin VB.TextBox Text2 
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   240
         Width           =   8055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input"
      Height          =   2175
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   8310
      Begin VB.TextBox Text1 
         Height          =   1815
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   240
         Width           =   8055
      End
   End
   Begin VB.ListBox List6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      ItemData        =   "Form1.frx":08CA
      Left            =   3480
      List            =   "Form1.frx":08CC
      TabIndex        =   5
      Top             =   6000
      Width           =   3255
   End
   Begin VB.ListBox List5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      ItemData        =   "Form1.frx":08CE
      Left            =   3480
      List            =   "Form1.frx":08D0
      TabIndex        =   4
      Top             =   6720
      Width           =   3255
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      ItemData        =   "Form1.frx":08D2
      Left            =   120
      List            =   "Form1.frx":08D4
      TabIndex        =   3
      Top             =   6720
      Width           =   3255
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      ItemData        =   "Form1.frx":08D6
      Left            =   120
      List            =   "Form1.frx":08D8
      TabIndex        =   2
      Top             =   6000
      Width           =   3255
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      ItemData        =   "Form1.frx":08DA
      Left            =   3480
      List            =   "Form1.frx":08DC
      TabIndex        =   1
      Top             =   7440
      Width           =   3255
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      ItemData        =   "Form1.frx":08DE
      Left            =   120
      List            =   "Form1.frx":08E0
      TabIndex        =   0
      Top             =   7440
      Width           =   3255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu Sep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuArah_Terjemah 
         Caption         =   "&Arah Terjemah"
         Begin VB.Menu mnuInd_Ing 
            Caption         =   "&Indonesia Ke Inggris"
         End
         Begin VB.Menu Sep3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuIng_Ind 
            Caption         =   "&Inggris Ke Indonesia"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu Sep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuType 
         Caption         =   "&Type Pengolahan"
      End
      Begin VB.Menu Sep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTypeCari 
         Caption         =   "Type Pencarian"
      End
      Begin VB.Menu Sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFont_Style 
         Caption         =   "&Font Style"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuEditor 
         Caption         =   "&Database Editor"
      End
      Begin VB.Menu Sep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "&Backup Database"
      End
      Begin VB.Menu Sep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore Database"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help..."
      End
      Begin VB.Menu Sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAboutMe 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Alhamdulilah
'Ucap syukur pada Allah S.W.T
'Terimakasihku untuk :
'Momon Suparman & Popong Rohayati kedua orang tuaku (i dedicated to...)
'Keluargaku & keponakanku Cherryl Aurelia dan Fiona
'mail me at : indra_g28@yahoo.co.id

Option Explicit
Dim i As Integer
Const ALPHA = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
Sub Pecah_Kalimat(txt As String)

'memecah terlebih dahulu satu kalimat menjadi kata demi kata

Dim i As Integer
Dim lcount As Long
Dim str() As String

str = Split(Trim(txt))

For lcount = 0 To UBound(str)
List3.AddItem str(lcount)
Next

End Sub
Sub Pisah_Karakter1()

'pisah karakter dengan karakter Non AlphaNumeric

    Dim i As Integer
    Dim str As String
    Dim iCntr As Integer
    Dim sWork As String, sChar As String
    Dim sOut1 As String, sOut2 As String

    
    
    str = List3.List(List3.ListIndex)
    sWork = Trim(str)
    
    If sWork = "" Then Exit Sub
    For iCntr = 1 To Len(sWork)
        sChar = Mid(sWork, iCntr, 1)
        
        
        If InStr(ALPHA, UCase(sChar)) <> 0 Then
            
            sOut1 = sOut1 & sChar

        Else
          
            sOut2 = sOut2 & sChar
            
        End If

    Next iCntr
    

   
   List4.AddItem sOut1
   List5.AddItem sOut2
   

End Sub
Public Function Pisah_Karakter2(ByVal Words As String) As String

'pisah karakter tanpa karakter Non AlphaNumeric
    
    Const SNG_SPACE = " "
    Const DBL_SPACE = "  "
    
    Dim iPos As Integer
    Dim sChar As String, sWork As String
    Dim lCntr As Long

    
    sWork = Trim(Text1.Text)
    
    
    For lCntr = 0 To 255
        If InStr(ALPHA, Chr(lCntr)) = 0 Then
            sWork = Replace(sWork, Chr(lCntr), SNG_SPACE)
        End If
    Next lCntr
    
    
    iPos = InStr(sWork, DBL_SPACE)
    While iPos > 0
        sWork = Replace(sWork, DBL_SPACE, SNG_SPACE)
        iPos = InStr(sWork, DBL_SPACE)
    Wend
    
    
    Pisah_Karakter2 = sWork

End Function
Sub terjemah1()

'Dengan karakter Non AlphaNumeric

If Text2 <> "" Then
Exit Sub
End If

For i = 0 To List3.ListCount - 1
List3.Selected(i) = True
Call Pisah_Karakter1
Next i


For i = 0 To List4.ListCount - 1

List4.ListIndex = i

With FrmCari
If .Option1.Value = True Then
List1.ListIndex = SendMessage(List1.hwnd, LB_FINDSTRINGEXACT, -1, ByVal List4.Text)
ElseIf .Option2.Value = True Then
List1.ListIndex = SendMessage(List1.hwnd, LB_FINDSTRING, -1, ByVal List4.Text)
End If
End With

List5.ListIndex = List4.ListIndex
List3.ListIndex = List4.ListIndex
List2.ListIndex = List1.ListIndex

If List1.Text = "" Then
List6.AddItem List3.List(List3.ListIndex)
Else
List6.AddItem List2.List(List2.ListIndex) & List5.List(List5.ListIndex)
End If

Next i

For i = 0 To List6.ListCount - 1
Text2.SelText = List6.List(i) & " "
Next i

End Sub
Sub terjemah2()
Dim indra As String

'tanpa karakter Non AlphaNumeric

If Text2 <> "" Then
Exit Sub
End If

Dim str, TmpStr As String
str = Pisah_Karakter2(Text1.Text)
TmpStr = str
Call Pecah_Kalimat(TmpStr)

For i = 0 To List3.ListCount - 1

List3.ListIndex = i

With FrmCari
If .Option1.Value = True Then
List1.ListIndex = SendMessage(List1.hwnd, LB_FINDSTRINGEXACT, -1, ByVal List3.Text)
ElseIf .Option2.Value = True Then
List1.ListIndex = SendMessage(List1.hwnd, LB_FINDSTRING, -1, ByVal List3.Text)
End If
End With

List2.ListIndex = List1.ListIndex

If List1.Text = "" Then
List6.AddItem List3.List(List3.ListIndex)
Else
List6.AddItem List2.List(List2.ListIndex)
End If

Next i

For i = 0 To List6.ListCount - 1
Text2.SelText = List6.List(i) & " "
Next i


End Sub
Private Sub Command1_Click()
Dim read As String
read = GetSetting("MultiTrans", "Setting", "Type")

Dummy.SetFocus
If Text1 = "" Then
MsgBox "Mohon masukan kata terlebih dahulu.... :-P", vbInformation, "MultiTranslator"
Text1.SetFocus
Exit Sub
End If

Text2 = ""
List3.Clear
List4.Clear
List5.Clear
List6.Clear

Select Case read

Case "1"
Call terjemah2

Case "2"
Call Pecah_Kalimat(Text1.Text)
Call terjemah1

End Select


End Sub


Private Sub Form_Load()

If App.PrevInstance = True Then
MsgBox "Application allready running...", vbInformation, "Application Info"
End
End If

'st.Panels.Item(1).Text = "indra_g28@yahoo.co.id"
MakeFlat Text1.hwnd
MakeFlat Text2.hwnd

On Error Resume Next

Me.Top = GetSetting("MultiTrans", "Setting", "Top")
Me.Left = GetSetting("MultiTrans", "Setting", "Left")

Dim read1
read1 = GetSetting("MultiTrans", "Setting", "Source")
Text1.Font = FrmFont.List1.Text
Text1.FontSize = FrmFont.Combo1.Text
Text2.Font = FrmFont.List1.Text
Text2.FontSize = FrmFont.Combo1.Text

Select Case read1
Case "Inggris"
mnuIng_Ind_Click
Case "Indonesia"
mnuInd_Ing_Click
End Select

Text1.Font = GetSetting("Multitrans", "Setting", "Font")
Text2.Font = GetSetting("Multitrans", "Setting", "Font")

MakeFlat Command1.hwnd

Me.Height = 6075

End Sub

Private Sub Form_Unload(Cancel As Integer)
mnuClose_Click
End Sub

Private Sub mnuAboutMe_Click()
FrmAbout.Show vbModal
End Sub

Private Sub mnuBackup_Click()
FrmBackup.Show vbModal
End Sub

Private Sub mnuClose_Click()
If FrmBackup.Check1.Value = 1 Then
Call Backup_Database(Me)
End If
SaveSetting "MultiTrans", "Setting", "Top", Me.Top
SaveSetting "MultiTrans", "Setting", "Left", Me.Left
End
End Sub

Private Sub mnuEditor_Click()
FrmEditor.Show vbModal
End Sub

Private Sub mnuFont_Style_Click()
FrmFont.Show vbModal
End Sub

Private Sub mnuHelp_Click()
  On Error Resume Next
  HHShowContents Me.hwnd
End Sub

Private Sub mnuInd_Ing_Click()
mnuIng_Ind.Checked = False
mnuInd_Ing.Checked = True
List1.Clear
List2.Clear
st.Panels.Item(2).Text = "Arah Terjemah : " & Mid(mnuInd_Ing.Caption, 2, Len(mnuInd_Ing.Caption))
Call Load_Data_List("Indonesian", List1, 0)
Call Load_Data_List("Indonesian", List2, 1)
Call Jumlah_Kata
SaveSetting "MultiTrans", "Setting", "Source", "Indonesia"
End Sub
Private Sub mnuIng_Ind_Click()
List1.Clear
List2.Clear
st.Panels.Item(2).Text = "Arah Terjemah : " & Mid(mnuIng_Ind.Caption, 2, Len(mnuIng_Ind.Caption))
Call Load_Data_List("English", List1, 0)
Call Load_Data_List("English", List2, 1)
Call Jumlah_Kata
mnuInd_Ing.Checked = False
mnuIng_Ind.Checked = True
SaveSetting "MultiTrans", "Setting", "Source", "Inggris"
End Sub


Private Sub mnuNew_Click()
Text1 = ""
Text2 = ""
End Sub

Private Sub mnuOpen_Click()
Dim F
On Error GoTo err

cdl.Filter = "Text|*.txt"
cdl.DialogTitle = "Buka Teks File..."
cdl.ShowOpen

F = FreeFile
Open cdl.FileName For Input As #F
Text1.Text = Input$(LOF(F), F)
Close #F

err:
End Sub

Private Sub mnuRestore_Click()
FrmRestore.Show vbModal
End Sub

Private Sub mnuSave_Click()
Dim F
On Error GoTo err

If Text2 = "" Then
    If MsgBox("Output teks masih kosong, lanjutkan??", vbInformation + vbYesNo, "MultiTranslator 1.0") = vbYes Then
    GoTo lanjut
    Else
    Exit Sub
    End If
End If

lanjut:

cdl.Filter = "Text|*.txt"
cdl.DialogTitle = "Simpan ke Teks File...."
cdl.ShowSave
F = FreeFile
Open cdl.FileName For Output As #F
Print #F, Text2.Text
Close #F
err:
End Sub

Private Sub mnuType_Click()
FrmType.Show vbModal
End Sub



Private Sub mnuTypeCari_Click()
FrmCari.Show
End Sub
