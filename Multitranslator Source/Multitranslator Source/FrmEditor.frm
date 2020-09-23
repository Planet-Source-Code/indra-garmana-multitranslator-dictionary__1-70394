VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmEditor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Database Editor"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton dummy 
      Height          =   255
      Left            =   8040
      TabIndex        =   18
      Top             =   2880
      Width           =   135
   End
   Begin VB.PictureBox Picture3 
      Height          =   375
      Left            =   4080
      ScaleHeight     =   315
      ScaleWidth      =   1995
      TabIndex        =   15
      Top             =   6120
      Width           =   2055
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "dari :"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   45
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "-"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   555
         TabIndex        =   16
         Top             =   45
         Width           =   60
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   13
      Top             =   6120
      Width           =   855
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Data ke"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   45
         Width           =   555
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3195
      ScaleWidth      =   5955
      TabIndex        =   9
      Top             =   2760
      Width           =   6015
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3200
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   5636
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   17
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Database Editor"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.OptionButton Option2 
         Caption         =   "Indonesia Ke Inggris"
         Height          =   255
         Left            =   3600
         TabIndex        =   12
         Top             =   1440
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Inggris Ke Indonesia"
         Height          =   255
         Left            =   3600
         TabIndex        =   11
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Cari"
         Height          =   375
         Left            =   4560
         TabIndex        =   8
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Hapus"
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Edit"
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Tambah"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   465
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1080
      Top             =   6120
      Width           =   2895
      _ExtentX        =   5106
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
      ForeColor       =   0
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
      Caption         =   ""
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
End
Attribute VB_Name = "FrmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo err
Text1.Text = Adodc1.Recordset.Fields(0)
Text2.Text = Adodc1.Recordset.Fields(1)
Adodc1.Caption = Adodc1.Recordset.AbsolutePosition
Label5.Caption = Adodc1.Recordset.RecordCount
err:
End Sub

Private Sub Command1_Click()
Dummy.SetFocus

If Command4.Caption = "Refresh" Then
Command4_Click
End If

Select Case Command1.Caption

Case "Tambah"
Call Grid_Size
Adodc1.Recordset.AddNew
Text1 = ""
Text2 = ""
Command2.Caption = "Simpan"
Command1.Caption = "Batal"
Text1.Locked = False
Text2.Locked = False
Text1.SetFocus

Case "Batal"

Adodc1.Recordset.Cancel
Adodc1.Refresh
On Error Resume Next
Text1 = Adodc1.Recordset.Fields(0)
Text2 = Adodc1.Recordset.Fields(1)
Command2.Caption = "Edit"
Command1.Caption = "Tambah"
Text1.Locked = True
Text2.Locked = True

Call Grid_Size

End Select

End Sub

Private Sub Command2_Click()
Dummy.SetFocus

Select Case Command2.Caption

Case "Edit"

Command1.Caption = "Batal"
Command2.Caption = "Simpan"

Text1.Locked = False
Text2.Locked = False
Text1.SetFocus


Case "Simpan"
Call Grid_Size
If Text1 = "" Or Text2 = "" Then
MsgBox "Mohon isi terlebih dahulu pada kolom yang kosong...", vbInformation, "MultiTranslator 1.0"
Exit Sub
End If

If Trim(Adodc1.Recordset.Fields(0)) = Trim(Text1.Text) Then
    If MsgBox("Data dengan kata " & Text1 & " telah ada..." & vbCrLf & "Lanjutkan simpan data??..", vbInformation + vbYesNo, "MultiTranslator 1.0") = vbYes Then
    GoTo lanjut
    Else
    Exit Sub
    End If
End If


lanjut:
Command1.Caption = "Tambah"
Command2.Caption = "Edit"
Command4.Caption = "Cari"
Adodc1.Recordset.Fields(0) = Trim(Text1.Text)
Adodc1.Recordset.Fields(1) = Trim(Text2.Text)
Adodc1.Recordset.Update
Text1.Locked = True
Text2.Locked = True

End Select

Adodc1.Refresh
Adodc1.Refresh

Call Grid_Size

End Sub

Private Sub Command3_Click()
Dummy.SetFocus

Call Grid_Size

If Adodc1.Recordset.AbsolutePosition <= -1 Then Exit Sub

If MsgBox("Hapus data " & "( " & Text1.Text & " = " & Text2.Text & " )", vbInformation + vbYesNo, "MultiTranslator 1.0") = vbYes Then
On Error Resume Next
Adodc1.Recordset.Delete
Adodc1.Recordset.Update

If Command4.Caption = "Refresh" Then
Command4_Click
End If

Else
Exit Sub
End If

End Sub

Private Sub Command4_Click()
Dummy.SetFocus
Dim str1 As String

Call Grid_Size

Select Case Command4.Caption

Case "Cari"

If Text1 = "" And Text2 = "" Then Command1_Click

str1 = InputBox("Masukan kata yang dicari dalam bahasa " & Label1.Caption & "......", "Cari kata dalam kamus...")
If str1 = "" Then Exit Sub
If Label1.Caption = "Inggris" Then
Adodc1.Recordset.Filter = "English ='" & str1 & "'"
ElseIf Label1.Caption = "Indonesia" Then
Adodc1.Recordset.Filter = "Indonesia ='" & str1 & "'"
End If
Command4.Caption = "Refresh"

Case "Refresh"
Adodc1.Refresh
Command4.Caption = "Cari"
End Select

End Sub

Private Sub Form_Load()
Dim read As String
read = GetSetting("MultiTrans", "Setting", "Editor")

If read = "" Then
Option1.Value = True
Option1_Click
End If

Call Grid_Size

Select Case read
Case "1"
Option1_Click
Option1.Value = True
Case "2"
Option2_Click
Option2.Value = True
End Select

MakeFlat Text1.hwnd
MakeFlat Text2.hwnd
MakeFlat Picture1.hwnd
MakeFlat Picture2.hwnd
MakeFlat Picture3.hwnd
MakeFlat Command1.hwnd
MakeFlat Command2.hwnd
MakeFlat Command3.hwnd
MakeFlat Command4.hwnd

End Sub

Private Sub Form_Unload(Cancel As Integer)

If Text1.Text = "" Or Text2.Text = "" Then
Adodc1.Refresh
Exit Sub
End If

If Main.mnuInd_Ing.Checked = True Then
Main.List1.Clear
Main.List2.Clear
Call Load_Data_List("Indonesian", Main.List1, 0)
Call Load_Data_List("Indonesian", Main.List2, 1)
Call Jumlah_Kata
ElseIf Main.mnuIng_Ind.Checked = True Then
Main.List1.Clear
Main.List2.Clear
Call Load_Data_List("English", Main.List1, 0)
Call Load_Data_List("English", Main.List2, 1)
Call Jumlah_Kata
End If

Adodc1.Recordset.Close
End Sub

Private Sub Option1_Click()
Label1.Caption = "Inggris"
Label2.Caption = "Indonesia"

Call Inggris
Call Grid_Size
Adodc1.Caption = Adodc1.Recordset.AbsolutePosition
Label5.Caption = Adodc1.Recordset.RecordCount

SaveSetting "MultiTrans", "Setting", "Editor", "1"

End Sub

Private Sub Option2_Click()
Label1.Caption = "Indonesia"
Label2.Caption = "Inggris"

Call Indonesia
Call Grid_Size
Adodc1.Caption = Adodc1.Recordset.AbsolutePosition
Label5.Caption = Adodc1.Recordset.RecordCount

SaveSetting "MultiTrans", "Setting", "Editor", "2"

End Sub

Sub Inggris()
Adodc1.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.path & "\Data.mdb")
Adodc1.RecordSource = "Select * From English"
Set DataGrid1.DataSource = Adodc1
On Error Resume Next
Text1 = ""
Text2 = ""
Adodc1.Refresh
End Sub

Sub Indonesia()
Adodc1.ConnectionString = ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.path & "\Data.mdb")
Adodc1.RecordSource = "Select * From Indonesian"
Set DataGrid1.DataSource = Adodc1
On Error Resume Next
Text1 = ""
Text2 = ""
Adodc1.Refresh
End Sub

Sub Grid_Size()
DataGrid1.Columns(0).Width = 2700
DataGrid1.Columns(1).Width = 2700
End Sub
