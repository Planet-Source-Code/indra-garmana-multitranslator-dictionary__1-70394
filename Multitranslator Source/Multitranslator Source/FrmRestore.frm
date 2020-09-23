VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmRestore 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Restore Database"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5655
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
   ScaleHeight     =   3600
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Dummy 
      Height          =   195
      Left            =   6840
      TabIndex        =   7
      Top             =   2880
      Width           =   75
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3480
      Top             =   3120
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   6600
      TabIndex        =   4
      Top             =   1440
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   6600
      TabIndex        =   3
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Restore"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data List"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin MSComctlLib.ProgressBar prg 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2520
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.ListBox List1 
         Height          =   2010
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   4800
      Width           =   90
   End
End
Attribute VB_Name = "FrmRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim restore
On Error Resume Next

Dummy.SetFocus

Label1.Caption = "1"

If List1.Text = "" Then
Label1.Caption = "0"
Exit Sub
End If

Kill App.path & "\Data.mdb"

FileCopy File1.path & "\" & List1.Text, App.path & "\Data.mdb"

Timer1.Enabled = True

End Sub

Private Sub Command2_Click()
On Error Resume Next

Dummy.SetFocus

If List1.Text = "" Then
Exit Sub
End If

If MsgBox("Hapus backup data tanggal " & Mid(List1.Text, 1, 10), vbInformation + vbYesNo, "MultiTrans 1.0") = vbYes Then
Kill File1.path & "\" & List1.Text
Timer1.Enabled = True
List1.RemoveItem List1.ListIndex
Else
Exit Sub
End If
End Sub

Private Sub Form_Load()

On Error Resume Next

Dir1.path = App.path & "\Backup Data"
File1.path = Dir1.path

Dim i As Integer
For i = 0 To File1.ListCount - 1
List1.AddItem File1.List(i)
Next i

MakeFlat prg.hwnd
MakeFlat Command1.hwnd
MakeFlat Command2.hwnd
MakeFlat List1.hwnd


End Sub

Private Sub Form_Unload(Cancel As Integer)
If Label1.Caption = "0" Then
Exit Sub
Else
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
End If
End Sub

Private Sub List1_Click()
File1.ListIndex = List1.ListIndex
End Sub

Private Sub Timer1_Timer()
prg.Value = prg.Value + 5
If prg.Value = 100 Then
Timer1.Enabled = False
prg.Value = 0
Unload Me
End If
End Sub
