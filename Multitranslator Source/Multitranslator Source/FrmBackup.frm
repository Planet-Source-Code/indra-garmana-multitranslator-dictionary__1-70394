VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmBackup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Backup Database"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5295
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
   ScaleHeight     =   2370
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Dummy 
      Height          =   195
      Left            =   6480
      TabIndex        =   5
      Top             =   840
      Width           =   195
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Backup"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5280
      Top             =   1920
   End
   Begin VB.Frame Frame1 
      Caption         =   "Backup"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.CheckBox Check1 
         Caption         =   "Backup Database On Exit"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   2295
      End
      Begin MSComctlLib.ProgressBar prg 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   4815
      End
   End
End
Attribute VB_Name = "FrmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
dummy.SetFocus
Label1.Caption = "Process..."
Call Backup_Database(Me)
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
MakeFlat prg.hWnd
MakeFlat Command1.hWnd

On Error GoTo lanjut
Check1.Value = GetSetting("MultiTrans", "Setting", "Auto Backup")
Label1.Caption = "Backup Terakhir Tanggal : " & GetSetting("MultiTrans", "Setting", "Backup")
MkDir App.path & "\Backup Data"
lanjut:
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting "MultiTrans", "Setting", "Auto Backup", Check1.Value
End Sub

Private Sub Timer1_Timer()
prg.Value = prg.Value + 5
If prg.Value = 100 Then
Timer1.Enabled = False
prg.Value = 0
Label1.Caption = "Backup Database Selesai..."
Unload Me
End If
End Sub

