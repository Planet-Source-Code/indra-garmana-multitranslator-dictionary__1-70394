VERSION 5.00
Begin VB.Form FrmType 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Type Pengolahan"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
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
   ScaleHeight     =   2025
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.OptionButton Option2 
         Caption         =   "Dengan Karakter ( With Char )"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tanpa Karakter  ( Without Char )"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
   End
End
Attribute VB_Name = "FrmType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Option1.Value = True Then
SaveSetting "MultiTrans", "Setting", "Type", "1"
ElseIf Option2.Value = True Then
SaveSetting "MultiTrans", "Setting", "Type", "2"
End If

Unload Me

End Sub

Private Sub Form_Load()
Dim read As String

MakeFlat Command1.hwnd

read = GetSetting("MultiTrans", "Setting", "Type")

Select Case read
Case "1"
Option1.Value = True
Case "2"
Option2.Value = True
End Select

End Sub
