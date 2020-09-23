VERSION 5.00
Begin VB.Form FrmFont 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Font Style"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5505
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
   ScaleHeight     =   5115
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Font.frx":0000
      Left            =   960
      List            =   "Font.frx":0028
      TabIndex        =   3
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Font Style"
      Height          =   4500
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.ListBox List1 
         Height          =   4155
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Font Size :"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   765
   End
End
Attribute VB_Name = "FrmFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
SaveSetting "MultiTrans", "Setting", "Font", List1.Text
SaveSetting "MultiTrans", "Setting", "Font Size", Combo1.Text
Main.Text1.Font = List1.Text
Main.Text1.FontSize = Combo1.Text
Main.Text2.Font = List1.Text
Main.Text2.FontSize = Combo1.Text
Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer

MakeFlat List1.hwnd
MakeFlat Command1.hwnd

For i = 0 To Screen.FontCount - 1
List1.AddItem Screen.Fonts(i)
Next

List1.Text = GetSetting("MultiTrans", "Setting", "Font")
Combo1.Text = GetSetting("MultiTrans", "Setting", "Font Size")

If Combo1.Text = "" Then
Combo1.ListIndex = 0
ElseIf List1.Text = "" Then
List1.Selected(0) = True
End If

End Sub
