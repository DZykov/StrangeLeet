VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Option"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3765
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Apply"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Turn on reminder"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Œ "
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
options_remind = Check1.Value
saveConf
loadConf
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
options_remind = Check1.Value
saveConf
loadConf
End Sub

Private Sub Form_Load()
Check1.Value = CInt(options_remind) ^ 2
End Sub
