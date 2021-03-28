VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   LinkTopic       =   "Form5"
   ScaleHeight     =   3855
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Languages..."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   1095
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim slhelp As ShellUIHelper
Private Sub Command1_Click()
slhelp.ShowBrowserUI "LanguageDialog", 0
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
   Set slhelp = New ShellUIHelper
End Sub
