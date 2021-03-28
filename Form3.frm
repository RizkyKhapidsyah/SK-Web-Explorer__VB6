VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H80000009&
   Caption         =   "Import Export Utility"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5760
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   4095
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   3120
      Width           =   5775
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   4080
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Next"
         Height          =   495
         Left            =   2880
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form4.Show
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
