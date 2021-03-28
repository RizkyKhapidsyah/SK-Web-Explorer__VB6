VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Import Export Utility"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   LinkTopic       =   "Form4"
   ScaleHeight     =   4035
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   3000
      Width           =   5775
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   3480
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Finish"
         Height          =   495
         Left            =   4560
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5775
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   120
         Picture         =   "Form4.frx":0000
         ScaleHeight     =   945
         ScaleWidth      =   5505
         TabIndex        =   8
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.ListBox List1 
      Height          =   1035
      ItemData        =   "Form4.frx":0C99
      Left            =   120
      List            =   "Form4.frx":0CA3
      TabIndex        =   2
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      Caption         =   "Description"
      Height          =   1575
      Left            =   2520
      TabIndex        =   0
      Top             =   1440
      Width           =   3255
      Begin VB.Label Label2 
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Chose An Action to perform"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   2175
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim slhelp As ShellUIHelper
Private Sub Command1_Click()
On Error GoTo ecomm
If List1.ListIndex = 0 Then
slhelp.ImportExportFavorites True, "c:\windows\desktop\Favorites.htm"
Unload Me
End If
If List1.ListIndex = 1 Then
slhelp.ImportExportFavorites False, "c:\windows\desktop\Favorites.htm"
Unload Me
End If
ecomm:
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Set slhelp = New ShellUIHelper
End Sub

Private Sub List1_Click()
If List1.ListIndex = 0 Then
Label2.Caption = "Import Favorites From File File must reside in c:\windows\desktop\Favorites.htm  Click Finish To Import it. "
End If
If List1.ListIndex = 1 Then
Label2.Caption = "Export Favorites to another File File will be created in c:\windows\desktop\Favorites.htm  Click Finish To Export it.  "
End If
End Sub

