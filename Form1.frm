VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form1 
   Caption         =   "Design Web Explorer"
   ClientHeight    =   6255
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   16
      Top             =   5970
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4455
      Left            =   0
      TabIndex        =   15
      Top             =   1440
      Width           =   7815
      ExtentX         =   13785
      ExtentY         =   7858
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton Command14 
         Caption         =   "Edit"
         Height          =   615
         Left            =   6480
         Picture         =   "Form1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Print"
         Height          =   615
         Left            =   5760
         Picture         =   "Form1.frx":052E
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Favorite"
         Height          =   615
         Left            =   5040
         Picture         =   "Form1.frx":0C30
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton Command12 
         Caption         =   "History"
         Height          =   615
         Left            =   4320
         Picture         =   "Form1.frx":12CA
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Search"
         Height          =   615
         Left            =   3600
         Picture         =   "Form1.frx":18E4
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Home"
         Height          =   615
         Left            =   2880
         Picture         =   "Form1.frx":1A92
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Refresh"
         Height          =   615
         Left            =   2160
         Picture         =   "Form1.frx":1C40
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Stop"
         Height          =   615
         Left            =   1440
         Picture         =   "Form1.frx":1DEE
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Forward"
         Height          =   615
         Left            =   720
         Picture         =   "Form1.frx":1F9C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Back"
         Height          =   615
         Left            =   0
         Picture         =   "Form1.frx":214A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   7575
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   "Address:"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Height          =   375
         Left            =   5160
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu new 
         Caption         =   "&New"
         Begin VB.Menu window 
            Caption         =   "&Window"
            Shortcut        =   ^N
         End
      End
      Begin VB.Menu open 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu saveas 
         Caption         =   "Save &As"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu psetup 
         Caption         =   "Page Set&up"
      End
      Begin VB.Menu print 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu importexport 
         Caption         =   "&Import and Export"
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu close 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu cut 
         Caption         =   "C&ut"
         Shortcut        =   ^X
      End
      Begin VB.Menu copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu line4 
         Caption         =   "-"
      End
      Begin VB.Menu selectall 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu find 
         Caption         =   "&Find(on This Page)..."
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu view 
      Caption         =   "&View"
      Begin VB.Menu toolbar 
         Caption         =   "&Toolbars"
         Begin VB.Menu standbutton 
            Caption         =   "&Standerd Buttons"
         End
         Begin VB.Menu addbar 
            Caption         =   "&Address Bar"
         End
      End
      Begin VB.Menu line5 
         Caption         =   "-"
      End
      Begin VB.Menu goto 
         Caption         =   "&Go to"
         Begin VB.Menu back 
            Caption         =   "&Back"
         End
         Begin VB.Menu forward 
            Caption         =   "&Forward"
         End
      End
      Begin VB.Menu stop 
         Caption         =   "&Stop"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu refresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu source 
         Caption         =   "&Source"
      End
      Begin VB.Menu font 
         Caption         =   "&Font(Text Size)..."
         Shortcut        =   ^T
      End
      Begin VB.Menu fullscreen 
         Caption         =   "&Full Screen"
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu favourite 
      Caption         =   "&Favorite"
      Begin VB.Menu addtofavorities 
         Caption         =   "&Add To Favorites..."
      End
      Begin VB.Menu organizefavorites 
         Caption         =   "&Organize Favorites... "
      End
      Begin VB.Menu line6 
         Caption         =   "-"
      End
      Begin VB.Menu channels 
         Caption         =   "&Channels"
      End
   End
   Begin VB.Menu tools 
      Caption         =   "&Tools"
      Begin VB.Menu Webupdate 
         Caption         =   "Webbrowser &Update"
      End
      Begin VB.Menu rlinks 
         Caption         =   "Show &Related Links"
      End
      Begin VB.Menu line7 
         Caption         =   "-"
      End
      Begin VB.Menu ioptions 
         Caption         =   "Internet &Options"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu cindex 
         Caption         =   "&Content and Index"
      End
      Begin VB.Menu tip 
         Caption         =   "Tip of the &Day"
      End
      Begin VB.Menu netscape 
         Caption         =   "For &Netscape Users"
      End
      Begin VB.Menu tour 
         Caption         =   "&Tour"
      End
      Begin VB.Menu olinesupport 
         Caption         =   "Online &Support"
      End
      Begin VB.Menu feedback 
         Caption         =   "Send Feedbac&k"
      End
      Begin VB.Menu line8 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "&About Design Web Explorer"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim fa As Integer
Dim cap As String
Dim url1 As String
Dim slhelp As ShellUIHelper

Private Sub about_Click()
frmAbout.Show
frmAbout.Icon = Form1.Icon
End Sub

Private Sub addtofavorities_Click()
On Error GoTo eaddfav
url1 = Text1.Text
slhelp.AddFavorite url1, "Website Title"
eaddfav:
End Sub

Private Sub back_Click()
On Error GoTo noback
WebBrowser1.GoBack
Exit Sub
noback:
End Sub

Private Sub cindex_Click()
CommonDialog1.HelpFile = "Iexplore.HLP"
CommonDialog1.HelpCommand = cdlHelpContents
CommonDialog1.ShowHelp
End Sub
Private Sub close_Click()
End
End Sub

Private Sub Command10_Click()
WebBrowser1.GoSearch

End Sub

Private Sub Command11_Click()
Dim nFolder As SpecialShellFolderIDs
  Dim pidl As Long
  Dim cbpidl As Integer
  Dim abpidl() As Byte
  Dim avpidl As Variant
  Dim sPath As Long
  nFolder = CSIDL_FAVORITES
  If SHGetSpecialFolderLocation(hWnd, nFolder, pidl) = NOERROR Then
    If pidl Then
       cbpidl = GetPIDLSize(pidl)
      If cbpidl Then
        ReDim abpidl(cbpidl - 1)
        MoveMemory abpidl(0), ByVal pidl, cbpidl
         avpidl = abpidl
        WebBrowser1.Navigate2 avpidl
        WebBrowser1.Visible = True
      End If
      Call CoTaskMemFree(pidl)
    End If
  End If
  
End Sub

Private Sub Command12_Click()
Dim nFolder As SpecialShellFolderIDs
  Dim pidl As Long
  Dim cbpidl As Integer
  Dim abpidl() As Byte
  Dim avpidl As Variant
  Dim sPath As Long
  nFolder = CSIDL_HISTORY
  If SHGetSpecialFolderLocation(hWnd, nFolder, pidl) = NOERROR Then
    If pidl Then
       cbpidl = GetPIDLSize(pidl)
      If cbpidl Then
        ReDim abpidl(cbpidl - 1)
        MoveMemory abpidl(0), ByVal pidl, cbpidl
         avpidl = abpidl
        WebBrowser1.Navigate2 avpidl
        WebBrowser1.Visible = True
      End If
      Call CoTaskMemFree(pidl)
    End If
  End If
End Sub

Private Sub Command13_Click()
 On Error Resume Next
    eQuery = WebBrowser1.QueryStatusWB(OLECMDID_PRINT)
    If Err.Number = 0 Then
        If eQuery And OLECMDF_ENABLED Then
            WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER, "", ""
        Else
            MsgBox "The print command is currently disabled."
        End If
    End If
End Sub

Private Sub Command14_Click()
On Error GoTo editerror
'Form2.Text1.Text = WebBrowser1.Document.documentElement.innerHTML
Form2.Text1.Text = WebBrowser1.Document.documentElement.outerHTML
Form2.Caption = cap & "-Design Web Explorer"
Form2.Icon = Form1.Icon
Form2.Show
editerror:
End Sub

Private Sub Command3_Click()
WebBrowser1.Navigate Text1.Text
End Sub

Private Sub Command5_Click()
On Error GoTo noback
WebBrowser1.GoBack
Exit Sub
noback:
End Sub

Private Sub Command6_Click()
On Error GoTo noforward
WebBrowser1.GoForward
Exit Sub
noforward:
End Sub

Private Sub Command7_Click()
On Error GoTo nstop
WebBrowser1.stop
nstop:
End Sub

Private Sub Command8_Click()
On Error GoTo erefresh
WebBrowser1.Refresh2
erefresh:
End Sub

Private Sub Command9_Click()
On Error GoTo ehome
WebBrowser1.GoHome
ehome:
End Sub

Private Sub copy_Click()
WebBrowser1.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub cut_Click()
WebBrowser1.ExecWB OLECMDID_CUT, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub find_Click()
'WebBrowser1.ExecWB OLECMDID_SHOWFIND, OLECMDEXECOPT_PROMPTUSER
'WebBrowser1.ExecWB OLECMDID_FIND, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub font_Click()
On Error GoTo efont
CommonDialog1.ShowFont
efont:
End Sub
Private Sub Form_Load()
WebBrowser1.StatusBar = True
CommonDialog1.DefaultExt = "HTM|*.htm"
CommonDialog1.Filter = "HTM|*.htm|HTML|*.html|Bitmap|*.bmp|gif images|*.gif|jpeg images|*.jpg"
 'ShowProgressInStatusBar True
     Set slhelp = New ShellUIHelper
End Sub

Private Sub Form_Resize()
If (Form1.Width - (8 * Screen.TwipsPerPixelX)) > 0 Then
        WebBrowser1.Width = Form1.Width - (8 * Screen.TwipsPerPixelX)
    End If
    If (Form1.Height - WebBrowser1.Top - StatusBar1.Height - (28 * Screen.TwipsPerPixelY)) > 0 Then
        WebBrowser1.Height = Form1.Height - WebBrowser1.Top - StatusBar1.Height - (28 * Screen.TwipsPerPixelY)
    End If
    If Form1.Width / 4 * 3 > 0 Then
        StatusBar1.Panels(1).Width = Form1.Width / 4 * 3
    End If
    'Toolbar1.Left = 0
    Frame2.Left = 0
    If (Form1.Width - Toolbar1.Left - (8 * Screen.TwipsPerPixelX)) > 0 Then
        'Toolbar1.Width = Form1.Width - Toolbar1.Left - (8 * Screen.TwipsPerPixelX)
   If (Form1.Width - Frame2.Left - (8 * Screen.TwipsPerPixelX)) > 0 Then
       Frame2.Width = Form1.Width - Frame2.Left - (8 * Screen.TwipsPerPixelX)
    End If
    End If
End Sub

Private Sub forward_Click()
On Error GoTo noforward
WebBrowser1.GoForward
Exit Sub
noforward:
End Sub

Private Sub fullscreen_Click()
Form1.Width = Screen.Width
Form1.Height = Screen.Height

End Sub

Private Sub importexport_Click()
Form3.Show
End Sub

Private Sub ioptions_Click()
Form5.Show
End Sub

Private Sub netscape_Click()
On Error GoTo enhelp
CommonDialog1.HelpFile = "navigator.HLP"
CommonDialog1.HelpCommand = cdlHelpContents
CommonDialog1.ShowHelp
enhelp:
CommonDialog1.HelpFile = "Iexplore.HLP"
CommonDialog1.HelpCommand = cdlHelpContents
CommonDialog1.ShowHelp

End Sub

Private Sub open_Click()
CommonDialog1.ShowOpen
WebBrowser1.Navigate (CommonDialog1.FileName)
End Sub

Private Sub open1_Click()

End Sub

Private Sub organizefavorites_Click()
slhelp.ShowBrowserUI "OrganizeFavorites", 0
End Sub

Private Sub paste_Click()
WebBrowser1.ExecWB OLECMDID_PASTESPECIAL, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub print_Click()
  
    On Error Resume Next
    eQuery = WebBrowser1.QueryStatusWB(OLECMDID_PRINT)
    If Err.Number = 0 Then
        If eQuery And OLECMDF_ENABLED Then
            WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER, "", ""
        Else
            MsgBox "The print command is currently disabled."
        End If
    End If
End Sub


Private Sub psetup_Click()
    Dim eQuery As OLECMDF
    
    On Error Resume Next
    eQuery = WebBrowser1.QueryStatusWB(OLECMDID_PRINT)
    If Err.Number = 0 Then
        If eQuery And OLECMDF_ENABLED Then
            WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER, "", ""
        Else
            MsgBox "The Page Setup For Print Command is currently disabled."
        End If
    End If
End Sub

Private Sub refresh_Click()
WebBrowser1.refresh
End Sub


Private Sub Scriptlet1_onscriptletevent(ByVal name As String, ByVal eventData As Variant)

End Sub

Private Sub rlinks_Click()
Dim s As String
Dim d As String
On Error GoTo Webupdate
s = "http://geocities.com/ali0001/rlinks.html"
d = "http://geocities.com/ali0001/rlinks.htm"
WebBrowser1.Navigate (s)
Webupdate:
WebBrowser1.Navigate (d)
End Sub

Private Sub save_Click()

WebBrowser1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub saveas_Click()
WebBrowser1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub selectall_Click()
WebBrowser1.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub source_Click()
On Error GoTo esource
Form2.Text1.Text = WebBrowser1.Document.documentElement.innerHTML
Form2.Caption = cap & "-Design Web Explorer"
Form2.Icon = Form1.Icon
Form2.Show
esource:
End Sub
Private Sub standbutton_Click()
If fa = 0 Then
Toolbar1.Visible = False
fa = 1
Exit Sub
End If
If fa = 1 Then
Toolbar1.Visible = True
fa = 0
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       WebBrowser1.Navigate2 Text1.Text
    End If
End Sub

Private Sub tip_Click()
formtip.Show

End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
Text1.Text = URL
End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
If Index = 0 Then
    If Progress >= 0 Then
    Label1.Caption = "Download Progress" & Progress & "/" & ProgressMax
         ' ProgressBar1.Value = ProgressMax
           If Progress = ProgressMax Then
            Label1.Caption = "DownLoad Complete"
        End If
    Else
    Label1.Caption = "Page Downloaded"
    End If
End If
End Sub
Private Sub WebBrowser1_TitleChange(ByVal Text As String)
cap = Text
Form1.Caption = cap & "-Design Web Explorer"
End Sub

Private Sub Webupdate_Click()
Dim s As String
Dim d As String
On Error GoTo Webupdate
s = "http://geocities.com/ali0001/webupdate.html"
d = "http://geocities.com/ali0001/webupdate.htm"
WebBrowser1.Navigate (s)
Webupdate:
WebBrowser1.Navigate (d)
End Sub

Private Sub window_Click()
   Dim frmWB As Form1
   Set frmWB = New Form1
   frmWB.WebBrowser1.RegisterAsBrowser = True
   Set ppDisp = frmWB.WebBrowser1.object
   frmWB.Visible = True
End Sub
Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
   Dim frmWB As Form1
 Set frmWB = New Form1
 frmWB.WebBrowser1.RegisterAsBrowser = True
 Set ppDisp = frmWB.WebBrowser1.object
 frmWB.Visible = True
End Sub
Private Sub webBrowser1_StatusTextChange(ByVal Text As String)
    If Text <> "Done" Then
        StatusBar1.Panels(1).Text = Text
    End If
End Sub
'Private Sub ShowProgressInStatusBar(ByVal bShowProgressBar As Boolean)

  '  Dim tRC As RECT
    
   ' If bShowProgressBar Then

    '    SendMessageAny StatusBar1.hWnd, SB_GETRECT, 1, tRC
     '   With tRC
    '        .Top = (.Top * Screen.TwipsPerPixelY)
     '       .Left = (.Left * Screen.TwipsPerPixelX)
      '      .Bottom = (.Bottom * Screen.TwipsPerPixelY) - .Top
     '       .Right = (.Right * Screen.TwipsPerPixelX) - .Left
      '  End With
      '  With ProgressBar1
      '      SetParent .hWnd, StatusBar1.hWnd
      '      .Move tRC.Left, tRC.Top, tRC.Right, tRC.Bottom
      '      .Visible = True
      '    .Value = 0
      '  End With
      '    Else
      '  SetParent ProgressBar1.hWnd, Me.hWnd
    '  '  ProgressBar1.Visible = False
    'End If
    
'End Sub

