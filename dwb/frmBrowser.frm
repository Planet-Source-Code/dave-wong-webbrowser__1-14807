VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmBrowser 
   BackColor       =   &H80000018&
   Caption         =   "David's Web Browser"
   ClientHeight    =   5625
   ClientLeft      =   3060
   ClientTop       =   3630
   ClientWidth     =   6720
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmBrowser.frx":030A
   ScaleHeight     =   5625
   ScaleWidth      =   6720
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   5370
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11351
            Text            =   "Done"
            TextSave        =   "Done"
         EndProperty
      EndProperty
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   6180
      Top             =   1860
   End
   Begin VB.PictureBox picAddress 
      Align           =   1  'Align Top
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   6720
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   540
      Width           =   6720
      Begin VB.ComboBox cboAddress 
         Height          =   315
         Left            =   45
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   240
         Width           =   3795
      End
      Begin VB.Label lblAddress 
         BackColor       =   &H80000018&
         Caption         =   "&Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Tag             =   "&Address:"
         Top             =   0
         Width           =   3075
      End
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   2670
      Top             =   2685
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":4B450
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":4BB62
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":4C274
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":4C986
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":4D098
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":4D7AA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   953
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Forward"
            Object.ToolTipText     =   "Forward"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   4
         EndProperty
      EndProperty
      MousePointer    =   1
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   5400
      ExtentX         =   9525
      ExtentY         =   6800
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
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPopup 
         Caption         =   "Show Popups"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuDivbar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuNav 
      Caption         =   "N&avigation"
      Begin VB.Menu mnuBack 
         Caption         =   "Ba&ck"
      End
      Begin VB.Menu mnuForward 
         Caption         =   "&Forward"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "&Stop"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refreash"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuBug 
         Caption         =   "&Bug Report"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StartingAddress As String
Dim mbDontNavigateNow As Boolean


Private Sub brwWebBrowser_NewWindow2(ppDisp As Object, Cancel As Boolean)
If mnuPopup.Checked = False Then
    Cancel = True
End If
End Sub

Private Sub brwWebBrowser_OnFullScreen(ByVal FullScreen As Boolean)
MsgBox ("The web site is trying to make this window FullScreen. This browser currently doesn't support this function.")

End Sub

Private Sub brwWebBrowser_OnMenuBar(ByVal MenuBar As Boolean)
MsgBox ("The web site is trying to hide the menu bar. This browser currently doesn't support this function.")

End Sub

Private Sub brwWebBrowser_OnQuit()
a = MsgBox("The web page you are at is trying to close this window.", vbOKCancel + vbInformation, "Close")
If a = vbOK Then
    End
End If
End Sub

Private Sub brwWebBrowser_OnStatusBar(ByVal StatusBar As Boolean)
MsgBox ("The web site is trying to hide the status bar. This browser currently doesn't support this function.")

End Sub

Private Sub brwWebBrowser_OnTheaterMode(ByVal TheaterMode As Boolean)
MsgBox ("The web site is trying to make this window TheaterMode (Whatever that means). This browser currently doesn't support this function.")

End Sub

Private Sub brwWebBrowser_OnToolBar(ByVal ToolBar As Boolean)
MsgBox ("The web site is trying to hide the toolbar. This browser currently doesn't support this function.")

End Sub

Private Sub brwWebBrowser_OnVisible(ByVal Visible As Boolean)
MsgBox ("The web site is trying to hide this window. This browser currently doesn't support this function.")

End Sub

Private Sub brwWebBrowser_StatusTextChange(ByVal Text As String)
StatusBar1.Panels.Item(1).Text = Text
End Sub

Private Sub brwWebBrowser_TitleChange(ByVal Text As String)
Me.Caption = "David's Web Browser - " & Text
End Sub

Private Sub brwWebBrowser_WindowClosing(ByVal IsChildWindow As Boolean, Cancel As Boolean)
a = MsgBox("The web page you are at is trying to close this window.", vbOKCancel + vbInformation, "Close")
If a = vbOK Then
    End
End If
End Sub

Private Sub brwWebBrowser_WindowSetHeight(ByVal Height As Long)
Me.Height = Height
End Sub

Private Sub brwWebBrowser_WindowSetLeft(ByVal Left As Long)
Me.Left = Left
End Sub

Private Sub brwWebBrowser_WindowSetTop(ByVal Top As Long)
Me.Top = Top
End Sub

Private Sub brwWebBrowser_WindowSetWidth(ByVal Width As Long)
Me.Width = Width
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Show
    tbToolBar.Refresh
    Form_Resize


    cboAddress.Move 50, lblAddress.Top + lblAddress.Height + 15


    If Len(StartingAddress) > 0 Then
        cboAddress.Text = StartingAddress
        cboAddress.AddItem cboAddress.Text
        'try to navigate to the starting address
        timTimer.Enabled = True
        brwWebBrowser.Navigate StartingAddress
    End If

brwWebBrowser.GoHome
End Sub

Private Sub brwWebBrowser_DownloadComplete()
    On Error Resume Next
    Me.Caption = brwWebBrowser.LocationName
End Sub


Private Sub brwWebBrowser_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    On Error Resume Next
    Dim i As Integer
    Dim bFound As Boolean
    Me.Caption = brwWebBrowser.LocationName
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = brwWebBrowser.LocationURL Then
            bFound = True
            Exit For
        End If
    Next i
    mbDontNavigateNow = True
    If bFound Then
        cboAddress.RemoveItem i
    End If
    cboAddress.AddItem brwWebBrowser.LocationURL, 0
    cboAddress.ListIndex = 0
    mbDontNavigateNow = False
End Sub


Private Sub cboAddress_Click()
    If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
    brwWebBrowser.Navigate cboAddress.Text
End Sub


Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    If KeyAscii = vbKeyReturn And shift = 2 Then 'This feature don't seem to work
        cboAddress = "http://www." & cboAddress
        aboaddress = cboAddress & ".com"
        cboAddress_Click
    End If
    End If
    
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    cboAddress.Width = Me.ScaleWidth - 50
    brwWebBrowser.Width = Me.ScaleWidth - 25
    brwWebBrowser.Height = Me.ScaleHeight - (picAddress.Top + picAddress.Height + StatusBar1.Height) - 25
End Sub


Private Sub Form_Unload(Cancel As Integer)
a = MsgBox("Are you sure you want to quit?", vbYesNo + vbQuestion, "Quit")
Select Case a
    Case vbYes
        End
End Select
End
End Sub

Private Sub mnuQuit_Click()
End
End Sub


Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuBack_Click()
brwWebBrowser.GoBack
End Sub

Private Sub mnuBug_Click()
a = MsgBox("E-Mail bug reports to syberdave001@hotmail.com", vbInformation, "Bug")
End Sub

Private Sub mnuFileExit_Click()
End
End Sub

Private Sub mnuForward_Click()
brwWebBrowser.GoForward
End Sub

Private Sub mnuRefreash_Click()
brwWebBrowser.Refresh
End Sub

Private Sub mnuPopup_Click()
If mnuPopup.Checked = True Then
    mnuPopup.Checked = False
Else
    mnuPopup.Checked = True
End If
End Sub

Private Sub mnuStop_Click()
timTimer.Enabled = False
brwWebBrowser.Stop
Label1.Caption = brwWebBrowser.LocationName
End Sub

Private Sub timTimer_Timer()
    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser.LocationName
    Else
    End If
End Sub


Private Sub tbToolBar_ButtonClick(ByVal Button As Button)
    On Error Resume Next
      

    timTimer.Enabled = True
      

    Select Case Button.Key
        Case "Back"
            brwWebBrowser.GoBack
        Case "Forward"
            brwWebBrowser.GoForward
        Case "Refresh"
            brwWebBrowser.Refresh
        Case "Home"
            brwWebBrowser.GoHome
        Case "Search"
            brwWebBrowser.GoSearch
        Case "Stop"
            timTimer.Enabled = False
            brwWebBrowser.Stop
            Label1.Caption = brwWebBrowser.LocationName
            
    End Select
    

End Sub

