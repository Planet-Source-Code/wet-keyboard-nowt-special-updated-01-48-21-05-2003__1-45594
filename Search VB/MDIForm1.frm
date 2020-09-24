VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{18D91AD0-D0BE-11D1-A6B4-00AA002075DA}#1.0#0"; "FLSHTRAY.OCX"
Begin VB.MDIForm MDIForm1 
   AutoShowChildren=   0   'False
   BackColor       =   &H000000FF&
   Caption         =   "Simple VB Search"
   ClientHeight    =   7320
   ClientLeft      =   165
   ClientTop       =   480
   ClientWidth     =   11880
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin TrayIconPrj.TrayIcon TrayIcon1 
      Left            =   720
      Top             =   2160
      _ExtentX        =   1905
      _ExtentY        =   953
      Icon            =   "MDIForm1.frx":030A
      ToolTipText     =   "Simple VB Search"
      Enabled         =   -1  'True
      TrueClick       =   0   'False
      Visible         =   -1  'True
      FlashSound      =   0
      FlashIcon       =   "MDIForm1.frx":0624
      FlashInterval   =   1000
      FlashEnabled    =   -1  'True
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1429
      BandCount       =   2
      BackColor       =   -2147483639
      _CBWidth        =   11880
      _CBHeight       =   810
      _Version        =   "6.7.8988"
      BandBackColor1  =   -2147483639
      MinHeight1      =   360
      Width1          =   2520
      BandPicture1    =   "MDIForm1.frx":093E
      UseCoolbarColors1=   0   'False
      NewRow1         =   0   'False
      MinHeight2      =   360
      Width2          =   9270
      NewRow2         =   -1  'True
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   190
         Left            =   5520
         TabIndex        =   14
         Top             =   120
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.TextBox txtStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5520
         TabIndex        =   13
         Top             =   80
         Width           =   6255
      End
      Begin VB.TextBox CodeID 
         Height          =   285
         Left            =   8160
         TabIndex        =   12
         Text            =   "Or code ID"
         Top             =   460
         Width           =   2175
      End
      Begin VB.TextBox PSC 
         Height          =   285
         Left            =   4440
         TabIndex        =   11
         Text            =   "Enter search term"
         Top             =   460
         Width           =   3615
      End
      Begin VB.CommandButton Login1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Login"
         Height          =   255
         Left            =   3360
         MaskColor       =   &H00FFC0C0&
         Picture         =   "MDIForm1.frx":BB29
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton Home 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Home"
         Height          =   255
         Left            =   4440
         MaskColor       =   &H00FFC0C0&
         Picture         =   "MDIForm1.frx":E75F
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   80
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton Refresh 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Refresh"
         Height          =   255
         Left            =   3360
         MaskColor       =   &H00FFC0C0&
         Picture         =   "MDIForm1.frx":11395
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   80
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton Stop 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Stop"
         Height          =   255
         Left            =   2280
         MaskColor       =   &H00FFC0C0&
         Picture         =   "MDIForm1.frx":13FCB
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   80
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton Forward 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Forward"
         Height          =   255
         Left            =   1200
         MaskColor       =   &H00FFC0C0&
         Picture         =   "MDIForm1.frx":16C01
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   80
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton Back 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Back"
         Height          =   255
         Left            =   120
         MaskColor       =   &H00FFC0C0&
         Picture         =   "MDIForm1.frx":19837
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   80
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Upload"
         Height          =   255
         Left            =   2280
         MaskColor       =   &H00FFC0C0&
         Picture         =   "MDIForm1.frx":1C46D
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Forum"
         Height          =   255
         Left            =   1200
         MaskColor       =   &H00FFC0C0&
         Picture         =   "MDIForm1.frx":1F0A3
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "New"
         Height          =   255
         Left            =   120
         MaskColor       =   &H00FFC0C0&
         Picture         =   "MDIForm1.frx":21CD9
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF80&
         Caption         =   "Find It!"
         Height          =   255
         Left            =   10440
         MaskColor       =   &H00FFC0C0&
         Picture         =   "MDIForm1.frx":2490F
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   465
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
   End
   Begin VB.Menu FileFile 
      Caption         =   "File"
      Begin VB.Menu FileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu TTools 
      Caption         =   "Tools"
      Visible         =   0   'False
      Begin VB.Menu TForum 
         Caption         =   "Forum"
      End
      Begin VB.Menu TLogin 
         Caption         =   "Login"
      End
      Begin VB.Menu TNewCode 
         Caption         =   "New Code"
      End
      Begin VB.Menu TUpload 
         Caption         =   "Upload"
      End
      Begin VB.Menu TSep 
         Caption         =   "-"
      End
      Begin VB.Menu TAbout 
         Caption         =   "About"
      End
      Begin VB.Menu THelp 
         Caption         =   "Help"
      End
      Begin VB.Menu TSe 
         Caption         =   "-"
      End
      Begin VB.Menu TExit 
         Caption         =   "Exit Program"
      End
   End
   Begin VB.Menu OOptions 
      Caption         =   "Options"
      Begin VB.Menu OSettings 
         Caption         =   "Settings"
      End
   End
   Begin VB.Menu HelpHlp 
      Caption         =   "Help"
      Begin VB.Menu HelpAbout 
         Caption         =   "About"
      End
      Begin VB.Menu HelpRepBug 
         Caption         =   "Report Bug"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CodeID_KeyPress(KeyAscii As Integer)
'Allow user to press return key to begin search
    If KeyAscii = vbKeyReturn Then
        Command1_Click
    Else
        
    End If
End Sub

Private Sub PSC_Click()
    PSC.Text = ""
    'Prevent the user from screwing up the search
    CodeID.Text = "Or code ID"
End Sub

Private Sub CodeID_Click()
    CodeID.Text = ""
    'Prevent the user from screwing up the search
    PSC.Text = "Enter search term"
End Sub

Private Sub Command1_Click()
'Start search
    If PSC.Text = "Enter search term" And CodeID.Text = "Or code ID" Then
       MsgBox "Please enter a search term or a code ID", vbExclamation + vbOKOnly, "Search term error"
    Else
        If CodeID.Text = "Or code ID" Then
            Form1.Show
            Form1.WebBrowser1.Navigate ("http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&lngWId=1&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&txtCriteria=" & (MDIForm1.PSC.Text))
            MDIForm1.Caption = MDIForm1.Caption & "Searching for: " & PSC.Text & " Please stand by..."
        ElseIf PSC.Text = "Enter search term" Then
            Form1.Show
            Form1.WebBrowser1.Navigate ("http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=" & (MDIForm1.CodeID.Text) & "&lngWId=1")
            MDIForm1.Caption = MDIForm1.Caption & "Searching for: " & CodeID.Text & " Please stand by..."
        End If
    End If
    
End Sub

Private Sub Command2_Click()
'Navigate to newest code
    Form1.Show
    Form1.WebBrowser1.Navigate ("http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?grpCategories=-1&optSort=DateDescending&txtMaxNumberOfEntriesPerPage=10&blnNewestCode=TRUE&blnResetAllVariables=TRUE&lngWId=1")
End Sub

Private Sub Command3_Click()
'Navigate to forums
    Form1.Show
    Form1.WebBrowser1.Navigate ("http://www.planet-source-code.com/vb/discussion/default.asp?lngWId=1")
End Sub

Private Sub Command4_Click()
'Navigate to author options
    Form1.Show
    Form1.WebBrowser1.Navigate ("http://www.planet-source-code.com/vb/authors/determine_author_type.asp?lngWId=1")
End Sub

Private Sub FileExit_Click()
'Unload forms and end program
    Dim frm As Form
    For Each frm In Forms
    Unload frm
    Set frm = Nothing
    Next
    Unload MDIForm1
    Set MDIForm1 = Nothing

End Sub

Private Sub HelpAbout_Click()
'Show About
    About.Show
End Sub

Private Sub OSettings_Click()

   Options.Show

End Sub

Private Sub PSC_KeyPress(KeyAscii As Integer)
'Allow user to press return key to begin search
    If KeyAscii = vbKeyReturn Then
        Command1_Click
    Else
        
    End If
End Sub

Private Sub HelpRepBug_Click()
'Browse to reportbug page
    Form1.WebBrowser1.Navigate ("http://www.your-website-com/reportbug.html")
End Sub

Private Sub Home_Click()
'Default homepage for Simple VB Search
    Form1.WebBrowser1.Navigate ("http://www.planet-source-code.com/vb/CodeOfTheDay/signup.asp?lngWId=1")
End Sub

Private Sub Login1_Click()

  If Login1.Caption = "Login" Then
    'Show login dialog
     Login.Show
  Else
    'Change to logout and provide logout URL (currently goes to Newest Code when logging out)
    
    Form1.WebBrowser1.Navigate ("http://www.planet-source-code.com/vb/authentication/DeleteCookies.asp?txtReturnURL=/vb/scripts/BrowseCategoryOrSearchResults.asp?grpCategories=-1&optSort=DateDescending&txtMaxNumberOfEntriesPerPage=10&blnNewestCode=TRUE&blnResetAllVariables=TRUE&lngWId=1")
    Login1.Caption = "Login"
  End If

End Sub

Private Sub MDIForm_Load()
'Show browser and navigate to blank page (loads quicker)
    Form1.Show
    Form1.WebBrowser1.Navigate ("about:blank")

   Options.Label3.Caption = GetSetting("Simple VB Search", "Settings", "AutoSave", "")

   If Options.Label3.Caption = "Autologin Enabled" Then
        Options.Check1.Value = vbChecked
   Else
        Options.Check1.Value = vbUnchecked
        Options.Label3.Caption = "Autologin Disabled"
   End If
        Options.Text1.Text = GetSetting("Simple VB Search", "Settings", "E-mail", "")
        Options.Text2.Text = GetSetting("Simple VB Search", "Settings", "Password", "")
   
   MDIForm1.Caption = "Simple VB Search - [Planet Source Code]"
   
End Sub

Private Sub Refresh_Click()
'Refresh page
    Form1.WebBrowser1.Refresh
End Sub

Private Sub Stop_Click()
'Stop loading
    Form1.WebBrowser1.Stop
End Sub

Private Sub Forward_Click()
'Forward
    Form1.WebBrowser1.GoForward
End Sub

Private Sub Back_Click()
'Go back
    Form1.WebBrowser1.GoBack
End Sub
Private Sub TAbout_Click()
'Systray menu item
    About.Show
End Sub

Private Sub TExit_Click()
'Systray menu item
    FileExit_Click
End Sub

Private Sub TForum_Click()
'Systray menu item
    Command3_Click
End Sub

Private Sub THelp_Click()
'Systray menu item
    MsgBox "Help not currently available", vbInformation, "Simple VB Search"
End Sub

Private Sub TLogin_Click()
'Systray menu item
    Login1_Click
End Sub

Private Sub TNewCode_Click()
'Systray menu item
    Command2_Click
End Sub

Private Sub TrayIcon1_LeftButtonClick()
'Show systray menu when left clicked
    PopupMenu MDIForm1.TTools
End Sub

Private Sub TrayIcon1_RightButtonClick()
'Show systray menu when right clicked
    PopupMenu MDIForm1.TTools
End Sub

Private Sub TUpload_Click()
'Systray menu item
    Command4_Click
End Sub
