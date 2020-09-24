VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Login 
   BackColor       =   &H000000FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5370
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H000000FF&
      Caption         =   "Remember"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   840
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   4335
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   5175
      ExtentX         =   9128
      ExtentY         =   2990
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
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   3840
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PROCESSING, PLEASE STAND BY..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   50
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
'save e-mail and password if checked
    If Check1.Value = Checked Then

        Call SaveSetting("Simple VB Search", "Settings", "E-mail", Text1.Text)
        Call SaveSetting("Simple VB Search", "Settings", "Password", Text2.Text)

    ElseIf Check1.Value = Unchecked Then
'Don't remember them
        Text1.Text = ""
        Text2.Text = ""
    End If
End Sub

Private Sub Command1_Click()

If MDIForm1.Login1.Caption = "Login" Then
'Process login
    WebBrowser1.Navigate ("https://www.exhedra.com/ads/authentication/LoginAction.asp?txtEmailAddress=" & Text1.Text & "&txtReturnURL=http%3A%2F%2Fwww%2Eplanet%2Dsource%2Dcode%2Ecom%2Fvb%2Fauthentication%2FLogin%2Easp%3FtxtReturnURL%3D%252Fvb%252Fscripts%252FBrowseCategoryOrSearchResults%252Easp%253FgrpCategories%253D%252D1%2526optSort%253DDateDescending%2526txtMaxNumberOfEntriesPerPage%253D10%2526blnNewestCode%253DTRUE%2526blnResetAllVariables%253DTRUE%2526lngWId%253D1&lngWId=&blnOutsideOfVBSubWeb=False&txtPassword=" & Text2.Text & "&chkRememberPassword=TRUE&cmOk=Ok&strPassKey=")
    Label3.Visible = True
'Stop browser
    Form1.WebBrowser1.Stop
    MDIForm1.Login1.Caption = "Logout"
ElseIf MDIForm1.Login1.Caption = "Logout" Then
    WebBrowser1.Navigate ("http://www.planet-source-code.com/vb/authentication/DeleteCookies.asp?txtReturnURL=" & Form1.WebBrowser1.LocationURL)
    Label3.Visible = True
'Stop browser
    Form1.WebBrowser1.Stop
    MDIForm1.Login1.Caption = "Login"
End If
End Sub

Private Sub Command2_Click()
'unload this form
    Unload Me
End Sub

Private Sub Login()
    'login
    Form1.WebBrowser1.Navigate ("http://www.planet-source-code.com/vb/authentication/Login.asp?txtReturnURL=%2Fvb%2Fscripts%2FBrowseCategoryOrSearchResults%2Easp%3FgrpCategories%3D%2D1%26optSort%3DDateDescending%26txtMaxNumberOfEntriesPerPage%3D10%26blnNewestCode%3DTRUE%26blnResetAllVariables%3DTRUE%26lngWId%3D1")

End Sub

Private Sub Form_Load()
'Stop browser (loads quicker and gets rid of an error that I couldn't get rid of without doing this)
    WebBrowser1.Stop
'load settings
    Text1.Text = GetSetting("Simple VB Search", "Settings", "E-mail", "")
    Text2.Text = GetSetting("Simple VB Search", "Settings", "Password", "")
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
'show message box once complete
    MsgBox "Thankyou for logging into Planet Source Code. Please wait whilst we refresh the page you are currently viewing", vbOKOnly, "Login Successful"
'refresh current page
    Form1.WebBrowser1.Refresh
'unload this form and maximize the browser again
    Unload Me
    Form1.WindowState = vbMaximized
End Sub

