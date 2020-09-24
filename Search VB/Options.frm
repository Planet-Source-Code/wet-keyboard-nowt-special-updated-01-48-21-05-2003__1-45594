VERSION 5.00
Begin VB.Form Options 
   Caption         =   "Options"
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   2070
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   3375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Autologin on program start"
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()

    If Check1.Value = vbChecked Then
    
       Label3.Caption = "Autologin Enabled"
       
    ElseIf Check1.Value = vbUnchecked Then
    
       Label3.Caption = "Autologin Disabled"
       
    End If

End Sub

Private Sub Command1_Click()

   Call SaveSetting("Simple VB Search", "Settings", "AutoSave", Label3.Caption)
   Call SaveSetting("Simple VB Search", "Settings", "E-mail", Text1.Text)
   Call SaveSetting("Simple VB Search", "Settings", "Password", Text2.Text)

End Sub

Private Sub Command2_Click()

   Unload Me

End Sub
