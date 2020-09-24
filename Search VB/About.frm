VERSION 5.00
Begin VB.Form About 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6675
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form2"
   ScaleHeight     =   3840
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   $"About.frx":0000
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   6495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright ©2003 Wet Keyboard"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   4815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Licence: Freeware"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Website: "
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Author : Wet Keyboard"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Version: 2.0"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Simple VB Search"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6480
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " About"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   0
      Left            =   240
      Picture         =   "About.frx":0126
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label2_Click()
    Unload Me
End Sub
