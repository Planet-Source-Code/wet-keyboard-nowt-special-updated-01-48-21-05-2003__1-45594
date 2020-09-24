VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   11715
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5745
   ScaleWidth      =   11715
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      ExtentX         =   20770
      ExtentY         =   10186
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
      Location        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    WebBrowser1.Width = Form1.ScaleWidth
    WebBrowser1.Height = Form1.ScaleHeight
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    MDIForm1.txtStatus.Text = "Document found"
End Sub
Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
      
      Dim CurrentValue As Integer
      Dim Step As Integer
      
      If ProgressMax > 0 And Progress > 0 Then
            Step = ProgressMax \ 100
            
            CurrentValue = (Progress \ Step) Mod 101
            
'            Form2.ProgressBar1.Value = CurrentValue
            MDIForm1.ProgressBar1.Value = CurrentValue
      Else
'            Form2.ProgressBar1.Value = 0
            MDIForm1.ProgressBar1.Value = 0
      End If
      
      Debug.Print "ProgressMax : " & ProgressMax
      Debug.Print "Progress : " & Progress
      Debug.Print "***********************"
      
End Sub

