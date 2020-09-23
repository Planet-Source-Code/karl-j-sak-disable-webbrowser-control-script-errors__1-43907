VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   4320
      Top             =   3960
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   9135
      ExtentX         =   16113
      ExtentY         =   9763
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



Private Sub Form_Load()
Timer1.Enabled = True

'WebBrowser1.Navigate "http://www.avril-lavigne.ws/"
WebBrowser1.Navigate "http://www.abcnews.com"


End Sub

Private Sub Timer1_Timer()
WebBrowser1.Silent = True
Timer1.Enabled = False
End Sub


Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
If CtrlKey <> True Then
    Cancel = True
End If

End Sub

Function CtrlKey() As Boolean
    CtrlKey = (GetAsyncKeyState(vbKeyControl) And &H8000)
End Function



