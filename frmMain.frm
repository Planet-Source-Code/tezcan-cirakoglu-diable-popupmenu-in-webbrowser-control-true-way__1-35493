VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "Shdocvw.dll"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WB 
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8535
      ExtentX         =   15055
      ExtentY         =   8493
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
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "&Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "C&opy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents IEDoc As HTMLDocument
Attribute IEDoc.VB_VarHelpID = -1

Private Sub Form_Load()
   WB.Navigate App.Path & "\index.htm"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set IEDoc = Nothing
End Sub

Private Function IEDoc_oncontextmenu() As Boolean
   IEDoc_oncontextmenu = False
   PopupMenu mnuEdit
End Function

Private Sub WB_DocumentComplete(ByVal pDisp As Object, URL As Variant)
   Set IEDoc = WB.Document
End Sub

