VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   Caption         =   "Window Resize"
   ClientHeight    =   1845
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5820
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imgListIcons 
      Left            =   5160
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser picPanel 
      Height          =   1575
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   1815
      ExtentX         =   3201
      ExtentY         =   2778
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "res://C:\WINNT\system32\shdoclc.dll/dnserror.htm#http:///"
   End
   Begin MSComctlLib.TreeView treWindows 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2778
      _Version        =   393217
      Style           =   7
      ImageList       =   "imgListIcons"
      Appearance      =   1
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   3120
      MousePointer    =   9  'Size W E
      ScaleHeight     =   1215
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   120
      Width           =   40
   End
   Begin VB.Menu mnuChange 
      Caption         =   "Control"
      Begin VB.Menu mnuResizeAll 
         Caption         =   "Recursive Resize"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuResize 
         Caption         =   "Resize"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuHidden 
         Caption         =   "Show Hidden Windows"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Resize other apps controls at runtime

Private Sub Form_Load()

    Me.Width = Screen.TwipsPerPixelX * 800
    Me.Height = Screen.TwipsPerPixelY * 600
    
    Me.picPanel.Navigate "about:Loading.."
    Me.picPanel.Navigate App.Path & "\resize.html"
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.picSplit.Move Me.picSplit.Left, 0, Me.picSplit.Width, Me.ScaleHeight
    Me.treWindows.Move 0, 0, Me.picSplit.Left, Me.ScaleHeight
    Me.picPanel.Move Me.picSplit.Left + Me.picSplit.Width, 0, Me.ScaleWidth - (Me.picSplit.Left + Me.picSplit.Width), Me.ScaleHeight
    Me.Refresh
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuHidden_Click()
    Me.mnuHidden.Checked = Not Me.mnuHidden.Checked
    main
End Sub

Private Sub mnuRefresh_Click()
    Me.picPanel.Refresh
    main
End Sub

Private Sub mnuResize_Click()
    If Me.mnuResize.Checked Then
        fixWindow hSelected
    Else
        resizeWindow hSelected
    End If
    Me.mnuResize.Checked = Not Me.mnuResize.Checked
    'selectWindow (hSelected)
End Sub

Private Sub mnuResizeAll_Click()
    If Me.mnuResizeAll.Checked Then
        fixAll hSelected
    Else
        resizeAll hSelected
    End If
    Me.mnuResizeAll.Checked = Not Me.mnuResizeAll.Checked
    'selectWindow (hSelected)
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Me.picSplit.Move Me.picSplit.Left + X, 0, 40, Me.ScaleHeight
        Form_Resize
    End If
End Sub

Private Sub treWindows_NodeClick(ByVal Node As MSComctlLib.Node)
    Debug.Print "Clicked " & Node.Key
    selectWindow hWndFromKey(Node.Key)
End Sub
