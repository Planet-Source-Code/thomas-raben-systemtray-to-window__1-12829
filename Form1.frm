VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "SystemTray"
   ClientHeight    =   1320
   ClientLeft      =   165
   ClientTop       =   690
   ClientWidth     =   1560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   88
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   104
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3060
      Top             =   1920
   End
   Begin VB.Menu menu 
      Caption         =   "menu"
      Begin VB.Menu GotoWin 
         Caption         =   "Window Systray"
      End
      Begin VB.Menu split 
         Caption         =   "-"
      End
      Begin VB.Menu EXIT 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EXIT_Click()
    Shell_NotifyIcon NIM_DELETE, nidProgramData
    End
    
End Sub

Private Sub Form_Load()
    Me.Hide
    
    TrayIcon

    
    SetWindowPos Me.hWnd, -1, Me.Left / 15, Me.Top / 15, 100, 100, 0
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If X = 515 Then
        Shell_NotifyIcon NIM_DELETE, nidProgramData
        Me.Show
        Me.menu.Visible = False
        BootUp
        Me.Timer1.Enabled = True
    ElseIf X = 517 Then
        PopupMenu Form1.menu
        
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Timer1.Enabled = False
    
    SetParent SysBox, Parent
    TrayIcon
    Cancel = True
    Me.Hide
    
End Sub

Private Sub GotoWin_Click()
    Shell_NotifyIcon NIM_DELETE, nidProgramData
    Me.Show
    Me.menu.Visible = False
    BootUp
    Me.Timer1.Enabled = True
    
End Sub

Private Sub Timer1_Timer()
    SetWindowPos SysBox, 0, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0
End Sub
