VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon It!"
   ClientHeight    =   1710
   ClientLeft      =   2145
   ClientTop       =   1635
   ClientWidth     =   3960
   Icon            =   "iconit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1710
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cHide 
      Caption         =   "Hide This Window"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.PictureBox cPic 
      Height          =   255
      Left            =   3840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      ToolTipText     =   "Refresh The List"
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox cCombo 
      Height          =   315
      Left            =   105
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Visible Windows"
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton cStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      ToolTipText     =   "Activate"
      Top             =   720
      Width           =   1215
   End
   Begin VB.Menu A7 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu A6 
         Caption         =   "Make Invisible"
      End
      Begin VB.Menu A8 
         Caption         =   "Main Window"
      End
      Begin VB.Menu A4 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Const GW_HWNDNEXT = 2
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim t As NOTIFYICONDATA
Dim tmp As Long
Dim Old_State As Long
Dim lProcOld As Long

Sub loadtasklist()
    Dim currwnd As Long, length As Long
    Dim Buffer As String
    currwnd = GetWindow(Me.hwnd, 0)
    While currwnd <> 0
        length = GetWindowTextLength(currwnd)
        Buffer = Space$(length + 1)
        length = GetWindowText(currwnd, Buffer, length + 1)
        If IsWindow(currwnd) Then
            If IsWindowVisible(currwnd) Then
                If Buffer <> Chr$(0) Then
                    If InStr(UCase$(Buffer), "ICON IT!") = 0 Then
                        cCombo.AddItem Buffer
                        cCombo.ItemData(cCombo.NewIndex) = currwnd
                    End If
                End If
            End If
        End If
        currwnd = GetWindow(currwnd, GW_HWNDNEXT)
        DoEvents
    Wend
End Sub


Private Sub cCombo_DropDown()
LockWindowUpdate cCombo.hwnd
cCombo.Clear
loadtasklist
LockWindowUpdate 0
cCombo.ListIndex = 0
End Sub


Private Sub cStart_Click()
If cStart.Caption = "Start" Then
    cHide.Visible = True
    cPic.Picture = Main.Icon
    tmp = cCombo.ItemData(cCombo.ListIndex)
    App.Title = "Icon It!  -  " & cCombo.List(cCombo.ListIndex)
    t.cbSize = Len(t)
    t.hwnd = cPic.hwnd
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = DetermineBestIcon(tmp)
    t.szTip = cCombo.List(cCombo.ListIndex) & Chr$(0)
    Shell_NotifyIcon NIM_ADD, t
    Me.Visible = False
    cCombo.Enabled = False
    cRefresh.Enabled = False
    cStart.Caption = "Stop"
    Old_State = GetWindowState(tmp)
    A6_Click
Else
    cHide.Visible = False
    cStart.Caption = "Start"
    A6.Caption = "Make Invisible"
    App.Title = "Icon It!"
    Me.Visible = True
    cCombo.Enabled = True
    cRefresh.Enabled = True
    Shell_NotifyIcon NIM_DELETE, t
    SetWindowState tmp, Old_State + 2
    cRefresh_Click
End If
            
End Sub
Private Sub cRefresh_Click()
cCombo.Clear
loadtasklist
cCombo.ListIndex = 0
End Sub

Private Sub cPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button Then
    Me.PopupMenu A7
End If
End Sub

Private Sub A8_Click()
Me.Show
End Sub

Private Sub cHide_Click()
Me.Visible = False
End Sub

Private Sub Form_Load()
Call SubClass(Main.hwnd)
Call loadtasklist
cCombo.ListIndex = 0
Old_State = 2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
t.cbSize = Len(t)
t.hwnd = cPic.hwnd
t.uId = 1&
Shell_NotifyIcon NIM_DELETE, t
SetWindowState tmp, Old_State + 2
SetWindowLong Me.hwnd, GWL_WNDPROC, lProcOld
End Sub

Private Sub A4_Click()
Unload Me
End Sub

Private Sub A6_Click()
If IsWindow(tmp) = False Then End
If IsWindowVisible(tmp) Then
    SetWindowState tmp, 0
    A6.Caption = "Make Visible"
    A4.Visible = False
Else
    SetWindowState tmp, Old_State + 2
    A6.Caption = "Make Invisible"
    A4.Visible = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form_QueryUnload 0, 0
End Sub
