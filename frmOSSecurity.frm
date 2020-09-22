VERSION 5.00
Begin VB.Form frmOSSecurity 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Disable Ctrl-Alt-Del"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   Icon            =   "frmOSSecurity.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2145
      TabIndex        =   11
      Top             =   4140
      Width           =   1005
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   780
      TabIndex        =   10
      Top             =   4140
      Width           =   1005
   End
   Begin VB.Frame frameWinNT 
      Caption         =   "Win NT"
      Height          =   2700
      Left            =   330
      TabIndex        =   1
      Top             =   1140
      Visible         =   0   'False
      Width           =   3345
      Begin VB.CheckBox chkShutDown 
         Caption         =   "Disable Shutdown"
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   2310
         Width           =   2400
      End
      Begin VB.CheckBox chkTaskMgr 
         Caption         =   "Disable Task Manager"
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   1905
         Width           =   2400
      End
      Begin VB.CheckBox chkLogOff 
         Caption         =   "Disable LogOff"
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   1515
         Width           =   2400
      End
      Begin VB.CheckBox chkLockWkSt 
         Caption         =   "Disable Workstation Locking"
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   1110
         Width           =   2400
      End
      Begin VB.CheckBox chkPassword 
         Caption         =   "Disable Password Change"
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   2400
      End
      Begin VB.Label lblWinNT 
         AutoSize        =   -1  'True
         Caption         =   "Windows NT/2K/XP "
         Height          =   195
         Left            =   210
         TabIndex        =   2
         Top             =   390
         Width           =   1530
      End
   End
   Begin VB.Frame frameWin9x 
      Caption         =   "Win 9x"
      Height          =   1005
      Left            =   315
      TabIndex        =   0
      Top             =   105
      Visible         =   0   'False
      Width           =   3360
      Begin VB.CheckBox chkWin9xDisable 
         Caption         =   "Disable Ctrl-Alt-Del"
         Height          =   240
         Left            =   450
         TabIndex        =   8
         Top             =   495
         Width           =   2355
      End
      Begin VB.Label lblWin9X 
         AutoSize        =   -1  'True
         Caption         =   "Windows 9x "
         Height          =   195
         Left            =   420
         TabIndex        =   9
         Top             =   225
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmOSSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'api decl for Win9x
Private Declare Function SystemParametersInfo Lib _
"user32" Alias "SystemParametersInfoA" (ByVal uAction _
As Long, ByVal uParam As Long, ByVal lpvParam As Any, _
ByVal fuWinIni As Long) As Long

Private Const SPI_SCREENSAVERRUNNING = 97


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
Dim objRegistry          As clsRegistry
Dim objOSversion    As clsOSVersion
Dim lngRV           As Long

On Error GoTo cmdSave_Click_Error
    Set objOSversion = New clsOSVersion
    If objOSversion.IsPlatformWin9X = True Then
        'if in Win9x, disable using api
        If chkWin9xDisable.Value = vbChecked Then
            lngRV = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, "1", 0&)
        Else
            lngRV = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, "1", 0&)
        End If
    Else
        'if in  WinNT, disable use registry
        Set objRegistry = New clsRegistry
        'taskmanager
        If chkTaskMgr.Value = vbChecked Then
            objRegistry.SetValue eHKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", CLng(1), eREG_DWORD
        Else
            objRegistry.SetValue eHKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", CLng(0), eREG_DWORD
        End If
        'Workstation locking
        If chkLockWkSt.Value = vbChecked Then
            objRegistry.SetValue eHKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableLockWorkstation", CLng(1), eREG_DWORD
        Else
            objRegistry.SetValue eHKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableLockWorkstation", CLng(0), eREG_DWORD
        End If
        'Password change
        If chkPassword.Value = vbChecked Then
            objRegistry.SetValue eHKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableChangePassword", CLng(1), eREG_DWORD
        Else
            objRegistry.SetValue eHKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableChangePassword", CLng(0), eREG_DWORD
        End If
        'Logoff
        If chkLogOff.Value = vbChecked Then
            objRegistry.SetValue eHKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogoff", CLng(1), eREG_DWORD
        Else
            objRegistry.SetValue eHKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogoff", CLng(0), eREG_DWORD
        End If
        'shutdown
        If chkShutDown.Value = vbChecked Then
            objRegistry.SetValue eHKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose", CLng(1), eREG_DWORD
        Else
            objRegistry.SetValue eHKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose", CLng(0), eREG_DWORD
        End If
            
        Set objRegistry = Nothing
    End If
    
    Set objOSversion = Nothing

    On Error GoTo 0
    Exit Sub
cmdSave_Click_Error:
    Select Case Err.Number
        Case 0
        Case Else
            '#If Debugging Then
                MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line# " & Erl & " in procedure cmdSave_Click of Form frmOSSecurity"
            '#End If
            'Call WriteErrorFile(Err.Number & "||" & Err.Description & "||" & Erl & "||" & "cmdSave_Click||frmOSSecurity" & "||" & g_CALL_STACK & "||" & Date & "||" & Time)
    End Select
End Sub

Private Sub Form_Load()

Dim objOSversion As clsOSVersion

On Error GoTo Form_Load_Error
    'set the frame position
    frameWinNT.Top = frameWin9x.Top
    frameWinNT.Left = frameWin9x.Left
    
    'show the options according to the OS present
    Set objOSversion = New clsOSVersion
    If objOSversion.IsPlatformWin9X = True Then
        'if platform is Windows 9x
        frameWin9x.Visible = True
        'adjust command button positions
        cmdSave.Top = frameWin9x.Top + frameWin9x.Height + 100
        cmdCancel.Top = frameWin9x.Top + frameWin9x.Height + 100
        'adjust the form height accordingly
        frmOSSecurity.Height = frameWin9x.Top + 2 * frameWin9x.Height
    Else
        'if platform is Win NT series
        frameWinNT.Visible = True
        'adjust command button positions
        cmdSave.Top = frameWinNT.Top + frameWinNT.Height + 100
        cmdCancel.Top = frameWinNT.Top + frameWinNT.Height + 100
        'adjust the form height accordingly
        frmOSSecurity.Height = frameWinNT.Top + frameWinNT.Height + (0.4 * frameWinNT.Height)
    End If
    Set objOSversion = Nothing

    On Error GoTo 0
    Exit Sub
Form_Load_Error:
    Select Case Err.Number
        Case 0
        Case Else
            '#If Debugging Then
                MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line# " & Erl & " in procedure Form_Load of Form frmOSSecurity"
            '#End If
            'Call WriteErrorFile(Err.Number & "||" & Err.Description & "||" & Erl & "||" & "Form_Load||frmOSSecurity" & "||" & g_CALL_STACK & "||" & Date & "||" & Time)
    End Select
    
End Sub
