VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Popup message - UPDATED! (1.0.2)"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9960
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command13 
      Caption         =   "Popup All (Local, External, PC Name)"
      Height          =   735
      Left            =   6960
      TabIndex        =   13
      Top             =   3840
      Width           =   2895
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Popup Computer Name"
      Height          =   735
      Left            =   6960
      TabIndex        =   12
      Top             =   3000
      Width           =   2895
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Popup Local && External IP Address"
      Height          =   735
      Left            =   6960
      TabIndex        =   11
      Top             =   2160
      Width           =   2895
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   10080
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Popup Local IP Address"
      Height          =   735
      Left            =   6960
      TabIndex        =   10
      Top             =   1320
      Width           =   2895
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   10080
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Popup External IP Address"
      Height          =   735
      Left            =   6960
      TabIndex        =   9
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Exit Program"
      Height          =   735
      Left            =   6960
      TabIndex        =   6
      Top             =   6360
      Width           =   2895
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Remove System Tray Icon"
      Height          =   735
      Left            =   6960
      TabIndex        =   8
      Top             =   5520
      Width           =   2895
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Add To System Tray"
      Height          =   735
      Left            =   6960
      TabIndex        =   7
      Top             =   4680
      Width           =   2895
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Standard - Popup Message"
      Height          =   855
      Left            =   4200
      TabIndex        =   5
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Info - Popup Message"
      Height          =   855
      Left            =   4200
      TabIndex        =   4
      Top             =   4680
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Warning - Popup Message"
      Height          =   855
      Left            =   4200
      TabIndex        =   3
      Top             =   3360
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "None - Popup Message"
      Height          =   855
      Left            =   4200
      TabIndex        =   2
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Error - Popup Message"
      Height          =   855
      Left            =   4200
      TabIndex        =   0
      Top             =   720
      Width           =   2415
   End
   Begin VB.Line Line8 
      X1              =   120
      X2              =   120
      Y1              =   480
      Y2              =   7080
   End
   Begin VB.Line Line7 
      X1              =   6840
      X2              =   6840
      Y1              =   480
      Y2              =   7080
   End
   Begin VB.Image Image1 
      Height          =   1035
      Left            =   240
      Picture         =   "Form1.frx":0442
      Top             =   720
      Width           =   3735
   End
   Begin VB.Image Image5 
      Height          =   1035
      Left            =   240
      Picture         =   "Form1.frx":CE20
      Top             =   6000
      Width           =   3765
   End
   Begin VB.Image Image4 
      Height          =   1050
      Left            =   240
      Picture         =   "Form1.frx":19A26
      Top             =   4680
      Width           =   3765
   End
   Begin VB.Image Image3 
      Height          =   1035
      Left            =   240
      Picture         =   "Form1.frx":26920
      Top             =   3360
      Width           =   3765
   End
   Begin VB.Image Image2 
      Height          =   960
      Left            =   240
      Picture         =   "Form1.frx":33526
      Top             =   2040
      Width           =   3375
   End
   Begin VB.Line Line6 
      X1              =   6840
      X2              =   120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line5 
      X1              =   6840
      X2              =   120
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line4 
      X1              =   6840
      X2              =   120
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line3 
      X1              =   6840
      X2              =   120
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line2 
      X1              =   6840
      X2              =   120
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      X1              =   6840
      X2              =   120
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UPDATED! - Don't forget to Vote!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   9975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'::::::::::::::::::::::::::::::::::::::::::::::::::::
' Popup Balloon UPDATED! (1.0.2)
'
'   Now Supports:
' > Different Icons
' > External IP Address Popups
' > Local IP Address Popups
' > Computer Name Popups
'::::::::::::::::::::::::::::::::::::::::::::::::::::

 Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long


Private Sub Command1_Click()
'Starts System Tray
   With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Sample" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
         End With
   Shell_NotifyIcon NIM_ADD, m_IconData
   
'Popup the balloon - Error Icon
Popup_error "This is a sample popup message!", "This is the Tittle of pop message"

End Sub

Private Sub Command10_Click()
'Starts System Tray
   With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Sample" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
         End With
   Shell_NotifyIcon NIM_ADD, m_IconData

Popup "Your Local IP Address is: " & Winsock1.LocalIP, "Local IP..."
End Sub

Private Sub Command11_Click()

'Starts System Tray
   With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Sample" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
         End With
   Shell_NotifyIcon NIM_ADD, m_IconData
   
Popup "Your Local IP Address is: " & Winsock1.LocalIP & vbNewLine & "Your External IP Address is: " & GetExternalIP(Inet), "IP Address..."
End Sub

Private Sub Command12_Click()

'Starts System Tray
   With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Sample" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
         End With
   Shell_NotifyIcon NIM_ADD, m_IconData

'Computer Name
Dim PCName As String
Dim P As Long
P = NameOfPC(PCName)

Popup "Your Computer Name is: " & PCName, "Computer Name..."

End Sub



Private Sub Command13_Click()

'Starts System Tray
   With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Sample" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
         End With
   Shell_NotifyIcon NIM_ADD, m_IconData


'Computer Name
Dim PCName As String
Dim P As Long
P = NameOfPC(PCName)
 
Popup "Your Local IP Address is: " & Winsock1.LocalIP & vbNewLine & vbNewLine & "Your External IP Address is: " & GetExternalIP(Inet) & vbNewLine & vbNewLine & "Your Computer Name is: " & PCName, "IP Address & Computer Name..."

End Sub

Private Sub Command2_Click()
'Starts System Tray
   With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Sample" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
         End With
   Shell_NotifyIcon NIM_ADD, m_IconData

'Popup the balloon - No Icon
Popup_none "This is a sample popup message!", "This is the Tittle of pop message"

End Sub

Private Sub Command3_Click()
'Starts System Tray
   With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Sample" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
         End With
   Shell_NotifyIcon NIM_ADD, m_IconData

'Popup the balloon - Error Warning
Popup_warning "This is a sample popup message!", "This is the Tittle of pop message"

End Sub

Private Sub Command4_Click()
'Starts System Tray
   With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Sample" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
         End With
   Shell_NotifyIcon NIM_ADD, m_IconData

'Popup the balloon - Info Icon
Popup_info "This is a sample popup message!", "This is the Tittle of pop message"

End Sub

Private Sub Command5_Click()
'Starts System Tray
   With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Sample" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
         End With
   Shell_NotifyIcon NIM_ADD, m_IconData

'Popup the balloon - Error Standard
Popup "This is a sample popup message!", "This is the Tittle of pop message"

End Sub

Private Sub Command6_Click()

'Removes Icon
Shell_NotifyIcon NIM_DELETE, m_IconData

'Exits Program
End

End Sub

Private Sub Command7_Click()

'Starts System Tray
   With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Sample" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
         End With
   Shell_NotifyIcon NIM_ADD, m_IconData
   
End Sub

Private Sub Command8_Click()
Shell_NotifyIcon NIM_DELETE, m_IconData
End Sub

Private Sub Command9_Click()

'Starts System Tray
   With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Sample" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
         End With
   Shell_NotifyIcon NIM_ADD, m_IconData

Popup "Your External IP Address is: " & GetExternalIP(Inet), "External IP..."

End Sub

Private Sub Form_Unload(Cancel As Integer)

'get rid of the icon in the system tray
 Shell_NotifyIcon NIM_DELETE, m_IconData
End Sub

'Gets External IP Address
Public Function GetExternalIP(IIN As Inet)
  Dim lngA As Long
  
  Dim strA As String
  Dim strB() As String
  
  'Grabs HTML and Phases Data
  strA = IIN.OpenURL("www.ipchicken.com")
  strB = Split(strA, Chr(10))
  
  GetExternalIP = Trim(Replace(Replace(strB(34), " ", ""), "<br>", ""))
  
End Function

'Gets Computer Name
Public Function NameOfPC(MachineName As String) As Long
    Dim NameSize As Long
    Dim X As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    X = GetComputerName(MachineName, NameSize)
End Function
