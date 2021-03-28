VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDialUp 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AutoConnect to VPN by BugMaster"
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmDialUp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Text            =   "Ras List"
      ToolTipText     =   "Ras Listesi"
      Top             =   0
      Width           =   4095
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   3600
      Top             =   960
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "Connection Status"
      ToolTipText     =   "Connection Status"
      Top             =   600
      Width           =   4095
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Connection Status Check Time Option"
      Top             =   360
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      _Version        =   393216
      Min             =   2
      Max             =   60
      SelStart        =   2
      Value           =   2
   End
End
Attribute VB_Name = "frmDialUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ftDUNConnection$

Private Type RASENTRYNAME95
    dwSize As Long
    szEntryname(256) As Byte
End Type

Private Declare Function RasEnumEntriesA Lib "RasApi32.DLL" _
    (ByVal reserved As String, ByVal lpszPhonebook As String, _
    lprasentryname As Any, lpcb As Long, lpcEntries As Long) _
    As Long
    
Private Declare Function LoadLibrary Lib "kernel32" _
  Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" _
  (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" _
  (ByVal hLibModule As Long) As Long

Public Sub DUN_Services(DUN_Array() As String)
'Pass in Empty array for DUN_Array
    Dim s As Long, ln As Long, conname As String, i As Long
    Dim r(255) As RASENTRYNAME95
    r(0).dwSize = 264
    s = 256 * r(0).dwSize
    
    Call RasEnumEntriesA(vbNullString, vbNullString, r(0), s, ln)
    ln = ln - 1
    ReDim DUN_Array(ln)
    For i = 0 To ln
        conname = StrConv(r(i).szEntryname(), vbUnicode)
        DUN_Array(i) = Left$(conname, InStr(conname, _
          vbNullChar) - 1)
    Next i
End Sub

Private Sub Form_Load()

   frmDialUp.Show
   frmDialUp.Width = 4185
   frmDialUp.Height = 1290
   
   With nid
      .cbSize = Len(nid)
      .hWnd = Me.hWnd
      .uId = vbNull
      .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
      .uCallBackMessage = WM_MOUSEMOVE
      .hIcon = Me.Icon
      .szTip = "Vpn Connection Manager "
   End With
   Shell_NotifyIcon NIM_ADD, nid
   Slider1.Value = (Timer1.Interval / 1000)
   
   If Not APIFunctionPresent("RasEnumEntriesA", "RasApi32.DLL") Then
            MsgBox "You need to have Dial-Up Network installed on your machined to run this program", vbCritical, "Cannot Continue"
            End
    End If
   
   'ras listesini yuklesi baslar
   Dim sArray() As String
   Dim iCtr As Integer
   DUN_Services sArray
   
   For iCtr = 0 To UBound(sArray)
     Combo1.AddItem sArray(iCtr)
   Next
   'ras listesini yuklemesi biter
   
   ftDUNConnection$ = GetSetting("AutoDUN", "Options", "Connection", "")
   Combo1.Text = ftDUNConnection$
   If Len(ftDUNConnection$) = 0 Then
    'ftDUNConnection$ = InputBox("Connection name", "Enter conn name", ftDUNConnection$)
      ftDUNConnection$ = Combo1.Text
   End If
End Sub


Private Sub Combo1_Click()
   ftDUNConnection$ = Combo1.Text
   If Not Len(ftDUNConnection$) = 0 Then
      SaveSetting "AutoDUN", "Options", "Connection", ftDUNConnection$
   End If

   Call Timer1_Timer
End Sub

Private Sub Combo1_Change()
   'Call Combo1_Click
End Sub


Private Sub Slider1_Change()
  Timer1.Enabled = False
  Timer1.Interval = (Slider1.Value * 1000)
  Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
   If RASCount() Then
      Text1.Text = "VPN already Connected"
      With nid
         .szTip = "Vpn Connection Manager Status: Connected"
      End With
      Shell_NotifyIcon NIM_MODIFY, nid
   Else
      retval = InternetDial(Me.hWnd, ftDUNConnection$, 2, 0, 0&)
      If RASCount() Then
         Text1.Text = "VPN Connected"
         With nid
            .szTip = "Vpn Connection Manager Status: Connected"
         End With
         Shell_NotifyIcon NIM_MODIFY, nid
      Else
         Text1.Text = "VPN Not connected"
         With nid
            .szTip = "Vpn Connection Manager Status: Not Connected"
         End With
         Shell_NotifyIcon NIM_MODIFY, nid
      End If
   End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim Result As Long
   Dim msg As Long
   If Me.ScaleMode = vbPixels Then
      msg = X
   Else
      msg = X / Screen.TwipsPerPixelX
   End If
      
   Select Case msg
      Case WM_LBUTTONUP
         frmDialUp.Show
         Me.WindowState = vbNormal
         Result = SetForegroundWindow(Me.hWnd)
      Case WM_RBUTTONUP
         frmDialUp.Show
         Me.WindowState = vbNormal
         Result = SetForegroundWindow(Me.hWnd)
   End Select
End Sub
Public Function APIFunctionPresent(ByVal FunctionName _
   As String, ByVal DllName As String) As Boolean

'USAGE:
'Dim bAvail as boolean
'bAvail = APIFunctionPresent("GetDiskFreeSpaceExA", "kernel32")

    Dim lHandle As Long
    Dim lAddr  As Long

    lHandle = LoadLibrary(DllName)
    If lHandle > 0 Then
        lAddr = GetProcAddress(lHandle, FunctionName)
        FreeLibrary lHandle
    End If
    
    APIFunctionPresent = (lAddr <> 0)

End Function
Private Sub Form_Resize()
   If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Shell_NotifyIcon NIM_DELETE, nid
   Unload Me
   End
End Sub
