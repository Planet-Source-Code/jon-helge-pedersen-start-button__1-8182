VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cStartButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hwnd As Long, ByVal hWndChild As Long, ByVal lpszClassName As Any, ByVal lpszWindow As Any) As Long
Private Declare Function LoadImage Lib "User32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal iImageType As Long, ByVal cx As Long, ByVal cy As Long, ByVal fFlags As Long) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowRect Lib "User32" (ByVal hwnd As Long, lpRect As Rect) As Long
Private Declare Function GetWindowPlacement Lib "User32" (ByVal hwnd As Long, lpwndpl As WindowPlacement) As Long

Private Type PointApi
  x As Long
  y As Long
End Type

Private Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type WindowPlacement
  Length As Long
  Flags As Long
  ShowCmd As Long
  MinPosition As PointApi
  MaxPosition As PointApi
  NormalPosition As Rect
End Type

Private Const Bm_GetImage = &HF6
Private Const Bm_SetImage = &HF7

Private Const Img_Bitmap = &H0

Private Const Lr_DefaultColor = &H0
Private Const Lr_LoadFromFile = &H10
Private Const Lr_DefaultSize = &H40

Private Const Swp_Position = &H1
Private Const Swp_Size = &H2

Private hTaskbar As Long, hStartButton As Long, hOriginal As Long, hCustom As Long
Private ButtonWidth As Long, ButtonHeight As Long, ButtonLeft As Long, ButtonTop As Long

Private Sub Class_Initialize()
  Dim r As Rect, wp As WindowPlacement: wp.Length = Len(wp)
    hTaskbar = FindWindowEx(0, 0, "Shell_TrayWnd", 0&)
    hStartButton = FindWindowEx(hTaskbar, 0, "BUTTON", 0&)
    hOriginal = SendMessage(hStartButton, Bm_GetImage, Img_Bitmap, ByVal 0&)
    GetWindowRect hStartButton, r
    GetWindowPlacement hStartButton, wp
    ButtonWidth = (r.Right - r.Left) * 15
    ButtonHeight = (r.Bottom - r.Top) * 15
    ButtonLeft = wp.NormalPosition.Left * 15
    ButtonTop = wp.NormalPosition.Top * 15
End Sub

Private Sub Class_Terminate()
  RestoreAll
End Sub

Sub RestoreAll()
  SendMessage hStartButton, Bm_SetImage, Img_Bitmap, ByVal hOriginal
  SetWindowPos hStartButton, 0, 0, 0, ButtonWidth / 15, ButtonHeight / 15, Swp_Size
  SetWindowPos hStartButton, 0, ButtonLeft / 15, ButtonTop / 15, 0, 0, Swp_Position
End Sub

Property Let Bitmap(ByVal Filename As String)
  If Filename <> "" Then
    hCustom = LoadImage(0&, Filename, Img_Bitmap, 0, 0, Lr_DefaultColor Or Lr_LoadFromFile Or lr_defaulsize)
    SendMessage hStartButton, Bm_SetImage, Img_Bitmap, ByVal hCustom
  Else
    SendMessage hStartButton, Bm_SetImage, Img_Bitmap, ByVal 0&
  End If
End Property

Property Get Width() As Long
  Dim r As Rect
    GetWindowRect hStartButton, r
    Width = (r.Right - r.Left) * 15
End Property

Property Let Width(ByVal Width As Long)
  SetWindowPos hStartButton, 0, 0, 0, Width / 15, 330 / 15, Swp_Size
End Property

Property Get Height() As Long
  Dim r As Rect
    GetWindowRect hStartButton, r
    Height = (r.Bottom - r.Top) * 15
End Property

Property Let Height(ByVal Height As Long)
  SetWindowPos hStartButton, 0, 0, 0, 810 / 15, Height / 15, Swp_Size
End Property

Property Get Left() As Long
  Dim wp As WindowPlacement
    wp.Length = Len(wp)
    GetWindowPlacement hStartButton, wp
    Left = wp.NormalPosition.Left * 15
End Property

Property Let Left(ByVal Left As Long)
  SetWindowPos hStartButton, 0, Left / 15, ButtonTop / 15, 0, 0, Swp_Position
End Property

Property Get Top() As Long
  Dim wp As WindowPlacement
    wp.Length = Len(wp)
    GetWindowPlacement hStartButton, wp
    Top = wp.NormalPosition.Top * 15
End Property

Property Let Top(ByVal Top As Long)
  SetWindowPos hStartButton, 0, ButtonLeft / 15, Top / 15, 0, 0, Swp_Position
End Property
