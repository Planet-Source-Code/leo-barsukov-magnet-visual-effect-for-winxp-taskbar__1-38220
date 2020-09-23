Attribute VB_Name = "Start"
Option Explicit

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWndChild As Long) As Long
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type WINDOWPLACEMENT
        Length As Long
        flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
End Type

Private Const GW_NEXT = 2
Private Const GW_CHILD = 5
Private Const BM_SETIMAGE = &HF7
Property Get hwnd() As Long
Dim CHwnd1 As Long, CHwnd2 As Long
Dim CLS_NM As String * 7

CHwnd1 = GetDesktopWindow
CHwnd1 = GetWindow(CHwnd1, GW_CHILD)
Do While CHwnd1 <> 0
    CHwnd2 = GetWindow(CHwnd1, GW_CHILD)
    Do While CHwnd2 <> 0
        GetClassName CHwnd2, CLS_NM, 7
        If Left(CLS_NM, 6) = "Button" Then
            hwnd = CHwnd2
            Exit Property
        End If
        CHwnd2 = GetWindow(CHwnd2, GW_NEXT)
    Loop
    CHwnd1 = GetWindow(CHwnd1, GW_NEXT)
Loop
End Property
Property Let hPic(ByVal hPicture As Long)
    PostMessage hwnd, BM_SETIMAGE, 0, hPicture
End Property
Property Let Width(ByVal sWidth As Long)
    SetWindowPos hwnd, 0, 0, 0, sWidth / 15, Height / 15, 2
End Property
Property Get Width() As Long
Dim tmpRECT As RECT

GetWindowRect hwnd, tmpRECT
Width = (tmpRECT.Right - tmpRECT.Left) * 15
End Property
Property Let Height(ByVal sHeight As Long)
    SetWindowPos hwnd, 0, 0, 0, Width / 15, sHeight / 15, 2
End Property
Property Get Height() As Long
Dim tmpRECT As RECT

GetWindowRect hwnd, tmpRECT
Height = (tmpRECT.Bottom - tmpRECT.Top) * 15
End Property
Property Let Left(ByVal lX As Long)
    SetWindowPos hwnd, 0, lX / 15, Top / 15, 0, 0, 1
End Property
Property Get Left() As Long
Dim tmpPLC As WINDOWPLACEMENT

tmpPLC.Length = Len(tmpPLC)
GetWindowPlacement Start.hwnd, tmpPLC

Left = tmpPLC.rcNormalPosition.Left * 15
End Property
Property Let Top(ByVal lY As Long)
    SetWindowPos hwnd, 0, Left / 15, lY / 15, 0, 0, 1
End Property
Property Get Top() As Long
Dim tmpPLC As WINDOWPLACEMENT

tmpPLC.Length = Len(tmpPLC)
GetWindowPlacement Start.hwnd, tmpPLC

Top = tmpPLC.rcNormalPosition.Top * 15
End Property
Property Let Parent(ByVal hParent As Long)
    SetParent hwnd, hParent
End Property
Property Get Parent() As Long
    Parent = GetParent(hwnd)
End Property


