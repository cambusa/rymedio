VERSION 5.00
Begin VB.UserControl CtlWheel 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   InvisibleAtRuntime=   -1  'True
   MaskColor       =   &H00E0E0E0&
   Picture         =   "wheel.ctx":0000
   ScaleHeight     =   450
   ScaleWidth      =   450
End
Attribute VB_Name = "CtlWheel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_CapWnd As Long
Dim m_Subclassed As Boolean

Event WheelScroll(Shift As Integer, zDelta As Integer, X As Single, Y As Single)

Private Sub UserControl_Resize()

    Size 32 * Screen.TwipsPerPixelX, 32 * Screen.TwipsPerPixelY
     
End Sub

Public Sub DisableWheel()

    If m_CapWnd = 0 Then
        Exit Sub
    End If
     
    If m_Subclassed = False Then
        Exit Sub
    End If

    UnSubclass Me, m_CapWnd
     
    m_Subclassed = False
     
End Sub

Public Sub EnableWheel()

    If m_CapWnd = 0 Then
        Exit Sub
    End If
     
    m_Subclassed = True
     
    Subclass Me, m_CapWnd
     
End Sub

Friend Property Get hwnd() As Long

    hwnd = UserControl.hwnd
     
End Property

Public Property Get hWndCapture() As Long

    hWndCapture = m_CapWnd
     
End Property

Public Property Let hWndCapture(ByVal vNewValue As Long)

    m_CapWnd = vNewValue
     
End Property

Friend Sub WndProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

    Dim wShift As Integer
    Dim wzDelta As Integer
    Dim wX As Single, wY As Single

    wShift = LOWORD(wParam)
    wzDelta = HIWORD(wParam)
    wX = LOWORD(lParam)
    wY = HIWORD(lParam)

    RaiseEvent WheelScroll(wShift, wzDelta, wX, wY)
     
End Sub

