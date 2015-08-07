Attribute VB_Name = "BasWheel"
Option Explicit

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC = (-4)
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Public Const WM_MOUSEWHEEL = &H20A
Public Const WM_MOUSELAST = &H20A
Public Const WHEEL_DELTA = 120 '/* Value for rolling one detent */

Public Function HIWORD(LongIn As Long) As Integer

    ' Mask off low word then do integer divide to
    ' shift right by 16.

    HIWORD = (LongIn And &HFFFF0000) \ &H10000
   
End Function

Public Function LOWORD(LongIn As Long) As Integer

    ' Low word retrieved by masking off high word.
    ' If low word is too large, twiddle sign bit.

    If (LongIn And &HFFFF&) > &H7FFF Then
        LOWORD = (LongIn And &HFFFF&) - &H10000
    Else
        LOWORD = LongIn And &HFFFF&
    End If
    
End Function

Public Function CtlWheelProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim OldProc As Long
    Dim CtlWnd As Long
    Dim CtlPtr As Long
    Dim IntObj As Object 'Intermediate object in between
    Dim MWObject As CtlWheel 'pointer and mousewheel control
    
    CtlWnd = GetProp(hwnd, "WheelWnd")
    CtlPtr = GetProp(CtlWnd, "WheelPtr")
    OldProc = GetProp(CtlWnd, "OldWheelProc")
    
    If wMsg = WM_MOUSEWHEEL Then
     
        CopyMemory IntObj, CtlPtr, 4
        
        Set MWObject = IntObj
        MWObject.WndProc hwnd, wMsg, wParam, lParam
        Set MWObject = Nothing
        
        CopyMemory IntObj, 0&, 4
        
        Exit Function
          
    End If

    CtlWheelProc = CallWindowProc(OldProc, hwnd, wMsg, wParam, lParam)
     
End Function

Public Sub Subclass(MWCtl As CtlWheel, ParentWnd As Long)

    If GetProp(MWCtl.hwnd, "OldWheelProc") <> 0 Then
        Exit Sub
    End If

    'Save the old window proc of the control's parent
    SetProp MWCtl.hwnd, "OldWheelProc", GetWindowLong(ParentWnd, GWL_WNDPROC)
    
    'Object pointer to the control
    SetProp MWCtl.hwnd, "WheelPtr", ObjPtr(MWCtl)
    
    'Save control's hWnd in its parent data
    SetProp ParentWnd, "WheelWnd", MWCtl.hwnd

    'Subclass the control's parent
    SetWindowLong ParentWnd, GWL_WNDPROC, AddressOf CtlWheelProc
    
End Sub

Public Sub UnSubclass(MWCtl As CtlWheel, ParentWnd As Long)

    Dim OldProc As Long

    OldProc = GetProp(MWCtl.hwnd, "OldWheelProc")
    
    If OldProc = 0 Then
        Exit Sub
    End If
    
    'Unsubclass control's parent
    SetWindowLong ParentWnd, GWL_WNDPROC, OldProc
    
    'Clean up properties
    RemoveProp ParentWnd, "WheelWnd"
    RemoveProp MWCtl.hwnd, "WheelPtr"
    RemoveProp MWCtl.hwnd, "OldWheelProc"
     
End Sub

