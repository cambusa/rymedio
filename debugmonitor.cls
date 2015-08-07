VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DebugMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim WithEvents FrmDebug  As FrmMonitor
Attribute FrmDebug.VB_VarHelpID = -1
Dim PropDirectory        As String
Dim PropQueryType        As EnumQueryType
Dim PropProcID           As String
Dim PropPolling          As Long
Dim PropModalMode        As Boolean

Public Event WriteCommand(strCommand As String)
Public Event ReadMonitor(strCode As String)
Public Event ReadStatus(strStatus As String)
Public Event StatusExists(ThereIs As Boolean)
Public Event MonitorExists(ThereIs As Boolean)
     
Public Sub DisplayMonitor()

    On Error Resume Next
    
    If Not FrmDebug Is Nothing Then
    
        Unload FrmDebug
        Set FrmDebug = Nothing
         
    End If
    
    Set FrmDebug = New FrmMonitor
    
    FrmDebug.Directory = Me.Directory
    FrmDebug.QueryType = Me.QueryType
    FrmDebug.ProcID = Me.ProcID
    FrmDebug.Polling = Me.Polling
    FrmDebug.ModalMode = PropModalMode
         
    FrmDebug.Show
     
End Sub

Public Sub HideMonitor()

    On Error Resume Next
    
    If Not FrmDebug Is Nothing Then
    
        Unload FrmDebug
        Set FrmDebug = Nothing
         
    End If
     
End Sub

Private Sub Class_Initialize()

    On Error Resume Next
    
    PropQueryType = qtEvent
    PropPolling = 200
    PropModalMode = False

End Sub

Private Sub Class_Terminate()

    On Error Resume Next

    Unload FrmDebug
    Set FrmDebug = Nothing

End Sub

Public Property Get Directory() As String

    On Error Resume Next
    
    If PropDirectory = "" Then
    
        PropDirectory = App.Path + "\rymedio"
        
        If PropQueryType = qtFile Then
        
            If Dir(PropDirectory, vbDirectory) = "" Then
                MkDir PropDirectory
            End If
             
        End If
         
    End If
    
    Directory = PropDirectory
     
End Property

Public Property Let Directory(ByVal NewValue As String)

    On Error Resume Next
    
    PropDirectory = NewValue
    
    If Dir(PropDirectory, vbDirectory) = "" Then
        MkDir PropDirectory
    End If
     
End Property

Public Property Get QueryType() As EnumQueryType

    On Error Resume Next

    QueryType = PropQueryType

End Property

Public Property Let QueryType(ByVal NewValue As EnumQueryType)

    On Error Resume Next
    
    PropQueryType = NewValue

End Property

Private Sub FrmDebug_ReadMonitor(strCode As String)

    On Error Resume Next
    
    RaiseEvent ReadMonitor(strCode)

End Sub

Private Sub FrmDebug_ReadStatus(strStatus As String)

    On Error Resume Next
    
    RaiseEvent ReadStatus(strStatus)

End Sub

Private Sub FrmDebug_StatusExists(ThereIs As Boolean)

    On Error Resume Next
    
    RaiseEvent StatusExists(ThereIs)

End Sub

Private Sub FrmDebug_MonitorExists(ThereIs As Boolean)

    On Error Resume Next
    
    RaiseEvent MonitorExists(ThereIs)

End Sub

Private Sub FrmDebug_WriteCommand(strCommand As String)

    On Error Resume Next
    
    RaiseEvent WriteCommand(strCommand)

End Sub

Public Property Get ProcID() As String

    On Error Resume Next

    If PropProcID = "" Then
     
        Randomize Timer
        PropProcID = Format(Now, "yyyymmddhhmmss") + Right("0000" & CLng(Rnd * 1000), 4)
     
    End If
    
    ProcID = PropProcID

End Property

Public Property Let ProcID(ByVal NewValue As String)

    On Error Resume Next

    PropProcID = NewValue

End Property

Public Property Get Polling() As Long

    On Error Resume Next
    
    Polling = PropPolling

End Property

Public Property Let Polling(ByVal NewValue As Long)

    On Error Resume Next
    
    PropPolling = NewValue

End Property

Public Property Get Handle() As Long

    On Error Resume Next
    
    If FrmDebug Is Nothing Then
        Handle = 0
    Else
        Handle = FrmDebug.hwnd
    End If

End Property

Public Property Get ModalMode() As Boolean

    On Error Resume Next
    
    ModalMode = PropModalMode

End Property

Public Property Let ModalMode(ByVal NewValue As Boolean)

    On Error Resume Next
    
    PropModalMode = NewValue

End Property




