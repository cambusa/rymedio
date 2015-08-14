VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Termina"
      Height          =   765
      Left            =   3030
      TabIndex        =   1
      Top             =   330
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   390
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents ObjDebug   As DebugParser
Attribute ObjDebug.VB_VarHelpID = -1
     
Dim CL         As Collection
Dim ObjXml     As DOMDocument30

Private Sub Command1_Click()

    Dim ObjMonitor As DebugMonitor

    Dim C1    As Collection
    Dim C2    As Collection
    Dim C3    As Collection

    Set C1 = New Collection
    Set C2 = New Collection
    Set C3 = New Collection

    C1.Add "Pippo"

    C2.Add "Agenore"
    C2.Add "Flangio"

    Set CL = New Collection

    CL.Add C1, "P"
    CL.Add C2, "Q"
    CL.Add C3, "R"

    Set ObjXml = New DOMDocument30
    ObjXml.loadXML "<xml tipo='spiaggia'><param/><dati nome='pippo' cognome='canino' /></xml>"


    Set ObjMonitor = New DebugMonitor
    ObjMonitor.QueryType = qtEvent
    ObjMonitor.Polling = 200

    ObjDebug.ProcID = ObjMonitor.ProcID
    Set ObjDebug.Monitor = ObjMonitor

    ObjMonitor.DisplayMonitor

    ObjDebug.SourceFile App.Path + "\prova.vbs"
    ObjDebug.Execute "Main"

    ObjMonitor.HideMonitor
    Set ObjMonitor = Nothing

    Set CL = Nothing

    Set C1 = Nothing
    Set C2 = Nothing
    Set C3 = Nothing
     
End Sub

Private Sub Command2_Click()

    ObjDebug.Terminate

End Sub

Private Sub Form_Load()

    Set ObjDebug = New DebugParser
    ObjDebug.QueryType = qtEvent
     
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set ObjDebug = Nothing

End Sub


Private Sub ObjDebug_InitScript(ScriptObject As Object)

    ScriptObject.addObject "Frm", Me
    ScriptObject.addObject "CL", CL
    ScriptObject.addObject "DXML", ObjXml

End Sub

