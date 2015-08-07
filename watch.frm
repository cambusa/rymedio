VERSION 5.00
Begin VB.Form FrmWatch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aggiunta watch"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10995
   Icon            =   "watch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   10995
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btChiudi 
      Caption         =   "&Chiudi"
      Height          =   285
      Left            =   1110
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1110
      Width           =   855
   End
   Begin VB.CommandButton btAggiungi 
      Caption         =   "&Aggiungi"
      Height          =   285
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1110
      Width           =   855
   End
   Begin VB.TextBox txRiga 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   450
      Width           =   10725
   End
   Begin VB.Label lbDescr 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Seleziona l'espressione da aggiungere"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   2700
   End
End
Attribute VB_Name = "FrmWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Cancel       As Boolean
Public Espressione  As String
Public Riga         As String
Public Codice       As String

Dim ObjScript       As ScriptControl

Dim MessaggioNiente As String

Private Sub btAggiungi_Click()

    On Error Resume Next
    
    Dim Valida     As Boolean
    
    Valida = False
    Me.Espressione = ""
    
    If Me.Riga <> "" Then
    
        If txRiga.SelLength > 0 Then
        
            Me.Espressione = Trim(Mid(txRiga, txRiga.SelStart + 1, txRiga.SelLength))
             
        End If
        
    Else
    
        Me.Espressione = Trim(txRiga)
    
    End If
    
    If Me.Espressione <> "" Then
    
        Valida = True
    
    End If
             
    If Valida Then
    
        ObjScript.Reset
        ObjScript.AddCode Me.Codice
        
        ObjScript.Error.Clear
        ObjScript.AddCode "sub test___tmp():x=" & Me.Espressione & ":end sub"
        
        If ObjScript.Error.Number = 0 Then
        
            Me.Cancel = False
            Me.Hide
             
        Else
        
            MsgBox ObjScript.Error.Description, , "Aggiunta watch"
        
        End If
        
    Else
    
        MsgBox MessaggioNiente, , "Aggiunta watch"
    
    End If
     
End Sub

Private Sub btChiudi_Click()

    On Error Resume Next
    
    Me.Hide

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next
    
    If KeyCode = 27 Then
    
        btChiudi_Click
    
    ElseIf KeyCode = 13 Then
    
        btAggiungi_Click
    
    End If

End Sub

Private Sub Form_Load()

    On Error Resume Next
    
    Set ObjScript = New ScriptControl
    ObjScript.Language = "vbscript"
    ObjScript.Timeout = -1
    
    txRiga = Trim(Me.Riga)
    
    If Me.Riga <> "" Then
        MessaggioNiente = "Nessuna espressione selezionata"
        lbDescr.Caption = "Seleziona l'espressione da aggiungere"
    Else
        MessaggioNiente = "Specificare l'espressione da aggiungere"
        lbDescr.Caption = "Espressione da aggiungere"
    End If
    
    Me.Espressione = ""
    Me.Codice = ""
    
    Me.Cancel = True
     
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    
    Set ObjScript = Nothing

End Sub
