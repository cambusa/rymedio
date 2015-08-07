VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ryMedio v1.3"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   Icon            =   "about.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pctFuoco 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   2880
      ScaleHeight     =   435
      ScaleWidth      =   825
      TabIndex        =   6
      Top             =   2520
      Width           =   825
   End
   Begin VB.CommandButton btChiudi 
      Caption         =   "&Close"
      Height          =   285
      Left            =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2580
      Width           =   855
   End
   Begin VB.Label lbDescr 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.rudyz.net"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   3
      Left            =   1230
      TabIndex        =   7
      Top             =   1440
      Width           =   2490
   End
   Begin VB.Label lbDescr 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "https://github.com/cambusa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   8
      Left            =   1230
      TabIndex        =   4
      Top             =   1860
      Width           =   3465
   End
   Begin VB.Label lbDescr 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Web sites"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   3
      Top             =   1500
      Width           =   705
   End
   Begin VB.Label lbDescr 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Powered by"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   900
      Width           =   840
   End
   Begin VB.Label lbDescr 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "VBScript debugger"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   180
      Width           =   3210
   End
   Begin VB.Label lbDescr 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Le Cose di Rudy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   0
      Left            =   1230
      TabIndex        =   0
      Top             =   780
      Width           =   2670
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ModalMode    As Boolean

Private Sub btChiudi_Click()

    On Error Resume Next
    
    Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    On Error Resume Next
    
    If KeyAscii = 27 Then
    
        Unload Me
         
    End If

End Sub

Private Sub Form_Load()

    On Error Resume Next
    
    If ModalMode Then
        PrimoPiano Me, True
    End If
     
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    
    If ModalMode Then
        PrimoPiano Me, False
    End If
     
End Sub

