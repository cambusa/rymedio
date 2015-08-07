VERSION 5.00
Begin VB.Form frmGuida 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Guida"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   Icon            =   "guida.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btChiudi 
      Caption         =   "&Chiudi"
      Height          =   285
      Left            =   150
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3840
      Width           =   855
   End
   Begin VB.PictureBox pctComandi 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3525
      Left            =   120
      ScaleHeight     =   3525
      ScaleWidth      =   5265
      TabIndex        =   0
      Top             =   120
      Width           =   5265
   End
End
Attribute VB_Name = "frmGuida"
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next
    
    If KeyCode = 27 Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()

    On Error Resume Next
    
    Dim PiX   As Long
    Dim PiY   As Long
    
    Const OffsetVal = 170
    
    If ModalMode Then
        PrimoPiano Me, True
    End If
    
    PiX = Screen.TwipsPerPixelX
    PiY = Screen.TwipsPerPixelY
    
    pctComandi.Cls
    
    pctComandi.FontSize = 10
    pctComandi.FontBold = True: pctComandi.Print "Tastiera"
    pctComandi.FontSize = 8
    
    pctComandi.CurrentY = pctComandi.CurrentY + 4 * PiY
    
    pctComandi.CurrentX = 5 * PiX
    pctComandi.FontBold = True
    pctComandi.Print " Singola istruzione: ";
    pctComandi.CurrentX = OffsetVal * PiX
    pctComandi.FontBold = False
    pctComandi.Print "<SHIFT><F8>"
    
    pctComandi.CurrentX = 5 * PiX
    pctComandi.FontBold = True
    pctComandi.Print " Entra nelle subroutine: ";
    pctComandi.CurrentX = OffsetVal * PiX
    pctComandi.FontBold = False
    pctComandi.Print "<F8>"
    
    pctComandi.CurrentX = 5 * PiX
    pctComandi.FontBold = True
    pctComandi.Print " Esecuzione normale: ";
    pctComandi.CurrentX = OffsetVal * PiX
    pctComandi.FontBold = False
    pctComandi.Print "<F5>"
    
    pctComandi.CurrentX = 5 * PiX
    pctComandi.FontBold = True
    pctComandi.Print " Ferma l'esecuzione: ";
    pctComandi.CurrentX = OffsetVal * PiX
    pctComandi.FontBold = False
    pctComandi.Print "<BREAK>"
    
    pctComandi.CurrentX = 5 * PiX
    pctComandi.FontBold = True
    pctComandi.Print " Termina: ";
    pctComandi.CurrentX = OffsetVal * PiX
    pctComandi.FontBold = False
    pctComandi.Print "<ESC>"
    
    pctComandi.CurrentY = pctComandi.CurrentY + 10 * Screen.TwipsPerPixelY
    
    pctComandi.FontSize = 10
    pctComandi.FontBold = True: pctComandi.Print "Metacomandi"
    pctComandi.FontSize = 8
    
    pctComandi.CurrentY = pctComandi.CurrentY + 4 * Screen.TwipsPerPixelY
    
    pctComandi.CurrentX = 5 * PiX
    pctComandi.FontBold = True
    pctComandi.Print " Valutazione: ";
    pctComandi.CurrentX = OffsetVal * PiX
    pctComandi.FontBold = False
    pctComandi.Print "'@Watch <espressione>"
    
    pctComandi.CurrentX = 5 * PiX
    pctComandi.FontBold = True
    pctComandi.Print " Valutazione membri: ";
    pctComandi.CurrentX = OffsetVal * PiX
    pctComandi.FontBold = False
    pctComandi.Print "'@Inside <espressione>"
    
    pctComandi.CurrentX = 5 * PiX
    pctComandi.FontBold = True
    pctComandi.Print " Ferma alla riga: ";
    pctComandi.CurrentX = OffsetVal * PiX
    pctComandi.FontBold = False
    pctComandi.Print "'@Stop"
    
    pctComandi.CurrentX = 5 * PiX
    pctComandi.FontBold = True
    pctComandi.Print " Salta alla riga: ";
    pctComandi.CurrentX = OffsetVal * PiX
    pctComandi.FontBold = False
    pctComandi.Print "'@Goto <riga>"
    
    pctComandi.CurrentX = 5 * PiX
    pctComandi.FontBold = True
    pctComandi.Print " Termina: ";
    pctComandi.CurrentX = OffsetVal * PiX
    pctComandi.FontBold = False
    pctComandi.Print "'@End"
    
    pctComandi.CurrentX = 5 * PiX
    pctComandi.FontBold = True
    pctComandi.Print " Commento persistente: ";
    pctComandi.CurrentX = OffsetVal * PiX
    pctComandi.FontBold = False
    pctComandi.Print "'@Rem (o anche solo Rem)"
    
    pctComandi.CurrentY = pctComandi.CurrentY + 4 * Screen.TwipsPerPixelY
    
    pctComandi.CurrentX = 5 * PiX
    pctComandi.FontBold = False: pctComandi.Print " Ogni istruzione VBScript può essere posta"
    pctComandi.CurrentX = 5 * PiX
    pctComandi.FontBold = False: pctComandi.Print " sotto metacomando '@"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    
    If ModalMode Then
        PrimoPiano Me, False
    End If
     
End Sub

