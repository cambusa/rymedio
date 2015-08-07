VERSION 5.00
Begin VB.Form FrmMonitor 
   Caption         =   "ryMedio - Monitor"
   ClientHeight    =   9555
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   15090
   Icon            =   "monitor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9555
   ScaleWidth      =   15090
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerResize 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   14640
      Top             =   1890
   End
   Begin VB.PictureBox pctBarraVert 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9315
      Left            =   11190
      MousePointer    =   9  'Size W E
      ScaleHeight     =   9315
      ScaleWidth      =   45
      TabIndex        =   3
      Top             =   150
      Width           =   45
   End
   Begin VB.PictureBox pctBarraOrizz 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   90
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   14955
      TabIndex        =   4
      Top             =   9480
      Visible         =   0   'False
      Width           =   14955
   End
   Begin VB.Timer TimerComando 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   14640
      Top             =   1230
   End
   Begin VB.Timer TimerMonitor 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   14610
      Top             =   660
   End
   Begin VB.Timer TimerStatus 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   14610
      Top             =   120
   End
   Begin VB.PictureBox pctAttesa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   525
      Left            =   4350
      ScaleHeight     =   495
      ScaleWidth      =   5535
      TabIndex        =   2
      Top             =   4140
      Visible         =   0   'False
      Width           =   5565
   End
   Begin VB.TextBox txValutazioni 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9375
      Left            =   11250
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   90
      Width           =   3795
   End
   Begin ryMedio.CtlMonitor TxMonitor 
      Height          =   9375
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   16536
   End
   Begin VB.Menu MnFile 
      Caption         =   "&File"
      Begin VB.Menu MnEsci 
         Caption         =   "&Esci"
      End
   End
   Begin VB.Menu MnVisualizza 
      Caption         =   "&Visualizza"
      Begin VB.Menu MnOrizzontale 
         Caption         =   "Ori&zzontale"
      End
      Begin VB.Menu MnVerticale 
         Caption         =   "Ver&ticale"
      End
   End
   Begin VB.Menu MnDebug 
      Caption         =   "&Debug"
      Begin VB.Menu MnAggiungiWatch2 
         Caption         =   "Aggiungi watch"
      End
      Begin VB.Menu MnDelBreakpoint2 
         Caption         =   "Elimina tutte le interruzioni"
      End
   End
   Begin VB.Menu MnEsegui 
      Caption         =   "&Esegui"
      Begin VB.Menu MnIstruzione2 
         Caption         =   "&Istruzione (SHIFT+F8)"
      End
      Begin VB.Menu MnNextSub2 
         Caption         =   "Entrando nelle subroutine (F8)"
      End
      Begin VB.Menu MnFree2 
         Caption         =   "Senza inte&rruzioni (F5)"
      End
      Begin VB.Menu MnInterrompi 
         Caption         =   "Interrompi (BREAK)"
      End
   End
   Begin VB.Menu MnHelp 
      Caption         =   "&?"
      Begin VB.Menu MnGuida 
         Caption         =   "&Guida"
      End
      Begin VB.Menu MnInformazioni 
         Caption         =   "Informazioni su r&yMedio"
      End
   End
   Begin VB.Menu MnPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu MnBreakPoint 
         Caption         =   "Esegui fino alla riga"
      End
      Begin VB.Menu MnGoTo 
         Caption         =   "Salta direttamente alla riga"
      End
      Begin VB.Menu MnMostraRiga 
         Caption         =   "Mostra riga corrente"
      End
      Begin VB.Menu MnSep 
         Caption         =   "-"
      End
      Begin VB.Menu MnAggiungiWatch 
         Caption         =   "Aggiungi watch"
      End
      Begin VB.Menu MnInterruzione 
         Caption         =   "Aggiungi interruzione"
      End
      Begin VB.Menu MnDelBreakpoint 
         Caption         =   "Elimina tutte le interruzioni"
      End
      Begin VB.Menu MnSep2 
         Caption         =   "-"
      End
      Begin VB.Menu MnIstruzione 
         Caption         =   "Esegui istruzione"
      End
      Begin VB.Menu MnNextSub 
         Caption         =   "Esegui entrando nelle subroutine"
      End
      Begin VB.Menu MnFree 
         Caption         =   "Esecuzione senza interruzioni"
      End
      Begin VB.Menu MnTermina 
         Caption         =   "Termina esecuzione"
      End
   End
End
Attribute VB_Name = "FrmMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Directory    As String
Public QueryType    As Long
Public ProcID       As String
Public Polling      As Long
Public ModalMode    As Boolean

Public Event WriteCommand(strCommand As String)
Public Event ReadMonitor(strCode As String)
Public Event ReadStatus(strStatus As String)
Public Event StatusExists(ThereIs As Boolean)
Public Event MonitorExists(ThereIs As Boolean)
     
Dim PuntiAttesa     As Long
Dim UltimaIstr      As Long

Dim UltimoComando   As Single

Dim Visualizzazione As Long
Dim InizialeW       As Single
Dim InizialeH       As Single
Dim PosBarra        As Single
Dim ContaResize     As Long

Const CorrezioneW = 20
Const CorrezioneH = 64
     
Dim StatoWin        As StructStatoWin
Dim FlagActivate    As Boolean
     
Public Sub ScriviComando(Comando As String)

    On Error GoTo ErrorProcedure
    
    Dim NumFile    As String
    Dim PathFile   As String
    
    If pctAttesa.Visible = False Or Comando = "break:0" Then
    
        If 1000 * Abs(Timer - UltimoComando) > Me.Polling Or Comando = "break:0" Then
    
            '------------------------------
            ' Accendo il segnale di attesa
            '------------------------------
            
            PuntiAttesa = 3
            AvanzaAttesa
            pctAttesa.Visible = True
            
            If QueryType = qtFile Then
            
                PathFile = Me.Directory + "\" + Me.ProcID + ".CMN"
            
                NumFile = FreeFile
                Open PathFile For Output As #NumFile
                Print #NumFile, Comando
                Close #NumFile
                NumFile = 0
            
            Else
            
                RaiseEvent WriteCommand(Comando)
            
            End If
            
            UltimoComando = Timer
            
            TimerStatus.Enabled = True
                            
        End If
        
    End If

Exit Sub

ErrorProcedure:

    Resume AbortProcedure

AbortProcedure:

    On Error Resume Next
    
    If NumFile > 0 Then
        Close #NumFile
        NumFile = 0
    End If
     
End Sub

Private Sub ComandoBreak(Riga As Long)

    On Error Resume Next
    
    ScriviComando "break:" & Riga

End Sub

Private Sub ComandoWatch(Espressione As String)

    On Error Resume Next
    
    ScriviComando "watch:" & Espressione

End Sub

Private Sub ComandoEnd()

    On Error Resume Next
    
    ScriviComando "end"

End Sub

Private Sub ComandoGoTo(Riga As Long)

    On Error Resume Next
    
    ScriviComando "goto:" & Riga

End Sub

Private Sub ComandoLibera()

    On Error Resume Next
    
    Dim Buffer     As String
     
    Buffer = TxMonitor.ListBreakPoint
    
    ScriviComando "free:" & Buffer

End Sub

Private Sub ComandoNext()

    On Error Resume Next
    
    ScriviComando "next"

End Sub

Private Sub ComandoNextSub()

    On Error Resume Next
    
    ScriviComando "next:1"

End Sub

Private Sub Form_Activate()

    On Error Resume Next
    
    If FlagActivate = False Then
    
        FlagActivate = True
        
        If StatoWin.M = 0 Then
        
            Me.Move StatoWin.L, StatoWin.T, StatoWin.W, StatoWin.H
            
        End If
        
        If Visualizzazione = 0 Then
        
            MnOrizzontale.Checked = False
            MnVerticale.Checked = True
            pctBarraVert.Left = StatoWin.B
            SpostaOrizz
            
        Else
        
            MnOrizzontale.Checked = True
            MnVerticale.Checked = False
            pctBarraOrizz.Top = StatoWin.B
            SpostaVert
            
        End If
        
        TimerResize.Enabled = True
         
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next
    
    Select Case KeyCode
    
    Case vbKeyF8
    
        If Shift = 1 Then
            ComandoNext
        ElseIf Shift = 0 Then
            ComandoNextSub
        End If
         
    Case vbKeyF5
    
        ComandoLibera
         
    Case vbKeyEscape
    
        ComandoEnd
         
    Case 19
    
        If pctAttesa.Visible Then
            ComandoBreak 0
        End If
        
    End Select

End Sub

Private Sub Form_Load()

    On Error Resume Next
    
    If ModalMode Then
        PrimoPiano Me, True
    End If
    
    UltimaIstr = 0
    InizialeW = Me.Width
    InizialeH = Me.Height
    
    FlagActivate = False
    
    StatoWin.L = Val(RegReadSetting(RegCurrentUser, "ryMedio", "", "Left", Str((Screen.Width - InizialeW) / 2)))
    StatoWin.T = Val(RegReadSetting(RegCurrentUser, "ryMedio", "", "Top", Str((Screen.Height - InizialeH) / 2)))
    StatoWin.W = Val(RegReadSetting(RegCurrentUser, "ryMedio", "", "Width", Str(InizialeW)))
    StatoWin.H = Val(RegReadSetting(RegCurrentUser, "ryMedio", "", "Height", Str(InizialeH)))
    StatoWin.M = Val(RegReadSetting(RegCurrentUser, "ryMedio", "", "WindowState", "0"))
    StatoWin.B = Val(RegReadSetting(RegCurrentUser, "ryMedio", "", "Offset", Str(pctBarraVert.Left)))
    Visualizzazione = Val(RegReadSetting(RegCurrentUser, "ryMedio", "", "TileMode", "0"))
    
    If StatoWin.M = 1 Then
        StatoWin.M = 0
    End If
    
    Me.WindowState = StatoWin.M
    
    '------------------------------
    ' Accendo il segnale di attesa
    '------------------------------

    PuntiAttesa = 3
    AvanzaAttesa
    pctAttesa.Visible = True
    
    TimerStatus.Interval = Me.Polling
    TimerMonitor.Enabled = True

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    
    TxMonitor.SetFocus

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    DisposizioneForm
     
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    
    TimerMonitor.Enabled = False
    TimerStatus.Enabled = False
    
    If Me.WindowState = 0 Then
    
        RegWriteSetting RegCurrentUser, "ryMedio", "", "Left", Str(Me.Left)
        RegWriteSetting RegCurrentUser, "ryMedio", "", "Top", Str(Me.Top)
        RegWriteSetting RegCurrentUser, "ryMedio", "", "Width", Str(Me.Width)
        RegWriteSetting RegCurrentUser, "ryMedio", "", "Height", Str(Me.Height)
         
    End If
    
    RegWriteSetting RegCurrentUser, "ryMedio", "", "WindowState", Str(Me.WindowState)
    RegWriteSetting RegCurrentUser, "ryMedio", "", "TileMode", Str(Visualizzazione)
    
    If Visualizzazione = 0 Then
        RegWriteSetting RegCurrentUser, "ryMedio", "", "Offset", Str(pctBarraVert.Left)
    Else
        RegWriteSetting RegCurrentUser, "ryMedio", "", "Offset", Str(pctBarraOrizz.Top)
    End If
    
    ScriviComando "end"

    If ModalMode Then
        PrimoPiano Me, False
    End If
     
End Sub

Private Sub MnInterruzione_Click()

    On Error Resume Next
    
    Dim Rg    As Long
    
    Rg = TxMonitor.RigaCliccata
    
    If Rg > 0 Then
    
        TxMonitor.BreakPoint(Rg) = Not TxMonitor.BreakPoint(Rg)
        TxMonitor.Refresh
         
    End If

End Sub

Private Sub MnAggiungiWatch_Click()

    On Error Resume Next
    
    AggiungiWatch -1

End Sub

Private Sub MnAggiungiWatch2_Click()

    On Error Resume Next
    
    AggiungiWatch TxMonitor.RigaCliccata

End Sub

Private Sub MnBreakPoint_Click()

    On Error Resume Next
    
    TimerComando.Tag = "break:" & TxMonitor.RigaCliccata
    TimerComando.Enabled = True

End Sub

Private Sub MnDelBreakpoint_Click()

    On Error Resume Next
    
    TxMonitor.ClearBreakPoint

End Sub

Private Sub MnDelBreakpoint2_Click()

    On Error Resume Next
    
    TxMonitor.ClearBreakPoint

End Sub

Private Sub MnEsci_Click()

    On Error Resume Next
    
    Unload Me

End Sub

Private Sub MnEsegui_Click()

    On Error Resume Next
    
    MnInterrompi.Enabled = pctAttesa.Visible

End Sub

Private Sub MnFree_Click()

    On Error Resume Next
    
    Dim Buffer     As String
    
    Buffer = TxMonitor.ListBreakPoint
    
    TimerComando.Tag = "free:" & Buffer
    TimerComando.Enabled = True

End Sub

Private Sub MnFree2_Click()

    On Error Resume Next
    
    Dim Buffer     As String
    
    Buffer = TxMonitor.ListBreakPoint
    
    TimerComando.Tag = "free:" & Buffer
    TimerComando.Enabled = True

End Sub

Private Sub MnGoTo_Click()

    On Error Resume Next
    
    TimerComando.Tag = "goto:" & TxMonitor.RigaCliccata
    TimerComando.Enabled = True

End Sub

Private Sub MnGuida_Click()

    On Error Resume Next
    
    If ModalMode Then
        PrimoPiano Me, False
    End If
    
    frmGuida.ModalMode = ModalMode
    frmGuida.Show 1
    Unload frmGuida
    Set frmGuida = Nothing
    
    If ModalMode Then
        PrimoPiano Me, True
    End If
     
End Sub

Private Sub MnInformazioni_Click()

    On Error Resume Next
    
    If ModalMode Then
        PrimoPiano Me, False
    End If
    
    frmAbout.ModalMode = ModalMode
    frmAbout.Show 1
    Unload frmAbout
    Set frmAbout = Nothing
    
    If ModalMode Then
        PrimoPiano Me, True
    End If
     
End Sub

Private Sub MnInterrompi_Click()

    On Error Resume Next
    
    TimerComando.Tag = "break:0"
    TimerComando.Enabled = True

End Sub

Private Sub MnIstruzione_Click()

    On Error Resume Next
    
    TimerComando.Tag = "next"
    TimerComando.Enabled = True

End Sub

Private Sub MnIstruzione2_Click()

    On Error Resume Next
    
    TimerComando.Tag = "next"
    TimerComando.Enabled = True

End Sub

Private Sub MnMostraRiga_Click()

    On Error Resume Next
    
    If TxMonitor.Riga - TxMonitor.VisibleRows \ 2 > 0 Then
        TxMonitor.SetTopRow TxMonitor.Riga - TxMonitor.VisibleRows \ 2
    Else
        TxMonitor.SetTopRow 0
    End If
    
    TxMonitor.Refresh
    
End Sub

Private Sub MnNextSub_Click()

    On Error Resume Next
    
    TimerComando.Tag = "next:1"
    TimerComando.Enabled = True

End Sub

Private Sub MnNextSub2_Click()

    On Error Resume Next
    
    ScriviComando "next:1"

End Sub

Private Sub MnOrizzontale_Click()

    Visualizzazione = 1
    MnOrizzontale.Checked = True
    MnVerticale.Checked = False
    DisposizioneForm

End Sub

Private Sub MnTermina_Click()

    On Error Resume Next
    
    TimerComando.Tag = "end"
    TimerComando.Enabled = True

End Sub

Private Sub MnVerticale_Click()

    On Error Resume Next
    
    Visualizzazione = 0
    MnOrizzontale.Checked = False
    MnVerticale.Checked = True
    DisposizioneForm

End Sub

Private Sub pctBarraOrizz_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    
    If Button = 1 Then
        PosBarra = Y
    End If

End Sub

Private Sub pctBarraOrizz_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    
    Dim NuovaPos    As Single
    
    If Button = 1 Then
    
        NuovaPos = pctBarraOrizz.Top + Y - PosBarra
        
        If NuovaPos > 200 * Screen.TwipsPerPixelY And NuovaPos < Me.Height - 100 * Screen.TwipsPerPixelY Then
        
            pctBarraOrizz.Top = NuovaPos
            SpostaVert
        
        End If
    
    End If
    
End Sub

Private Sub pctBarraOrizz_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    
    TxMonitor.SetFocus

End Sub

Private Sub pctBarraVert_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    
    If Button = 1 Then
        PosBarra = X
    End If

End Sub

Private Sub pctBarraVert_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    
    Dim NuovaPos    As Single
    
    If Button = 1 Then
    
        NuovaPos = pctBarraVert.Left + X - PosBarra
        
        If NuovaPos > 200 * Screen.TwipsPerPixelX And NuovaPos < Me.Width - 200 * Screen.TwipsPerPixelX Then
        
            pctBarraVert.Left = NuovaPos
            SpostaOrizz
        
        End If
    
    End If
     
End Sub

Private Sub pctBarraVert_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    
    TxMonitor.SetFocus

End Sub

Private Sub TimerComando_Timer()

    On Error Resume Next
    
    TimerComando.Enabled = False
    ScriviComando TimerComando.Tag
     
End Sub

Private Sub TimerMonitor_Timer()

    On Error Resume Next
    
    Dim Buffer          As String
    Dim NumFile         As Integer
    Dim PathFile        As String
    Dim Esiste          As Boolean
    
    If QueryType = qtFile Then
    
        PathFile = Me.Directory + "\" + Me.ProcID + ".MON"
        
        Esiste = Dir(PathFile) <> ""
         
    Else
    
        Esiste = False
        RaiseEvent MonitorExists(Esiste)
    
    End If
    
    If Esiste Then
    
        TimerMonitor.Enabled = False
        
        If QueryType = qtFile Then
        
            NumFile = FreeFile
            Open PathFile For Binary As #NumFile
            Buffer = Space(LOF(NumFile))
            Get #NumFile, , Buffer
            Close #NumFile
            NumFile = 0
            
        Else
        
            RaiseEvent ReadMonitor(Buffer)
        
        End If
        
        TxMonitor = Buffer
        
        TimerStatus.Enabled = True
        
    End If
     
End Sub

Private Sub TimerResize_Timer()

    On Error Resume Next
    
    TimerResize.Enabled = False
    
    DisposizioneForm
     
End Sub

Private Sub TimerStatus_Timer()

    On Error Resume Next
    
    Dim Buffer          As String
    Dim NumFile         As Integer
    Dim PathFile        As String
    
    Dim Riga            As Long
    Dim Commento        As String
    Dim MessErr         As String
    Dim MaxEval         As Long
    Dim I               As Long
    Dim Esiste          As Boolean
    Dim V               As Variant
    
    If QueryType = qtFile Then
    
        PathFile = Me.Directory + "\" + Me.ProcID + ".STA"
        
        Esiste = (Dir(PathFile) <> "")
        
    Else
    
        Esiste = False
        RaiseEvent StatusExists(Esiste)
    
    End If
    
    If Esiste Then
    
        TimerStatus.Enabled = False
        
        If QueryType = qtFile Then
        
            NumFile = FreeFile
            Open PathFile For Binary As #NumFile
            Buffer = Space(LOF(NumFile))
            Get #NumFile, , Buffer
            Close #NumFile
            NumFile = 0
            
            Kill PathFile
            
        Else
        
            RaiseEvent ReadStatus(Buffer)
        
        End If
        
        V = Split(Buffer, vbCrLf)
        
        MessErr = V(0)
        
        If MessErr = "end" Then
        
            Unload Me
        
        Else
        
            Commento = V(1)
            UltimaIstr = Val(V(2))
            Riga = Val(V(3))
            
            '-------------
            ' Valutazioni
            '-------------
            
            MaxEval = UBound(V) - 3
            
            If Commento <> "" Then
                Buffer = "Commento: " + Commento + vbCrLf
            Else
                Buffer = ""
            End If
            
            For I = 1 To MaxEval
                Buffer = Buffer & V(I + 3) & vbCrLf
            Next I
            
            txValutazioni = Buffer
            
            '---------
            ' Monitor
            '---------
            
            If Riga <= 0 Then
                Riga = 1
            End If
            
            TxMonitor.Riga = Riga - 1
            
            If Riga > TxMonitor.VisibleRows \ 2 Then
                TxMonitor.SetTopRow Riga - TxMonitor.VisibleRows \ 2
            End If
            
            TxMonitor.Refresh
            
            If MessErr <> "" Then
                MsgBox MessErr, , "ryMedio - Riscontrato errore!"
            End If
        
        End If
        
        '-----------------------------
        ' Spengo il segnale di attesa
        '-----------------------------
        
        pctAttesa.Visible = False
    
    Else
    
        AvanzaAttesa
    
    End If

End Sub

Private Sub AvanzaAttesa()

    On Error Resume Next
    
    pctAttesa.Cls
    pctAttesa.CurrentX = 130 * Screen.TwipsPerPixelX
    pctAttesa.CurrentY = 5 * Screen.TwipsPerPixelY
    pctAttesa.Print "Attendere" + String(PuntiAttesa, ".")
    PuntiAttesa = PuntiAttesa + 1
    
    If PuntiAttesa > 10 Then
        PuntiAttesa = 0
    End If
     
End Sub

Private Sub TxMonitor_DblClick(Riga As Long, Numerazione As Boolean)

    On Error Resume Next
    
    If Numerazione Then
    
        TxMonitor.BreakPoint(Riga) = Not TxMonitor.BreakPoint(Riga)
        TxMonitor.Refresh
         
    Else
    
        AggiungiWatch Riga
    
    End If

End Sub

Private Sub TxMonitor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    
    Dim R     As Long
    
    If Button = 2 Then
    
         If pctAttesa.Visible = False Then
         
            If UltimaIstr = 0 Then
            
                R = TxMonitor.RigaCliccata
            
                If R >= 0 Then
            
                    MnBreakPoint.Enabled = True
                    MnBreakPoint.Caption = "Esegui fino alla riga " & R
                    
                    MnGoTo.Enabled = True
                    MnGoTo.Caption = "Salta direttamente alla riga " & R
                    
                    MnInterruzione.Enabled = True
            
                    If TxMonitor.BreakPoint(R) Then
                        MnInterruzione.Caption = "Togli interruzione alla riga " & R
                    Else
                        MnInterruzione.Caption = "Aggiungi interruzione alla riga " & R
                    End If
            
                Else
            
                    MnBreakPoint.Enabled = False
                    MnBreakPoint.Caption = "Esegui fino alla riga"
                    
                    MnGoTo.Enabled = False
                    MnGoTo.Caption = "Salta direttamente alla riga"
                    
                    MnInterruzione.Enabled = False
                    MnInterruzione.Caption = "Aggiungi interruzione"
                    
                End If
            
                MnAggiungiWatch.Enabled = True
                MnDelBreakpoint.Enabled = True
                MnFree.Enabled = True
                MnIstruzione.Enabled = True
                MnNextSub.Enabled = True
            
            Else
              
                MnBreakPoint.Enabled = False
                MnBreakPoint.Caption = "Esegui fino alla riga"
                
                MnGoTo.Enabled = False
                MnGoTo.Caption = "Salta direttamente alla riga"
                
                MnInterruzione.Enabled = False
                MnInterruzione.Caption = "Aggiungi interruzione"
                        
                MnAggiungiWatch.Enabled = False
                MnDelBreakpoint.Enabled = False
                MnFree.Enabled = False
                MnIstruzione.Enabled = False
                MnNextSub.Enabled = False
                
            End If
              
            PopupMenu MnPopup
                        
        End If
         
    End If

End Sub

Public Sub DisposizioneForm()

    On Error Resume Next
    
    Dim Proporzione     As Single
    Dim EditW           As Single
    Dim EditH           As Single
    Dim ValuW           As Single
    Dim ValuH           As Single
    Dim FlagTimer       As Boolean
    
    ContaResize = ContaResize + 1
    
    If ContaResize = 1 Then
     
        FlagTimer = False
     
        If Me.Width < InizialeW Or Me.Height < InizialeH Then
          
            If Me.Width < InizialeW And Me.Height < InizialeH Then
               
                Me.Move Me.Left, Me.Top, InizialeW, InizialeH
                FlagTimer = True
                    
                    
            ElseIf Me.Width < InizialeW Then
               
                Me.Move Me.Left, Me.Top, InizialeW
                FlagTimer = True
               
            ElseIf Me.Height < InizialeH Then
               
                Me.Move Me.Left, Me.Top, Me.Width, InizialeH
                FlagTimer = True
               
            End If
          
        End If
          
        If FlagTimer Then
          
            TimerResize.Enabled = True
               
        Else
          
            If Visualizzazione = 0 Then
                         
                If pctBarraVert.Visible Then
                    
                    Proporzione = txValutazioni.Width / TxMonitor.Width
                    EditW = (Me.Width - pctBarraVert.Width - CorrezioneW * Screen.TwipsPerPixelX) / (1 + Proporzione)
                         
                Else
                    
                    EditW = (Me.Width - CorrezioneW * Screen.TwipsPerPixelX) * (3 / 4)
                         
                End If
                    
                ValuW = Me.Width - EditW - pctBarraVert.Width - CorrezioneW * Screen.TwipsPerPixelX
                    
                EditH = Me.Height - CorrezioneH * Screen.TwipsPerPixelY
                ValuH = EditH
                    
                TxMonitor.Move TxMonitor.Left, TxMonitor.Top, EditW, EditH
                pctBarraVert.Move TxMonitor.Left + TxMonitor.Width, TxMonitor.Top, pctBarraVert.Width, EditH
                txValutazioni.Move pctBarraVert.Left + pctBarraVert.Width, TxMonitor.Top, ValuW, ValuH
                    
                pctBarraVert.Visible = True
                pctBarraOrizz.Visible = False
                    
            Else
               
                If pctBarraOrizz.Visible Then
                    
                    Proporzione = txValutazioni.Height / TxMonitor.Height
                    EditH = (Me.Height - pctBarraOrizz.Height - CorrezioneH * Screen.TwipsPerPixelY) / (1 + Proporzione)
                         
                Else
                    
                    EditH = (Me.Height - CorrezioneH * Screen.TwipsPerPixelY) * (3 / 4)
                         
                End If
                    
                ValuH = Me.Height - EditH - pctBarraOrizz.Height - CorrezioneH * Screen.TwipsPerPixelX
                    
                EditW = Me.Width - CorrezioneW * Screen.TwipsPerPixelX
                ValuW = EditW
                    
                TxMonitor.Move TxMonitor.Left, TxMonitor.Top, EditW, EditH
                pctBarraOrizz.Move TxMonitor.Left, TxMonitor.Top + TxMonitor.Height, EditW
                txValutazioni.Move TxMonitor.Left, pctBarraOrizz.Top + pctBarraOrizz.Height, ValuW, ValuH
                    
                pctBarraVert.Visible = False
                pctBarraOrizz.Visible = True
               
            End If
               
            DisponiAttendi
          
        End If
          
    End If
     
    ContaResize = ContaResize - 1
     
End Sub

Public Sub SpostaOrizz()

    On Error Resume Next
    
    Dim EditW           As Single
    Dim EditH           As Single
    Dim ValuW           As Single
    Dim ValuH           As Single
    
    ContaResize = ContaResize + 1
    
    If ContaResize = 1 Then
     
        EditW = pctBarraVert.Left - TxMonitor.Left
        ValuW = Me.Width - EditW - pctBarraVert.Width - CorrezioneW * Screen.TwipsPerPixelX
          
        EditH = TxMonitor.Height
        ValuH = EditH
        
        TxMonitor.Move TxMonitor.Left, TxMonitor.Top, EditW, EditH
        txValutazioni.Move pctBarraVert.Left + pctBarraVert.Width, TxMonitor.Top, ValuW, ValuH
        
        'GripBarra pctBarraVert, False
        
        DisponiAttendi
               
    End If
     
    ContaResize = ContaResize - 1
     
End Sub

Public Sub SpostaVert()

    On Error Resume Next
    
    Dim EditW           As Single
    Dim EditH           As Single
    Dim ValuW           As Single
    Dim ValuH           As Single
    
    ContaResize = ContaResize + 1
    
    If ContaResize = 1 Then
    
        EditH = pctBarraOrizz.Top - TxMonitor.Top
        ValuH = Me.Height - EditH - pctBarraOrizz.Height - CorrezioneH * Screen.TwipsPerPixelX
        
        EditW = TxMonitor.Width
        ValuW = EditW
        
        TxMonitor.Move TxMonitor.Left, TxMonitor.Top, EditW, EditH
        txValutazioni.Move TxMonitor.Left, pctBarraOrizz.Top + pctBarraOrizz.Height, ValuW, ValuH
        
        'GripBarra pctBarraVert, False
        
        DisponiAttendi
         
    End If
    
    ContaResize = ContaResize - 1
     
End Sub

Private Sub GripBarra(ObjPic As PictureBox, FlagOrizz As Boolean)

    On Error GoTo ErrorProcedure
    
    Dim mx         As Long
    Dim my         As Long
    Dim dx         As Long
    Dim dy         As Long
    Dim px         As Long
    Dim py         As Long
    Dim P          As Long
    Dim Colore     As Long
    
    ObjPic.Cls
    
    px = Screen.TwipsPerPixelX
    py = Screen.TwipsPerPixelY
    
    mx = (ObjPic.Width / 2)
    my = (ObjPic.Height / 2)
    
    If FlagOrizz Then
        dx = 3 * px
        dy = 0
    Else
        dx = 0
        dy = 3 * py
    End If
         
    For P = -4 To 4
    
        ObjPic.PSet (mx + P * dx, my + P * dy), vbWhite
        ObjPic.PSet (mx + P * dx + px, my + P * dy), vbBlack
        ObjPic.PSet (mx + P * dx, my + P * dy + py), vbBlack
        ObjPic.PSet (mx + P * dx + px, my + P * dy + py), vbBlack
         
    Next P
         
Exit Sub

ErrorProcedure:

    Resume AbortProcedure
     
AbortProcedure:
     
End Sub

Private Sub DisponiAttendi()

    On Error Resume Next
    
    pctAttesa.Move TxMonitor.Left + (TxMonitor.Width - pctAttesa.Width) / 2, TxMonitor.Top + (TxMonitor.Height - pctAttesa.Height) / 2

End Sub

Private Sub AggiungiWatch(Riga As Long)

    On Error Resume Next

    FrmWatch.Codice = TxMonitor.Text
    
    If Riga >= 1 Then
        FrmWatch.Riga = TxMonitor.GetRow(Riga)
    Else
        FrmWatch.Riga = ""
    End If
    
    FrmWatch.Show 1
    
    If FrmWatch.Cancel = False Then
    
        ComandoWatch FrmWatch.Espressione
    
    End If
    
    Unload FrmWatch
    Set FrmWatch = Nothing
          
End Sub
