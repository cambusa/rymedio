Attribute VB_Name = "BasDebug"
'************************************************************************
'* Name:            debug.bas                                           *
'* Project:         ryMedio                                             *
'* Version:         1.3                                                 *
'* Description:     Debugger for VBScript                               *
'* Copyright (C):   2015 Rodolfo Calzetti                               *
'*                  License GNU LESSER GENERAL LICENSE Version 3        *
'* Contact:         https://github.com/cambusa                          *
'*                  postmaster@rudyz.net                                *
'************************************************************************

Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long

Public GlobalErrNumber             As Long
Public GlobalErrDescription        As String

Enum EnumTipoRiga

    trNull
    trDim
    trStatement
    trIf
    trIfThen
    trIfThenElse
    trElseIf
    trElse
    trEndIf
    trFor
    trForStep
    trForEach
    trExitFor
    trNext
    trDo
    trDoUntil
    trDoWhile
    trExitDo
    trLoop
    trLoopUntil
    trLoopWhile
    trSelect
    trCase
    trCaseElse
    trEndSelect
    trFunct
    trEndFunct
    trSub
    trEndSub
    trWith
    trEndWith
    trExitFunction
    trExitSub
    trResumeNext
    trGoToZero
    trGoTo
    trStop
    trEnd
    trRem
    trInside

End Enum

Type StructRiga

    Orig           As String
    Pura           As String
    Tipo           As EnumTipoRiga
    
    Inizio         As Long
    Madre          As Long
    Figlia         As Long
    Sorella        As Long
    
    Livello        As Long
    
    Nome           As String
    Espr1          As String
    Espr2          As String
    Espr3          As String
    
    CaseEspr       As Variant
    
    DefaultName    As String
    
    FlagCiclo      As Boolean
    Confronto      As Variant
    Index          As Long
     
End Type

Type StructLivello

    Inizio         As Long
    Sorella        As Long
    DefaultName    As String

End Type

Type StructValutazione

    Espressione    As String
    Valore         As Variant
    FlagInside     As Boolean

End Type

Type StructFunction

    Nome           As String
    Normale        As String
    EntryPoint     As Long
    Argomenti      As String

End Type

Type StructStatoWin

    L    As Single      ' left
    T    As Single      ' top
    W    As Single      ' width
    H    As Single      ' height
    B    As Single      ' posizione barra
    M    As Long        ' statewindows

End Type

Type StructParentesi

    Id        As String
    In        As Long
    Arg       As Long
    Fi        As Long
    Liv       As Long
    MaxComma  As Long
    Comma     As Variant

End Type

Type StructRipristino

    Argomento      As String
    Espressione    As String
    Ripristino     As String
     
End Type

Public Function AnalizzaSorgente(ObjRegEx As RegExp, Directory As String, ByVal Text As String, VettRighe() As StructRiga, VettEval() As StructValutazione, VettFunct() As StructFunction) As Boolean

    On Error GoTo ErrorProcedure
    
    Dim Esito      As Boolean
    Dim MaxRiga    As Long
    Dim M          As MatchCollection
    Dim S          As MatchCollection
    Dim R          As Long
    Dim I          As Long
    Dim Buff       As String
    Dim PosOffset  As Long
    Dim StrLen     As Long
    Dim V          As Variant
    Dim ElencoVal  As String
    
    Dim UltimaRiga As Long
    Dim UltimoTipo As EnumTipoRiga
    
    Dim FlagVuoto  As Boolean
    
    Dim PrevLivello     As Long
    
    Dim CurrDefaultName As String
    
    Dim CurrLivello     As Long
    ReDim VettLivelli(100) As StructLivello
    
    Dim MaxEval         As Long
    ReDim VettEval(100) As StructValutazione
    
    Dim MaxFunct        As Long
    ReDim VettFunct(100) As StructFunction
    
    ElencoVal = "|"
    
    '-------------------------
    ' Sostituzione tabulatori
    '-------------------------
    
    Text = Replace(Text, vbTab, Space(5))
    
    '----------------------------------
    ' Giunzione delle righe continuate
    '----------------------------------
    
    ObjRegEx.Pattern = " _\r\n *"
    Text = ObjRegEx.Replace(Text, " ")
    
    ObjRegEx.Pattern = " _\r *"
    Text = ObjRegEx.Replace(Text, " ")
    
    ObjRegEx.Pattern = " _\n *"
    Text = ObjRegEx.Replace(Text, " ")
    
    '-----------------------------------
    ' Eliminazione prefisso metacomando
    '-----------------------------------
    
    ObjRegEx.Pattern = "((^|\r|\n) *'[@])"
    Set M = ObjRegEx.Execute(Text)
    
    For R = M.Count - 1 To 0 Step -1
         
        Buff = M(R).SubMatches(0)
        I = M(R).FirstIndex
        Text = Left(Text, I) + Left(Buff, Len(Buff) - 2) + Mid(Text, I + Len(Buff) + 1)
         
    Next R
    
    Set M = Nothing
    
    '---------------------------------
    ' Disabilitazione Option Explicit
    '---------------------------------
    
    ObjRegEx.Pattern = "^ *option +explicit"
    Text = ObjRegEx.Replace(Text, "'option explicit")
    
    '-------------------------------------
    ' Sostituzione caratteri di controllo
    '-------------------------------------
    
    ObjRegEx.Pattern = "[\x00-\x09]"
    Text = ObjRegEx.Replace(Text, "?")
    
    ObjRegEx.Pattern = "[\x0B\x0C]"
    Text = ObjRegEx.Replace(Text, "?")
    
    ObjRegEx.Pattern = "[\x0E\x1F]"
    Text = ObjRegEx.Replace(Text, "?")
    
    '---------------
    ' Analisi righe
    '---------------
    
    ObjRegEx.Pattern = "(.*)(\r|\n|$)"
    Set M = ObjRegEx.Execute(Text)
    
    MaxRiga = M.Count
    
    ReDim VettRighe(MaxRiga) As StructRiga
    
    For R = 1 To MaxRiga
    
        Buff = Replace(Replace(M(R - 1).SubMatches(0), vbLf, ""), vbCr, "")
        
        VettRighe(R).Orig = Buff
        
        ObjRegEx.Pattern = Chr(34) + "[^" + Chr(34) + "]*" + Chr(34)
        Set S = ObjRegEx.Execute(Buff)
        
        For I = 0 To S.Count - 1
        
            Mid(Buff, S(I).FirstIndex + 1) = String(S(I).Length, "§")
        
        Next I
        
        PosOffset = InStr(Buff, "'")
        
        If PosOffset > 0 Then
            Buff = Left(Buff, PosOffset - 1)
        End If
        
        VettRighe(R).Pura = RTrim(Buff)
        
        Set S = Nothing
    
    Next R
    
    Set M = Nothing
    
    CurrLivello = 1
    UltimaRiga = 0
    UltimoTipo = trNull
    MaxEval = 0
    CurrDefaultName = ""
    
    For R = 1 To MaxRiga
    
        VettRighe(R).Tipo = trNull
        
        If Trim(VettRighe(R).Pura) <> "" Then
        
            '-----
            ' DIM
            '-----
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^( *dim )(.+)"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trDim
                    
                    PosOffset = M(0).FirstIndex + Len(M(0).SubMatches(0)) + 1
                    StrLen = Len(M(0).SubMatches(1))
                    
                    VettRighe(R).Nome = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))
                    
                    AggiungiValutazione ElencoVal, VettEval(), MaxEval, VettRighe(R).Nome
                    
                End If
                
                Set M = Nothing
        
            End If
        
            '-------
            ' WATCH
            '-------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^( *watch )(.+)"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trDim
                    
                    PosOffset = M(0).FirstIndex + Len(M(0).SubMatches(0)) + 1
                    StrLen = Len(M(0).SubMatches(1))
                    
                    VettRighe(R).Nome = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))
                    
                    AggiungiValutazione ElencoVal, VettEval(), MaxEval, VettRighe(R).Nome
                    
                End If
                
                Set M = Nothing
        
            End If
        
            '--------
            ' INSIDE
            '--------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^( *inside )(.+)"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trInside
                    
                    PosOffset = M(0).FirstIndex + Len(M(0).SubMatches(0)) + 1
                    StrLen = Len(M(0).SubMatches(1))
                    
                    VettRighe(R).Nome = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))
                    
                    MaxEval = MaxEval + 1
                    
                    If MaxEval > UBound(VettEval) Then
                        ReDim Preserve VettEval(MaxEval + 100)
                    End If
                    
                    VettEval(MaxEval).Espressione = VettRighe(R).Nome
                    VettEval(MaxEval).Valore = Empty
                    VettEval(MaxEval).FlagInside = True
                    
                End If
                
                Set M = Nothing
        
            End If
        
            '------
            ' GOTO
            '------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^( *goto )(\d+) *$"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trGoTo
                    
                    PosOffset = M(0).FirstIndex + Len(M(0).SubMatches(0)) + 1
                    StrLen = Len(M(0).SubMatches(1))
                    
                    VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))
                    
                End If
                
                Set M = Nothing
        
            End If
        
            '------
            ' REM
            '------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^( *rem )(.+)"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trRem
                    
                    PosOffset = M(0).FirstIndex + Len(M(0).SubMatches(0)) + 1
                    StrLen = Len(M(0).SubMatches(1))
                    
                    VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))
                    
                End If
                
                Set M = Nothing
        
            End If
        
            '------
            ' STOP
            '------
            
            If VettRighe(R).Tipo = trNull Then
            
                If LCase(Trim(VettRighe(R).Pura)) = "stop" Then
                
                    VettRighe(R).Tipo = trStop
                    
                End If
                 
            End If
        
            '-----
            ' END
            '-----
            
            If VettRighe(R).Tipo = trNull Then
            
                If LCase(Trim(VettRighe(R).Pura)) = "end" Then
                
                    VettRighe(R).Tipo = trEnd
                    
                End If
                 
            End If
        
            '----------
            ' FUNCTION
            '----------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^((private|public|) *function )(.+)[(](.*)[)]"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trFunct
                    
                    PosOffset = M(0).FirstIndex + Len(M(0).SubMatches(0)) + 1
                    StrLen = Len(M(0).SubMatches(2))
                    
                    VettRighe(R).Nome = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))
                    
                    PosOffset = PosOffset + StrLen + 1
                    StrLen = Len(M(0).SubMatches(3))
                    
                    VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))
                    
                    MaxFunct = MaxFunct + 1
                    
                    If MaxFunct > UBound(VettFunct) Then
                        ReDim Preserve VettFunct(MaxFunct + 100)
                    End If
                    
                    VettFunct(MaxFunct).Nome = VettRighe(R).Nome
                    VettFunct(MaxFunct).Normale = LCase(VettRighe(R).Nome)
                    VettFunct(MaxFunct).Argomenti = Replace(VettRighe(R).Espr1, " ", "")
                    VettFunct(MaxFunct).EntryPoint = R
                    
                    If VettFunct(MaxFunct).Argomenti <> "" Then
                    
                        V = Split(VettFunct(MaxFunct).Argomenti, ",")
                        
                        For I = 0 To UBound(V)
                        
                            AggiungiValutazione ElencoVal, VettEval(), MaxEval, (V(I))
                        
                        Next I
                         
                    End If
                    
                End If
                
                Set M = Nothing
                
            End If
        
            '--------------
            ' END FUNCTION
            '--------------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^ *end +function"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trEndFunct
                    
                End If
                
                Set M = Nothing
        
            End If
        
            '-----
            ' SUB
            '-----
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^((private|public|) *sub )(.+)[(](.*)[)]"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trSub
                    
                    PosOffset = M(0).FirstIndex + Len(M(0).SubMatches(0)) + 1
                    StrLen = Len(M(0).SubMatches(2))
                    
                    VettRighe(R).Nome = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))
                    
                    PosOffset = PosOffset + StrLen + 1
                    StrLen = Len(M(0).SubMatches(3))
                    
                    VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))
                    
                    MaxFunct = MaxFunct + 1
                    
                    If MaxFunct > UBound(VettFunct) Then
                        ReDim Preserve VettFunct(MaxFunct + 100)
                    End If
                    
                    VettFunct(MaxFunct).Nome = VettRighe(R).Nome
                    VettFunct(MaxFunct).Normale = LCase(VettRighe(R).Nome)
                    VettFunct(MaxFunct).Argomenti = Replace(VettRighe(R).Espr1, " ", "")
                    VettFunct(MaxFunct).EntryPoint = R
                    
                    If VettFunct(MaxFunct).Argomenti <> "" Then
                    
                        V = Split(VettFunct(MaxFunct).Argomenti, ",")
                        
                        For I = 0 To UBound(V)
                        
                            AggiungiValutazione ElencoVal, VettEval(), MaxEval, (V(I))
                        
                        Next I
                         
                    End If
                    
                End If
                
                Set M = Nothing
                
            End If
        
            '---------
            ' END SUB
            '---------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^ *end +sub"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trEndSub
                    
                End If
                
                Set M = Nothing
        
            End If
        
            '----------
            ' EXIT SUB
            '----------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^ *exit +sub"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trExitSub
                    
                End If
                
                Set M = Nothing
        
            End If
        
            '---------------
            ' EXIT FUNCTION
            '---------------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^ *exit +function"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trExitFunction
                    
                End If
                
                Set M = Nothing
                
            End If
        
            '----------------------
            ' ON ERROR RESUME NEXT
            '----------------------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^ *on +error +resume +next"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trResumeNext
                    
                    VettRighe(R).Espr1 = Trim(VettRighe(R).Pura)
                    
                End If
                
                Set M = Nothing
        
            End If
        
            '-----------------
            ' ON ERROR GOTO 0
            '-----------------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^ *on +error +goto +0"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trGoToZero
                    
                    VettRighe(R).Espr1 = Trim(VettRighe(R).Pura)
                    
                End If
                
                Set M = Nothing
        
            End If
        
            '--------
            ' ELSEIF
            '--------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^( *elseif )(.+) then"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trElseIf
                    
                    PosOffset = M(0).FirstIndex + Len(M(0).SubMatches(0)) + 1
                    StrLen = Len(M(0).SubMatches(1))
                    
                    VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))
                    
                End If
                
                Set M = Nothing
        
            End If
        
            '--------------
            ' IF THEN ELSE
            '--------------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^ *if (.+) then (.+) else (.+)"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trIfThenElse
                    VettRighe(R).Espr1 = Trim(VettRighe(R).Orig)
                    
                End If
                
                Set M = Nothing
                 
            End If
            
            '---------
            ' IF THEN
            '---------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^ *if (.+) then (.+)"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trIfThen
                    VettRighe(R).Espr1 = Trim(VettRighe(R).Orig)
                    
                End If
                
                Set M = Nothing
                 
            End If
        
            '----
            ' IF
            '----
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^( *)if (.+) then"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trIf
                    
                    PosOffset = M(0).FirstIndex + Len(M(0).SubMatches(0)) + 4
                    StrLen = Len(M(0).SubMatches(1))
                    
                    VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))
                    
                End If
                
                Set M = Nothing
        
            End If
        
            '-------------
            ' SELECT CASE
            '-------------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^( *select +case )(.+)"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trSelect
                    
                    PosOffset = M(0).FirstIndex + Len(M(0).SubMatches(0)) + 1
                    StrLen = Len(M(0).SubMatches(1))
                    
                    VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))
                    
                End If
                
                Set M = Nothing
        
            End If
        
            '-----------
            ' CASE ELSE
            '-----------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^ *case +else"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trCaseElse
                    
                End If
                
                Set M = Nothing
        
            End If
        
            '------
            ' CASE
            '------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^( *case )(.+)"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trCase
                    
                    PosOffset = M(0).FirstIndex + Len(M(0).SubMatches(0)) + 1
                    StrLen = Len(M(0).SubMatches(1))
                    
                    VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))
                    
                    VettRighe(R).CaseEspr = AnalizzaCase(VettRighe(R).Espr1)
                    
                End If
                
                Set M = Nothing
        
            End If
        
            '------------
            ' END SELECT
            '------------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^ *end +select"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trEndSelect
                    
                End If
                
                Set M = Nothing
        
            End If
        
            '------
            ' ELSE
            '------
            
            If VettRighe(R).Tipo = trNull Then
            
                If LCase(Trim(VettRighe(R).Pura)) = "else" Then
                
                    VettRighe(R).Tipo = trElse
                    
                End If
                 
            End If
        
            '--------
            ' END IF
            '--------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^ *end +if"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trEndIf
                    
                End If
                
                Set M = Nothing
                
            End If
        
            '----------
            ' END WITH
            '----------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^ *end +with"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trEndWith
                    
                End If
                
                Set M = Nothing
        
            End If
        
            '------
            ' WITH
            '------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^( *with )(.+)"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trWith
                    
                    PosOffset = M(0).FirstIndex + Len(M(0).SubMatches(0)) + 1
                    StrLen = Len(M(0).SubMatches(1))
                    
                    VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))
                    
                    CurrDefaultName = VettRighe(R).Espr1
                    
                End If
                
                Set M = Nothing
        
            End If
        
            '-------------
            ' FOR EACH IN
            '-------------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^( *for +each )(.+) in (.+)"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trForEach
                    
                    PosOffset = M(0).FirstIndex + Len(M(0).SubMatches(0)) + 1
                    StrLen = Len(M(0).SubMatches(1))
                    VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))
                    
                    PosOffset = PosOffset + StrLen + 4
                    StrLen = Len(M(0).SubMatches(2))
                    VettRighe(R).Espr2 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))
                    
                    VettRighe(R).Nome = VettRighe(R).Espr1
                    
                End If
                
                Set M = Nothing
                 
            End If
            
            '-------------
            ' FOR TO STEP
            '-------------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^( *for )(.+) to (.+) step (.+)"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trForStep
                    
                    PosOffset = M(0).FirstIndex + Len(M(0).SubMatches(0)) + 1
                    StrLen = Len(M(0).SubMatches(1))
                    VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))
                    
                    PosOffset = PosOffset + StrLen + 4
                    StrLen = Len(M(0).SubMatches(2))
                    VettRighe(R).Espr2 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))
                    
                    PosOffset = PosOffset + StrLen + 6
                    StrLen = Len(M(0).SubMatches(3))
                    VettRighe(R).Espr3 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))
                    
                    PosOffset = InStr(VettRighe(R).Espr1, "=")
                    If PosOffset > 0 Then
                        VettRighe(R).Nome = Trim(Mid(VettRighe(R).Espr1, 1, PosOffset - 1))
                    End If
                    
                End If
                
                Set M = Nothing
                
            End If
            
            '--------
            ' FOR TO
            '--------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^( *for )(.+) to (.+)"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trFor
                    
                    PosOffset = M(0).FirstIndex + Len(M(0).SubMatches(0)) + 1
                    StrLen = Len(M(0).SubMatches(1))
                    VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))
                    
                    PosOffset = PosOffset + StrLen + 4
                    StrLen = Len(M(0).SubMatches(2))
                    VettRighe(R).Espr2 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))
                    
                    VettRighe(R).Espr3 = "1"
                    
                    PosOffset = InStr(VettRighe(R).Espr1, "=")
                    If PosOffset > 0 Then
                        VettRighe(R).Nome = Trim(Mid(VettRighe(R).Espr1, 1, PosOffset - 1))
                    End If
                    
                End If
                
                Set M = Nothing
                 
            End If
            
            '----------
            ' EXIT FOR
            '----------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^ *exit +for"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trExitFor
                    
                End If
                
                Set M = Nothing
                
            End If
        
            '------
            ' NEXT
            '------
            
            If VettRighe(R).Tipo = trNull Then
            
                If LCase(Trim(VettRighe(R).Pura)) = "next" Then
                
                    VettRighe(R).Tipo = trNext
                    
                End If
                 
            End If
        
            '----------
            ' DO UNTIL
            '----------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^( *do +until )(.+)"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trDoUntil
                    
                    PosOffset = M(0).FirstIndex + Len(M(0).SubMatches(0)) + 1
                    StrLen = Len(M(0).SubMatches(1))
                    
                    VettRighe(R).Espr1 = "Not (" + Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen)) + ")"
                    
                End If
                
                Set M = Nothing
        
            End If
        
            '----------
            ' DO WHILE
            '----------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^( *do +while )(.+)"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trDoWhile
                    
                    PosOffset = M(0).FirstIndex + Len(M(0).SubMatches(0)) + 1
                    StrLen = Len(M(0).SubMatches(1))
                    
                    VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))
                    
                End If
                
                Set M = Nothing
        
            End If
        
            '---------
            ' EXIT DO
            '---------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^ *exit +do"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trExitDo
                    
                End If
                
                Set M = Nothing
        
            End If
        
            '------------
            ' LOOP UNTIL
            '------------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^( *loop +until )(.+)"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trLoopUntil
                    
                    PosOffset = M(0).FirstIndex + Len(M(0).SubMatches(0)) + 1
                    StrLen = Len(M(0).SubMatches(1))
                    
                    VettRighe(R).Espr1 = "Not (" + Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen)) + ")"
                    
                End If
                
                Set M = Nothing
        
            End If
        
            '------------
            ' LOOP WHILE
            '------------
            
            If VettRighe(R).Tipo = trNull Then
            
                ObjRegEx.Pattern = "^( *loop +while )(.+)"
                Set M = ObjRegEx.Execute(VettRighe(R).Pura)
                
                If M.Count = 1 Then
                    
                    VettRighe(R).Tipo = trLoopWhile
                    
                    PosOffset = M(0).FirstIndex + Len(M(0).SubMatches(0)) + 1
                    StrLen = Len(M(0).SubMatches(1))
                    
                    VettRighe(R).Espr1 = Trim(Mid(VettRighe(R).Orig, PosOffset, StrLen))
                    
                End If
                
                Set M = Nothing
        
            End If
        
            '----
            ' DO
            '----
            
            If VettRighe(R).Tipo = trNull Then
            
                If LCase(Trim(VettRighe(R).Pura)) = "do" Then
                    
                    VettRighe(R).Tipo = trDo
                    
                    VettRighe(R).Espr1 = "-1"
                    
                End If
                 
            End If
        
            '------
            ' LOOP
            '------
            
            If VettRighe(R).Tipo = trNull Then
            
                If LCase(Trim(VettRighe(R).Pura)) = "loop" Then
                    
                    VettRighe(R).Tipo = trLoop
                    
                    VettRighe(R).Espr1 = "-1"
                    
                End If
                 
            End If
        
            '----------
            ' CHIAMATA
            '----------
            
            If VettRighe(R).Tipo = trNull Then
            
                VettRighe(R).Tipo = trStatement
                VettRighe(R).Espr1 = Trim(VettRighe(R).Orig)
                 
            End If
            
            '----------------------------
            ' Gestione livello
            '----------------------------
            
            VettRighe(R).Livello = 0
            VettRighe(R).Inizio = 0
            VettRighe(R).Madre = 0
            VettRighe(R).Figlia = 0
            
            FlagVuoto = BloccoVuoto(UltimoTipo, VettRighe(R).Tipo)
            
            Select Case UltimoTipo
            
            Case trCase, trCaseElse, trDo, trDoUntil, trDoWhile, trElse, trElseIf, trFor, trForStep, trForEach, trFunct, trIf, trSelect, trSub, trWith
            
                VettLivelli(CurrLivello).Inizio = UltimaRiga
                
                If Not FlagVuoto Then
                
                    CurrLivello = CurrLivello + 1
                    
                    VettLivelli(CurrLivello).Inizio = R
                    
                    VettRighe(UltimaRiga).Figlia = R
                     
                End If
                      
            End Select
            
            Select Case VettRighe(R).Tipo
        
            Case trCase, trCaseElse, trDo, trDoUntil, trDoWhile, trElse, trElseIf, trFor, trForStep, trForEach, trFunct, trIf, trSelect, trSub, trWith
            
                If CurrDefaultName <> "" Then
                
                    VettLivelli(CurrLivello + 1).DefaultName = VettLivelli(CurrLivello).DefaultName + CurrDefaultName
                    
                    CurrDefaultName = ""
                
                Else
                
                    VettLivelli(CurrLivello + 1).DefaultName = VettLivelli(CurrLivello).DefaultName
                     
                End If
                 
            End Select
                
            Select Case VettRighe(R).Tipo
            
            Case trEndFunct, trEndIf, trEndSelect, trEndSub, trEndWith, trLoop, trLoopUntil, trLoopWhile, trNext
            
                If Not FlagVuoto Then
                
                    CurrLivello = CurrLivello - 1
                     
                End If
                
                If VettRighe(R).Tipo = trEndSelect Then
                    CurrLivello = CurrLivello - 1
                End If
                     
                VettRighe(R).Livello = CurrLivello
                VettRighe(R).Madre = VettLivelli(CurrLivello - 1).Inizio
                
            Case trElse, trElseIf
            
                If Not FlagVuoto Then
                     
                    CurrLivello = CurrLivello - 1
                          
                End If
                
                VettRighe(R).Livello = CurrLivello
                VettRighe(R).Madre = VettLivelli(CurrLivello - 1).Inizio
                
            Case trCase, trCaseElse
            
                If Not FlagVuoto Then
                     
                    If VettRighe(UltimaRiga).Tipo <> trSelect Then
                        CurrLivello = CurrLivello - 1
                    Else
                        VettRighe(UltimaRiga).Figlia = R
                    End If
                          
                End If
                
                VettRighe(R).Livello = CurrLivello
                VettRighe(R).Madre = VettLivelli(CurrLivello - 1).Inizio
                
            Case trDo, trDoUntil, trDoWhile, trFor, trForStep, trForEach, trFunct, trIf, trSelect, trSub, trWith
                
                VettRighe(R).Livello = CurrLivello
                VettRighe(R).Madre = VettLivelli(CurrLivello - 1).Inizio
                VettLivelli(CurrLivello).Inizio = R
                
            Case Else
            
                If UltimaRiga > 0 Then
                
                    If CurrLivello = VettRighe(UltimaRiga).Livello Then
                        VettRighe(R).Madre = VettRighe(UltimaRiga).Madre
                    Else
                        VettRighe(R).Madre = UltimaRiga
                    End If
                    
                    VettRighe(R).Livello = CurrLivello
                     
                Else
                
                    VettRighe(R).Livello = CurrLivello
                    VettLivelli(CurrLivello).Inizio = R
                
                End If
                      
            End Select
            
            VettRighe(R).DefaultName = VettLivelli(CurrLivello).DefaultName
            
            If VettRighe(R).DefaultName <> "" Then
            
                NormalizzaPredefinito VettRighe(R).Espr1, VettRighe(R).DefaultName
                NormalizzaPredefinito VettRighe(R).Espr2, VettRighe(R).DefaultName
                NormalizzaPredefinito VettRighe(R).Espr3, VettRighe(R).DefaultName
                 
            End If
            
            UltimaRiga = R
            UltimoTipo = VettRighe(R).Tipo
            
        End If
         
    Next R
    
    UltimaRiga = 0      ' Precedente riga non vuota
    PrevLivello = 0     ' Livello di UltimaRiga
    CurrLivello = 0     ' Livello di R
    
    For R = MaxRiga To 1 Step -1
    
        If Trim(VettRighe(R).Pura) <> "" Then
        
            CurrLivello = VettRighe(R).Livello
            
            If CurrLivello > PrevLivello Then
            
                '-------------------------------
                ' Indentazione: resetto Sorella
                '-------------------------------
                
                VettLivelli(CurrLivello).Sorella = 0
                 
            End If
            
            '--------------------------------------------------------
            ' Se non siamo sull'ultima riga dello script,
            ' memorizzo in Sorella del precedente livello UltimaRiga
            '--------------------------------------------------------
            
            If PrevLivello > 0 Then
                VettLivelli(PrevLivello).Sorella = UltimaRiga
            End If
            
            '--------------------------------------------------------------------------
            ' Assegno Sorella corrente che ho bufferizziato nella struttura di livello
            '--------------------------------------------------------------------------
            
            VettRighe(R).Sorella = VettLivelli(CurrLivello).Sorella
            
            '-------------------------------------------------------------------------------
            ' Se sono all'inizio di un ciclo FOR o DO, assegno inizio alla NEXT o alla LOOP
            '-------------------------------------------------------------------------------
            
            I = VettRighe(R).Sorella
            
            If I > 0 Then
            
                Select Case VettRighe(I).Tipo
                
                Case trNext, trLoop, trLoopUntil, trLoopWhile
                    VettRighe(I).Inizio = R
                          
                End Select
                 
            End If
            
            '-------------------------------
            ' Memorizzo lo stato precedente
            '-------------------------------
            
            PrevLivello = CurrLivello
            UltimaRiga = R
        
        End If
    
    Next R
    
    'ScriviHTML Directory + "\~debug.htm", VettRighe(), 0, ""
    
    ReDim Preserve VettEval(MaxEval)
    ReDim Preserve VettFunct(MaxFunct)
    
    AnalizzaSorgente = Esito
     
Exit Function

ErrorProcedure:

    GlobalErrDescription = Err.Description
    GlobalErrNumber = Err.Number
    
    Resume AbortProcedure
     
AbortProcedure:

    On Error Resume Next
    
    Set M = Nothing
    Set S = Nothing
    
    ReDim Preserve VettEval(0)
    ReDim Preserve VettFunct(0)
    
    AnalizzaSorgente = False
     
End Function

Private Function BloccoVuoto(PrecTipo As EnumTipoRiga, CurrTipo As EnumTipoRiga) As Boolean

    Dim Esito As Boolean
    
    Esito = False
    
    Select Case PrecTipo
         
    Case trCase, trCaseElse

        Select Case CurrTipo
             
        Case trCase, trCaseElse, trEndSelect
        
            Esito = True
                  
        End Select
    
    Case trDo, trDoUntil, trDoWhile

        Select Case CurrTipo
             
        Case trLoop, trLoopUntil, trLoopWhile
        
            Esito = True
                  
        End Select
    
    Case trElse, trElseIf, trIf

        Select Case CurrTipo
             
        Case trElse, trElseIf, trEndIf
        
            Esito = True
                  
        End Select
    
    Case trFor, trForStep, trForEach

        Select Case CurrTipo
             
        Case trNext
        
            Esito = True
                  
        End Select
    
    Case trFunct, trSub

        Select Case CurrTipo
             
        Case trEndFunct, trEndSub
        
            Esito = True
                  
        End Select
    
    Case trSelect

        Select Case CurrTipo
             
        Case trEndSelect
        
            Esito = True
                  
        End Select
    
    Case trWith

        Select Case CurrTipo
             
        Case trEndWith
        
            Esito = True
                  
        End Select
         
    End Select
    
    BloccoVuoto = Esito

End Function

Public Sub ScriviHTML(PathFile As String, VettRighe() As StructRiga, CurrentRow As Long, Messaggio As String)

    On Error GoTo ErrorProcedure
    
    Dim R               As Long
    Dim Buffer          As String
    Dim NumFile         As Integer
    Dim TestataHTML     As String
    Dim CodaHTML        As String
    Dim StrClass        As String
    
    ReDim VettBuff(UBound(VettRighe))
    
    For R = 1 To UBound(VettRighe)
    
        If R = CurrentRow Then
            StrClass = " class='selected'"
        Else
            StrClass = ""
        End If
        
        If Trim(VettRighe(R).Pura) <> "" Then
            VettBuff(R - 1) = "<tr" + StrClass + "><td>" & R & "</td><td style='padding-left:20px;'>" + Replace(VettRighe(R).Orig, " ", "&nbsp;") + "</td><td style='padding-left:20px;'>" & VettRighe(R).Livello & "</td><td style='padding-left:20px;'>" & VettRighe(R).Madre & "</td><td style='padding-left:20px;'>" & VettRighe(R).Figlia & "</td><td style='padding-left:20px;'>" & VettRighe(R).Sorella & "</td><td style='padding-left:20px;'>" & VettRighe(R).Inizio & "</td><tr>"
        Else
            VettBuff(R - 1) = "<tr><td>" & R & "</td><td colspan='5' style='font-size:4px;'>&nbsp;</td><tr>"
        End If
    
    Next R
    
    TestataHTML = ""
    TestataHTML = TestataHTML + "<html>" + vbCrLf
    TestataHTML = TestataHTML + "<style>" + vbCrLf
    TestataHTML = TestataHTML + ".selected{background-color:yellow;}" + vbCrLf
    TestataHTML = TestataHTML + ".cell{padding-left:20px;}" + vbCrLf
    TestataHTML = TestataHTML + "</style>" + vbCrLf
    TestataHTML = TestataHTML + "<body style='font-family:verdana;font-size:12px;'>" + vbCrLf
    
    If Messaggio <> "" Then
        TestataHTML = TestataHTML + "<div style='background-color:red;color:white;'>" + Messaggio + "</div>" + vbCrLf
    End If
    
    TestataHTML = TestataHTML + "<table style='font-family:courier new;font-size:12px;border-collapse:collapse;'>" + vbCrLf
    TestataHTML = TestataHTML + "<tr><th></th><th></th><th style='padding-left:20px;'>Livello</th><th style='padding-left:20px;'>Madre</th><th style='padding-left:20px;'>Figlia</th><th style='padding-left:20px;'>Sorella</th><th style='padding-left:20px;'>Inizio</th></tr>" + vbCrLf
    
    Buffer = Join(VettBuff, vbCrLf)
    
    CodaHTML = ""
    CodaHTML = CodaHTML + "</table>" + vbCrLf
    CodaHTML = CodaHTML + "</body>" + vbCrLf
    CodaHTML = CodaHTML + "</html>" + vbCrLf
    
    Buffer = TestataHTML + Buffer + CodaHTML
    
    NumFile = FreeFile
    Open PathFile For Output As #NumFile
    Print #NumFile, Buffer
    Close #NumFile
    NumFile = 0
     
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

Public Sub NormalizzaPredefinito(Espressione As String, DefaultName As String)

    Dim I          As Long
    Dim Buffer     As String
    Dim FlagStr    As Boolean
    
    Dim C          As String
    Dim P          As String
    Dim S          As String
    
    Dim L          As Long
    
    If Espressione <> "" Then
    
        If InStr(Espressione, ".") > 0 Then
        
            P = ""
            FlagStr = False
            L = Len(Espressione)
            
            For I = 1 To L
            
                C = Mid(Espressione, I, 1)
                
                If C = Chr(34) Then
                    FlagStr = Not FlagStr
                End If
                
                If FlagStr = False Then
                     
                    If C = "." Then
                    
                        If I < L Then
                            S = Mid(Espressione, I + 1, 1)
                        Else
                            S = ""
                        End If
                        
                        If S < "0" Or "9" < S Then
                        
                            Select Case P
                            
                            Case "", " ", ":", "=", "<", ">", "+", "-", "*", "/", "\", "^", "("
                                C = DefaultName + "."
                            
                            End Select
                        
                        End If
                         
                    End If
                
                End If
                
                Buffer = Buffer + C
                
                P = C
                 
            Next I
            
            Espressione = Buffer
        
        End If
         
    End If

End Sub

Sub PollingAttesa(Ritardo As Single)

    On Error Resume Next
    
    Dim Tempo As Single

    Tempo = Timer

    Do
         
        DoEvents
        Sleep 1
         
    Loop Until Abs(Timer - Tempo) > Ritardo

End Sub

Sub PrimoPiano(Mask As Form, FlagPrimo As Boolean)

    On Error Resume Next
    
    Dim Esito  As Long
    Dim Scelta As Long

    If FlagPrimo Then
        Scelta = -1
    Else
        Scelta = -2
    End If

    Esito = SetWindowPos(Mask.hwnd, Scelta, 0, 0, 0, 0, 3)

End Sub

Public Sub AggiungiValutazione(ElencoVal As String, VettEval() As StructValutazione, MaxEval As Long, Id As String)

    On Error Resume Next
    
    If InStr(1, ElencoVal, "|" & Id & "|", vbTextCompare) = 0 Then
    
        MaxEval = MaxEval + 1
        
        If MaxEval > UBound(VettEval) Then
            ReDim Preserve VettEval(MaxEval + 100)
        End If
        
        VettEval(MaxEval).Espressione = Id
        VettEval(MaxEval).Valore = Empty
        
        ElencoVal = ElencoVal + Id + "|"
         
    End If
    
End Sub

Private Function AnalizzaCase(CaseArgs As String) As Variant

    On Error GoTo ErrorProcedure
    
    Dim V          As Variant
    Dim I          As Long
    Dim K          As String
    Dim FlagStr    As Boolean
    Dim MaxArg     As Long
    Dim PosPrev    As Long
    
    MaxArg = 0
    ReDim V(20) As String
    PosPrev = 0
    FlagStr = False
    
    For I = 1 To Len(CaseArgs)
    
        K = Mid(CaseArgs, I, 1)

        If FlagStr Then

            If K = Chr(34) Then
                FlagStr = False
            End If
             
        ElseIf K = Chr(34) Then
        
            FlagStr = True

        Else

            Select Case UCase(K)

            Case ","
            
                MaxArg = MaxArg + 1
                
                If MaxArg > UBound(V) Then
                    ReDim Preserve V(MaxArg + 20)
                End If
                
                V(MaxArg) = Trim(Mid(CaseArgs, PosPrev + 1, I - PosPrev - 1))
                
                PosPrev = I
                      
            End Select
             
        End If
         
    Next I
    
    MaxArg = MaxArg + 1
    
    If MaxArg > UBound(V) Then
        ReDim Preserve V(MaxArg + 20)
    End If
    
    V(MaxArg) = Trim(Mid(CaseArgs, PosPrev + 1))
    
    ReDim Preserve V(MaxArg)
    
    AnalizzaCase = V
     
Exit Function

ErrorProcedure:

    Resume AbortProcedure
     
AbortProcedure:

    On Error Resume Next
    
    ReDim Preserve V(0)
    
    AnalizzaCase = V
     
End Function

Public Function RisolviOggetto(ObjColl As Object, Indice As Long) As Object

    On Error GoTo ErrorProcedure
    
    Dim CurrInd    As Long
    Dim Oggetto    As Object
    
    CurrInd = 0
    
    For Each Oggetto In ObjColl
    
        CurrInd = CurrInd + 1
        
        If CurrInd >= Indice Then
        
            Exit For
        
        End If
    
    Next
    
    If CurrInd <> Indice Then
        Set Oggetto = Nothing
    End If
    
    Set RisolviOggetto = Oggetto
    Set Oggetto = Nothing

Exit Function

ErrorProcedure:

    Resume AbortProcedure
     
AbortProcedure:

    On Error Resume Next
    
    Set RisolviOggetto = Nothing
    Set Oggetto = Nothing
     
End Function
