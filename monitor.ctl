VERSION 5.00
Begin VB.UserControl CtlMonitor 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   KeyPreview      =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin ryMedio.CtlWheel ObjWheel 
      Left            =   4140
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.HScrollBar ScrollOrizzontale 
      Enabled         =   0   'False
      Height          =   255
      LargeChange     =   10
      Left            =   300
      Max             =   0
      SmallChange     =   10
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2910
      Width           =   3735
   End
   Begin VB.VScrollBar ScrollVerticale 
      Enabled         =   0   'False
      Height          =   2535
      Left            =   4380
      Max             =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Width           =   255
   End
   Begin VB.PictureBox pctSchermo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   300
      MousePointer    =   1  'Arrow
      ScaleHeight     =   1695
      ScaleWidth      =   3135
      TabIndex        =   0
      Top             =   600
      Width           =   3135
      Begin VB.Shape ShapeSelettore 
         BackColor       =   &H00000000&
         BorderColor     =   &H00000000&
         BorderStyle     =   3  'Dot
         DrawMode        =   2  'Blackness
         Height          =   345
         Left            =   720
         Top             =   870
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape ShapeStatement 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         DrawMode        =   7  'Invert
         Height          =   345
         Left            =   270
         Top             =   210
         Visible         =   0   'False
         Width           =   195
      End
   End
   Begin VB.Image ImgBreak 
      Height          =   180
      Left            =   3870
      Picture         =   "monitor.ctx":0000
      Top             =   1050
      Width           =   180
   End
End
Attribute VB_Name = "CtlMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public RigaCliccata As Long
Dim ColNumerazione  As Boolean
Dim Dizio           As String
Dim CurrentRow      As Long
Dim CurrentCol      As Long
Dim TopRow          As Long
Dim LeftCol         As Long
Dim SchermoWidth    As Long
Dim SchermoHeight   As Long
Dim MaxRow          As Long
Dim MaxCol          As Long
Dim PrevCol         As Long

Private Type StructMonRow
    Rg   As String
    Bk   As Boolean
End Type

Dim CodeRows()      As StructMonRow

Dim BorderX         As Single
Dim BorderY         As Single

Const ColRem = &H7F00
Const ColId = &H7F
Const ColConst = &H0
Const ColFunct = &HA00000

Dim SospendiRefresh As Boolean

Public Event DblClick(Riga As Long, Numerazione As Boolean)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim PrevSelettoreRiga    As Long

Dim CarNumeri  As Long
Dim LargBordo  As Long

Private Sub pctSchermo_DblClick()

    On Error Resume Next
    
    If RigaCliccata >= 0 Then
        RaiseEvent DblClick(RigaCliccata, ColNumerazione)
    End If

End Sub

Private Sub pctSchermo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo ErrorProcedure
    
    Dim Riga       As String
    
    RigaCliccata = TopRow + Y \ pctSchermo.TextHeight("X")
    
    If RigaCliccata > MaxRow Then
        RigaCliccata = MaxRow
    End If
    
    If Trim(CodeRows(RigaCliccata).Rg) = "" Then
    
        RigaCliccata = -1
        
    Else
    
        Riga = Trim(CodeRows(RigaCliccata).Rg)
        
        If Left(Riga, 1) = "'" Then
            RigaCliccata = -1
        ElseIf LCase(Left(Riga, 4)) = "rem " Then
            RigaCliccata = -1
        ElseIf LCase(Left(Riga, 4)) = "dim " Then
            RigaCliccata = -1
        ElseIf LCase(Left(Riga, 6)) = "watch " Then
            RigaCliccata = -1
        ElseIf LCase(Left(Riga, 7)) = "inside " Then
            RigaCliccata = -1
        Else
            RigaCliccata = RigaCliccata + 1
        End If
    
    End If
    
    ColNumerazione = X \ pctSchermo.TextWidth("X") <= LargBordo
    
    DisponiSelettore Y
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
    Exit Sub
    
ErrorProcedure:
    
    Resume AbortProcedure
    
AbortProcedure:
    
    On Error Resume Next
    
    RigaCliccata = -1
    
End Sub

Private Sub pctSchermo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    DisponiSelettore Y

End Sub

Private Sub pctSchermo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    DisponiSelettore Y
     
End Sub

Private Sub ScrollOrizzontale_Change()

    LeftCol = ScrollOrizzontale.Value
    PosizionaSelettore
    Me.Refresh

End Sub

Private Sub ScrollOrizzontale_GotFocus()

    pctSchermo.SetFocus

End Sub

Private Sub ScrollVerticale_Change()

    TopRow = ScrollVerticale.Value
    Me.Refresh

End Sub

Private Sub ScrollVerticale_GotFocus()

    pctSchermo.SetFocus

End Sub

Private Sub UserControl_EnterFocus()

    PosizionaSelettore
     
End Sub

Private Sub UserControl_ExitFocus()

    On Error GoTo ErrorProcedure
    
    ParseRow CurrentRow, True
                              
Exit Sub

ErrorProcedure:

    Resume AbortProcedure
     
AbortProcedure:

End Sub

Private Sub UserControl_Initialize()

    On Error Resume Next
    
    Dizio = ""
    Dizio = Dizio + ",And,Call,Case,Const,Dim,Watch,Inside,Stop,Do,Each"
    Dizio = Dizio + ",Else,ElseIf,Empty,End,Eqv,Err,Error,Exit,Explicit"
    Dizio = Dizio + ",False,For,Function,Goto,If,Imp,In,Is,Let,Like,Loop"
    Dizio = Dizio + ",New,Next,Not,Nothing,Null,On,Option,Or,Private"
    Dizio = Dizio + ",Public,Rem,Resume,Select,Set,Step,Sub,Then,To,True"
    Dizio = Dizio + ",Until,Wend,While,With,Abs,Asc,Array,Atn,Chr,Cos,Date"
    Dizio = Dizio + ",Erase,Exp,Filter,Fix,Hex,Hour,Int,Join,Left,Len,Log"
    Dizio = Dizio + ",Mid,Minute,Now,Oct,Randomize,Replace,Right,Rnd,Round"
    Dizio = Dizio + ",Sgn,Second,Sin,Space,Split,Sqr,String,Tan,Time,Trim"
    Dizio = Dizio + ",Weekday,Year,Month,Day,CBool,CByte,CCur,CDate,CDbl,CInt"
    Dizio = Dizio + ",CLng,CSng,CStr,LBound,LCase,LTrim,RTrim,UBound,UCase"
    Dizio = Dizio + ",AscB,AscW,ChrB,ChrW,CreateObject,DateAdd,DateDiff,DatePart"
    Dizio = Dizio + ",DateSerial,DateValue,FormatCurrency,FormatDateTime,FormatNumber"
    Dizio = Dizio + ",FormatPercent,GetObject,InputBox,InStr,InStrB,InStrRev"
    Dizio = Dizio + ",IsArray,IsDate,IsEmpty,IsNull,IsNumeric,IsObject,LeftB"
    Dizio = Dizio + ",LenB,LoadPicture,MidB,MsgBox,ReDim,RGB,RightB,StrComp"
    Dizio = Dizio + ",StrReverse,TimeSerial,TimeValue,TypeName,VarType,WeekdayName"
    Dizio = Dizio + ","
    
    MaxRow = 0
    MaxCol = 0
    CurrentRow = 0
    CurrentCol = 0
    PrevCol = 0
    LeftCol = 0
    TopRow = 0
    
    BorderX = 2 * Screen.TwipsPerPixelX
    BorderY = 2 * Screen.TwipsPerPixelY
    ReDim CodeRows(1023)
    ShapeStatement.Width = pctSchermo.Width
    ShapeStatement.Height = pctSchermo.TextHeight("X")
    PosizionaSelettore
    
    ObjWheel.hWndCapture = UserControl.hwnd
    ObjWheel.EnableWheel
     
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo ErrorProcedure
    
    Select Case KeyCode
    
    Case vbKeyLeft
        If LeftCol - 20 > 0 Then
            LeftCol = LeftCol - 20
        Else
            LeftCol = 0
        End If
        Me.Refresh
    
    Case vbKeyRight
        If LeftCol + 20 <= ScrollOrizzontale.Max Then
            LeftCol = LeftCol + 20
        Else
            LeftCol = ScrollOrizzontale.Max
        End If
        Me.Refresh
    
    Case vbKeyUp
        If TopRow - 1 > 0 Then
            TopRow = TopRow - 1
        Else
            TopRow = 0
        End If
        Me.Refresh
    
    Case vbKeyDown
        If TopRow + 1 <= ScrollVerticale.Max Then
            TopRow = TopRow + 1
        Else
            TopRow = ScrollVerticale.Max
        End If
        Me.Refresh
    
    Case vbKeyPageUp
        If TopRow - SchermoHeight + 1 > 0 Then
            TopRow = TopRow - SchermoHeight + 1
        Else
            TopRow = 0
        End If
        Me.Refresh
    
    Case vbKeyPageDown
        If TopRow + SchermoHeight - 1 <= ScrollVerticale.Max Then
            TopRow = TopRow + SchermoHeight - 1
        Else
            TopRow = ScrollVerticale.Max
        End If
        Me.Refresh
    
    Case vbKeyHome
        If (Shift And 2) <> 0 Then
            TopRow = 0
        Else
            LeftCol = 0
        End If
        Me.Refresh
    
    Case vbKeyEnd
        If (Shift And 2) <> 0 Then
            TopRow = ScrollVerticale.Max
        Else
            LeftCol = ScrollOrizzontale.Max
        End If
        Me.Refresh
    
    End Select
    
    Exit Sub
    
ErrorProcedure:
    
    Resume AbortProcedure
    
AbortProcedure:

End Sub

Private Sub UserControl_Resize()

    On Error GoTo ErrorProcedure
     
    ScrollVerticale.Move _
        UserControl.Width - ScrollVerticale.Width - 4 * Screen.TwipsPerPixelX, _
        0, _
        ScrollVerticale.Width, _
        UserControl.Height - ScrollOrizzontale.Height - 4 * Screen.TwipsPerPixelY
                      
    ScrollOrizzontale.Move _
        0, _
        UserControl.Height - ScrollOrizzontale.Height - 4 * Screen.TwipsPerPixelY, _
        UserControl.Width - ScrollVerticale.Width - 4 * Screen.TwipsPerPixelX
     
    pctSchermo.Move 0, 0, ScrollVerticale.Left, ScrollOrizzontale.Top
     
    SchermoWidth = pctSchermo.Width \ pctSchermo.TextWidth("X")
    SettaMaxScrollCol
     
    SchermoHeight = pctSchermo.Height \ pctSchermo.TextHeight("X")
    SetMaxScrollRow
    
    ScrollVerticale.LargeChange = SchermoHeight - 1
    
    Me.Refresh

Exit Sub

ErrorProcedure:

    Resume AbortProcedure
     
AbortProcedure:
     
End Sub

Private Sub WriteRow(Row As Long)

    Dim Parola          As String
    Dim Car             As String
    Dim InizioParola    As Long
    Dim I               As Long
    Dim OnRem           As Boolean
    Dim OnString        As Boolean
    Dim Colore          As Long
    Dim Normale         As String
    Dim ObjId           As String
    Dim Separatore      As String
    
    pctSchermo.CurrentX = BorderX
    pctSchermo.CurrentY = BorderY + (Row - TopRow) * pctSchermo.TextHeight("X")
    
    InizioParola = 0
    
    pctSchermo.ForeColor = ColConst
                        
    If MaxRow > 0 And Row < MaxRow Then
        pctSchermo.Print "  " & Right(Space(CarNumeri) & (Row + 1), CarNumeri) & " ";
    End If

    For I = 1 To Len(CodeRows(Row).Rg)

        Car = Mid(CodeRows(Row).Rg, I, 1)

        If OnRem Then
        
            Parola = Parola + Car
        
        ElseIf OnString Then
        
            If Car = Chr(34) Then
            
                Parola = Parola + Car
                PrintWord Parola, ObjId, OnRem, OnString, Separatore, InizioParola
                OnString = False
                 
            Else
            
                Parola = Parola + Car
            
            End If
        
        Else
        
            If Car = "'" Then
            
                PrintWord Parola, ObjId, OnRem, OnString, Separatore, InizioParola
                OnRem = True
                InizioParola = I
                Parola = "'"
            
            ElseIf Car = Chr(34) Then
            
                PrintWord Parola, ObjId, OnRem, OnString, Separatore, InizioParola
                OnString = True
                InizioParola = I
                Parola = Chr(34)
            
            Else
            
                Select Case Car

                Case "0" To "9", "a" To "z", "A" To "Z", "_"
                    Parola = Parola + Car
                    If InizioParola = 0 Then
                        InizioParola = I
                    End If
                
                Case Else
                    Separatore = Car
                    PrintWord Parola, ObjId, OnRem, OnString, Separatore, InizioParola
                    
                    If I - LeftCol > SchermoWidth + 2 Then
                        Exit For
                    ElseIf I > LeftCol Then
                        pctSchermo.Print Car;
                    End If
                    
                End Select
                 
            End If
        
        End If
    
    Next I

    PrintWord Parola, ObjId, OnRem, OnString, Separatore, InizioParola

    pctSchermo.Print Space(SchermoWidth);

    If MaxRow > 0 Then
        If CodeRows(Row).Bk Then
            pctSchermo.PaintPicture ImgBreak.Picture, BorderX, 2 * BorderY + (Row - TopRow) * pctSchermo.TextHeight("X")
        End If
    End If

End Sub

Public Sub PrintWord( _
    Parola As String, _
    ObjId As String, _
    OnRem As Boolean, _
    OnString As Boolean, _
    Separatore As String, _
    InizioParola As Long _
)

    Dim Colore  As Long
    Dim Normale As String
    
    If InizioParola > 0 Then
    
        If OnRem Then
            pctSchermo.ForeColor = ColRem
        ElseIf OnString Then
            pctSchermo.ForeColor = ColConst
        ElseIf Riconoscimento(Parola, Colore, Normale, ObjId) Then
            pctSchermo.ForeColor = Colore
            Parola = Normale
        ElseIf IsNumeric(Parola) Then
            pctSchermo.ForeColor = ColConst
        Else
            pctSchermo.ForeColor = ColId
        End If
    
        If InizioParola + Len(Parola) > LeftCol Then
            If InizioParola > LeftCol Then
                 pctSchermo.Print Parola;
            Else
                 pctSchermo.Print Mid(Parola, LeftCol - InizioParola + 2);
            End If
        End If
    
        pctSchermo.ForeColor = ColConst
    
        If Separatore = "." Then
            ObjId = Parola
        Else
            ObjId = ""
        End If
    
        Parola = ""
        InizioParola = 0
    
    Else
    
        ObjId = ""
        
        pctSchermo.ForeColor = ColConst
    
    End If

End Sub

Private Sub PosizionaSelettore()

    If 0 <= (CurrentRow - TopRow) And (CurrentRow - TopRow) <= SchermoHeight + 2 Then
     
        If LeftCol = 0 Then
            ShapeStatement.Move BorderX, (CurrentRow - TopRow) * pctSchermo.TextHeight("X") + BorderY, pctSchermo.Width, pctSchermo.TextHeight("X")
        Else
            ShapeStatement.Move -BorderY, (CurrentRow - TopRow) * pctSchermo.TextHeight("X") + BorderY, pctSchermo.Width + 2 * BorderY, pctSchermo.TextHeight("X")
        End If
        
        ShapeStatement.Visible = True
         
    Else
     
        ShapeStatement.Visible = False
     
    End If

    If ScrollOrizzontale.Value <> LeftCol Then
        If ScrollOrizzontale.Max < LeftCol Then
            MaxCol = CurrentCol
            ScrollOrizzontale.Max = LeftCol
        End If
        ScrollOrizzontale.Value = LeftCol
    End If
     
    If ScrollVerticale.Value <> TopRow Then
        If TopRow <= ScrollVerticale.Max Then
            ScrollVerticale.Value = TopRow
        Else
            ScrollVerticale.Value = ScrollVerticale.Max
        End If
    End If
     
End Sub

Public Sub Refresh()

    On Error GoTo ErrorRefresh
    
    Dim R     As Long
    
    ShapeSelettore.Visible = False
    
    If SospendiRefresh = False Then
    
        pctSchermo.Cls
        
        For R = TopRow To TopRow + SchermoHeight
            WriteRow R
        Next R
        
        SospendiRefresh = True
        PosizionaSelettore
        SospendiRefresh = False
        
        pctSchermo.Refresh
    
    End If
     
Exit Sub

ErrorRefresh:

    Resume AbortRefresh
     
AbortRefresh:

    SospendiRefresh = False
     
End Sub

Private Sub SetMaxScrollRow()

    If MaxRow + 2 > SchermoHeight Then
        ScrollVerticale.Max = MaxRow - SchermoHeight + 2
        ScrollVerticale.Enabled = True
    ElseIf ScrollVerticale.Enabled Then
        ScrollVerticale.Max = 0
        ScrollVerticale.Enabled = False
    End If
     
End Sub

Private Sub SettaMaxScrollCol()

    If MaxCol + 2 > SchermoWidth Then
        ScrollOrizzontale.Max = MaxCol - SchermoWidth + 2
        ScrollOrizzontale.Enabled = True
    ElseIf ScrollOrizzontale.Enabled Then
        ScrollOrizzontale.Max = 0
        ScrollOrizzontale.Enabled = False
    End If
    
End Sub

Private Function Riconoscimento(Parola As String, Colore As Long, Normale As String, ObjId As String) As Boolean

    Dim Ret     As Boolean
    Dim Posiz   As Long
    
    Ret = False
    Colore = ColConst
    Normale = Parola
    
    Posiz = InStr(LCase(Dizio), "," + LCase(Parola) + ",")
    
    If Posiz > 0 Then
    
        Ret = True
        Colore = ColFunct
        Normale = Mid(Dizio, Posiz + 1, Len(Parola))
        
    End If
    
     Riconoscimento = Ret
     
End Function

Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0

    On Error GoTo ErrorProcedure
    
    Dim I As Long
    
    ReDim CopiaRighe(MaxRow) As String
    
    If MaxRow = 0 And CodeRows(0).Rg = "" Then
     
        Text = ""
     
    Else
      
        For I = 0 To MaxRow
            CopiaRighe(I) = CodeRows(I).Rg
        Next I
          
        Text = Join(CopiaRighe, vbCrLf)
          
    End If
     
Exit Property

ErrorProcedure:

    Resume AbortProcedure
     
AbortProcedure:

    Text = ""

End Property

Public Property Let Text(ByVal NewValue As String)

    On Error GoTo ErrorProcedure
    
    Dim NewLine    As String
    Dim BuffRows   As Variant
    Dim I          As Long
    
    UserControl_Initialize
    
    If InStr(NewValue, vbCrLf) > 0 Then
        NewLine = vbCrLf
    ElseIf InStr(NewValue, Chr(10)) > 0 Then
        NewLine = Chr(10)
    Else
        NewLine = vbCrLf
    End If
     
    '------------------------
    ' Scomposizione in righe
    '------------------------
    
    If Right(NewValue, Len(NewLine)) <> NewLine Then
         NewValue = NewValue + NewLine
    End If
    
    BuffRows = Split(NewValue, NewLine)
    
    MaxCol = 0
    MaxRow = UBound(BuffRows)
    CarNumeri = Len("" & (MaxRow + 1))
    LargBordo = CarNumeri + 3
    SchermoWidth = pctSchermo.Width \ pctSchermo.TextWidth("X") - LargBordo
    
    ReDim CodeRows(MaxRow + 1024)
    
    For I = 0 To MaxRow
     
        CodeRows(I).Rg = BuffRows(I)
        ParseRow I, False

        If MaxCol < Len(CodeRows(I).Rg) Then
            MaxCol = Len(CodeRows(I).Rg)
        End If
     
    Next I
     
    SetMaxScrollRow
    SettaMaxScrollCol
          
    Me.Refresh
     
Exit Property

ErrorProcedure:

    Resume AbortProcedure
     
AbortProcedure:
     
    On Error Resume Next
    
    SetMaxScrollRow
    SettaMaxScrollCol
         
    Me.Refresh
    
End Property

Private Sub UserControl_Terminate()

    On Error Resume Next
    
    ObjWheel.DisableWheel
    
End Sub

Private Sub ParseRow(Riga As Long, FlagStampa As Boolean)

    Dim Parola         As String
    Dim Car            As String
    Dim PrevCar        As String
    Dim InizioParola   As Long
    Dim I              As Long
    Dim OnRem          As Boolean
    Dim OnString       As Boolean
    Dim Colore         As Long
    Dim Normale        As String
    Dim ObjId          As String
    Dim Separatore     As String
    Dim FlagInizio     As Boolean
    Dim SuccPos        As Long
    Dim OnIf           As Boolean
    Dim OnThen         As Boolean
    Dim OnFor          As Boolean
    Dim OnTo           As Boolean
    Dim OnNext         As Boolean
    Dim OnEnd          As Boolean
    Dim OnExit         As Boolean
    
    InizioParola = 0
    FlagInizio = False
    I = 0
    
    Do
    
        I = I + 1
        
        If I <= Len(CodeRows(Riga).Rg) Then
          
            If I > 1 Then
                PrevCar = Mid(CodeRows(Riga).Rg, I - 1, 1)
            Else
                PrevCar = ""
            End If
               
            Car = Mid(CodeRows(Riga).Rg, I, 1)
            
            If Car < " " Then
               
                Mid(CodeRows(Riga).Rg, I) = " "
                Car = " "
                    
            End If
               
            If OnRem Then
               
                Parola = Parola + Car
                    
            ElseIf OnString Then
               
                If Car = Chr(34) Then
                    
                    Parola = Parola + Car
                    ParseWord Riga, Parola, ObjId, OnRem, OnString, OnIf, OnThen, OnFor, OnTo, OnNext, OnEnd, OnExit, Separatore, InizioParola
                    OnString = False
                         
                Else
                    
                    Parola = Parola + Car
                    
                End If
               
            Else
               
                If Car = "'" Then
                    
                    ParseWord Riga, Parola, ObjId, OnRem, OnString, OnIf, OnThen, OnFor, OnTo, OnNext, OnEnd, OnExit, Separatore, InizioParola
                    OnRem = True
                    InizioParola = I
                    Parola = "'"
                    
                ElseIf Car = Chr(34) Then
                    
                    ParseWord Riga, Parola, ObjId, OnRem, OnString, OnIf, OnThen, OnFor, OnTo, OnNext, OnEnd, OnExit, Separatore, InizioParola
                    OnString = True
                    InizioParola = I
                    Parola = Chr(34)
                    
                Else
                    
                    Select Case Car
                    
                    Case "0" To "9", "a" To "z", "A" To "Z", "_"
                         
                        FlagInizio = True
                        Parola = Parola + Car
                        
                        If InizioParola = 0 Then
                            InizioParola = I
                        End If
                    
                    Case Else
                         
                        Separatore = Car
                        ParseWord Riga, Parola, ObjId, OnRem, OnString, OnIf, OnThen, OnFor, OnTo, OnNext, OnEnd, OnExit, Separatore, InizioParola
                              
                    End Select
                    
                End If
                    
            End If
               
        Else
          
            Exit Do
               
        End If
          
    Loop
     
    ParseWord Riga, Parola, ObjId, OnRem, OnString, OnIf, OnThen, OnFor, OnTo, OnNext, OnEnd, OnExit, Separatore, InizioParola
    
    If FlagStampa Then
        WriteRow Riga
    End If
     
End Sub

Public Sub ParseWord( _
    Riga As Long, _
    Parola As String, _
    ObjId As String, _
    OnRem As Boolean, _
    OnString As Boolean, _
    OnIf As Boolean, _
    OnThen As Boolean, _
    OnFor As Boolean, _
    OnTo As Boolean, _
    OnNext As Boolean, _
    OnEnd As Boolean, _
    OnExit As Boolean, _
    Separatore As String, _
    InizioParola As Long _
)

    Dim Colore  As Long
    Dim Normale As String
    
    If InizioParola > 0 Then
    
        If OnRem Or OnString Then
         
        ElseIf Riconoscimento(Parola, Colore, Normale, ObjId) Then
         
            Parola = Normale
            
            Mid(CodeRows(Riga).Rg, InizioParola) = Parola
            
            If Colore = ColFunct Then
              
                If Parola = "If" Then
                    OnIf = True
                End If
                
                If Parola = "Then" Then
                    OnThen = True
                End If
                
                If Parola = "For" Then
                    OnFor = True
                End If
                
                If Parola = "To" Or Parola = "Each" Then
                    OnTo = True
                End If
                
                If Parola = "Next" Then
                    OnNext = True
                End If
                
                If Parola = "End" Then
                    OnEnd = True
                End If
                
                If Parola = "Exit" Then
                    OnExit = True
                End If
                   
            End If
              
        End If
          
        If Separatore = "." Then
            ObjId = Parola
        Else
            ObjId = ""
        End If
        
        Parola = ""
        InizioParola = 0
     
    Else
     
        ObjId = ""
          
    End If

End Sub

Public Property Get Riga() As Long

    Riga = CurrentRow

End Property

Public Property Let Riga(ByVal NewValue As Long)

    CurrentRow = NewValue

End Property

Public Sub SetTopRow(ByVal NewValue As Long)

    On Error GoTo ErrorProcedure
    
    If NewValue < ScrollVerticale.Max Then
        TopRow = NewValue
    Else
        TopRow = ScrollVerticale.Max
    End If

Exit Sub

ErrorProcedure:

    Resume AbortProcedure
     
AbortProcedure:

End Sub

Private Sub ObjWheel_WheelScroll(Shift As Integer, zDelta As Integer, X As Single, Y As Single)

    On Error GoTo ErrorProcedure
    
    If zDelta > 0 Then
        If TopRow > 5 Then
            TopRow = TopRow - 5
        Else
            TopRow = 0
        End If
    Else
        If TopRow + 5 < ScrollVerticale.Max Then
            TopRow = TopRow + 5
        Else
            TopRow = ScrollVerticale.Max
        End If
    End If
    
    PosizionaSelettore
    Me.Refresh
    
Exit Sub

ErrorProcedure:

    Resume AbortProcedure
     
AbortProcedure:

End Sub

Public Property Get VisibleRows() As Long

    On Error Resume Next
    
    VisibleRows = SchermoHeight

End Property


Private Sub DisponiSelettore(Y As Single)

    On Error Resume Next
    
    Dim Rg    As Long
    
    Rg = Y \ pctSchermo.TextHeight("X")
    
    If PrevSelettoreRiga <> Rg Then
    
        If Rg < MaxRow - TopRow Then
         
            If LeftCol = 0 Then
                ShapeSelettore.Move BorderX, Rg * pctSchermo.TextHeight("X") + BorderY, pctSchermo.Width, pctSchermo.TextHeight("X")
            Else
                ShapeSelettore.Move -BorderY, Rg * pctSchermo.TextHeight("X") + BorderY, pctSchermo.Width + 2 * BorderY, pctSchermo.TextHeight("X")
            End If
            
            
            If ShapeSelettore.Visible = False Then
                ShapeSelettore.Visible = True
            End If
              
        Else
         
            ShapeSelettore.Visible = False
    
        End If
         
    End If
    
    PrevSelettoreRiga = Rg

End Sub

Public Property Get BreakPoint(Riga As Long) As Boolean

    On Error Resume Next
    
    BreakPoint = CodeRows(Riga - 1).Bk

End Property

Public Property Let BreakPoint(Riga As Long, ByVal NewValue As Boolean)

    On Error Resume Next
    
    CodeRows(Riga - 1).Bk = NewValue

End Property

Public Sub ClearBreakPoint()

    On Error Resume Next
    
    Dim I     As Long
    
    For I = 0 To MaxRow
        CodeRows(I).Bk = False
    Next I
    
    Me.Refresh

End Sub

Public Function ListBreakPoint() As String

    On Error Resume Next
    
    Dim I          As Long
    Dim Buffer     As String
    
    For I = 0 To MaxRow
        If CodeRows(I).Bk Then
            Buffer = Buffer & (I + 1) & "|"
        End If
    Next I
    
    If Buffer <> "" Then
        Buffer = "|" & Buffer
    End If
    
    ListBreakPoint = Buffer

End Function

Public Property Get GetRow(Riga As Long) As String

    On Error Resume Next
    
    GetRow = CodeRows(Riga - 1).Rg

End Property

