option explicit

Dim Nome
Dim Cognome

Sub Main ()

    '@Watch Frm
    '@Watch Frm.Visible
    dim x
    dim y
    dim z
    dim a
    '@Inside Frm
    '@Inside v
    '@Inside CL
    '@Inside DXML.firstchild
     
    for I=1 to 10

        x=now

    next
     
    redim V(3)

    if Pippo(2,3) = 5 then

        msgbox "OK"

    end if

    Select Case "3"

    case "2", "4", "3"

        msgbox _
        "Eureka"

    end select

     
    Passaggio x

    Passaggio Pippo(2,1)


    V(0)=now
    v(1)="Frugolo"
    V(2)=4.6
    set r = createobject("adodb.recordset")

    a=1

    Nome = "Rudy"
    Cognome = "Calzetti"

    'exit sub

    x = Pippo(7, Pippo(3,500))-37

    if a=1 then

        if Nome="Rudy" then

            if cognome="calzetti" then

                z="OK"

            else

                z="dasdasd"

            end if

        else

            z=""

        end if

    elseif a=1 then

        z="PIPPO"

    end if
     
     
    On error Resume next

    x=1/0

    z=0

    Rem "Pippo"

    with Frm
     
        with .command1
     
            msgbox .Caption

            select case lcase("CamelCase")

            case "camelcase", "lower"
                z=1
                '@Stop

            case else
                z=2

            end select
        
        end with
        
        '@Stop
        
     end with
     
    If Nome="Rudy" then

        msgbox "OK"
        x="Y"

    ElseIf Nome="Virgo" then

        If 1=1 then

            Y=1

        end if

        msgbox "URGH"     
        x="E"

    else

    End if
     
    '@stop

    select case 2

    case 1

    case 2

        x="B"
        y="D"
        z="X"

    case else

        x="C"
        y="D"
        z="X"

    End select
     
    '@End

    x=0

    for I=42 to 1 step -1

        x=x+1

        if x=23 then
            exit for
        end if

    next
     
    msgbox x

    x=0

    do 

        X=x+1

        if x=3 then
            exit do
        end if

    loop  while x<5

    msgbox x
     
End Sub

Function Pippo(t, y)

    Pippo = t + y
     
End Function

Sub Passaggio(x)

    Set x= createobject("ADODB.recordset")

End Sub