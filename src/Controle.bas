Attribute VB_Name = "Controle"
'==============================================================================
' Controle: liga o teclado ao jogo.
'
'   Jogar  - monta a tela, comeca uma partida e captura as teclas
'   Sair   - devolve as teclas ao Excel
'
' Nao ha laco de jogo nem chamada ao Windows: cada tecla e ligada a uma macro
' com Application.OnKey, e cada macro avanca um turno e redesenha a tela.
' Enquanto o jogo esta ligado, as letras W A S D E R nao digitam nas celulas;
' Esc (ou trocar de aba) devolve o teclado.
'==============================================================================
Option Explicit

Private mLigado As Boolean

Public Sub Jogar()
    Tela.Preparar
    Jogo.NovoJogo
    Teclas True
    Tela.Desenhar
End Sub

Public Sub Sair()
    Teclas False
End Sub

Public Sub Reiniciar()
    If Not Pronto() Then Exit Sub
    Jogo.NovoJogo
    Tela.Desenhar
End Sub

' Liga ou desliga a captura do teclado.
Private Sub Teclas(ByVal ligar As Boolean)
    Dim nomes As Variant, macros As Variant
    Dim i As Long

    nomes = Array("{UP}", "{DOWN}", "{LEFT}", "{RIGHT}", "w", "s", "a", "d", "e", "r", "{ESC}")
    macros = Array("TeclaCima", "TeclaBaixo", "TeclaEsquerda", "TeclaDireita", _
                   "TeclaCima", "TeclaBaixo", "TeclaEsquerda", "TeclaDireita", _
                   "TeclaEsperar", "Reiniciar", "Sair")

    On Error Resume Next
    For i = LBound(nomes) To UBound(nomes)
        If ligar Then
            Application.OnKey nomes(i), macros(i)
        Else
            Application.OnKey nomes(i)
        End If
    Next i
    On Error GoTo 0
    mLigado = ligar
End Sub

' O jogo so responde com a aba dele na frente; fora dela, solta o teclado.
Private Function Pronto() As Boolean
    If Not mLigado Then Exit Function
    If Not Tela.EmFoco() Then
        Teclas False
        Exit Function
    End If
    Pronto = True
End Function

Public Sub TeclaCima()
    Passo 0, -1
End Sub

Public Sub TeclaBaixo()
    Passo 0, 1
End Sub

Public Sub TeclaEsquerda()
    Passo -1, 0
End Sub

Public Sub TeclaDireita()
    Passo 1, 0
End Sub

Public Sub TeclaEsperar()
    If Not Pronto() Then Exit Sub
    Jogo.Esperar
    DepoisDoTurno
End Sub

Private Sub Passo(ByVal dx As Long, ByVal dy As Long)
    Dim resposta As Variant

    If Not Pronto() Then Exit Sub
    Jogo.Mover dx, dy

    If Jogo.PedirCodigo Then
        Tela.Desenhar
        resposta = Application.InputBox(Ac("C\'odigo de 4 d\'igitos:"), "Cofre", Type:=2)
        If VarType(resposta) = vbBoolean Then
            Jogo.PedirCodigo = False            ' cancelou
        Else
            Jogo.DigitarCodigo CStr(resposta)
        End If
    End If

    DepoisDoTurno
End Sub

Private Sub DepoisDoTurno()
    Tela.Desenhar

    If Jogo.Fim = PERDEU Then
        Beep
        MsgBox Ac("Aquilo te alcan\,cou no turno " & Jogo.Turnos & "." & vbCrLf & vbCrLf & "A noite recome\,ca."), vbCritical, "Hora Extra"
        Jogo.NovoJogo
        Tela.Desenhar
    ElseIf Jogo.Fim = VENCEU Then
        Teclas False
        MsgBox Ac("Voc\^e saiu do pr\'edio em " & Jogo.Turnos & " turnos." & vbCrLf & vbCrLf & _
                  "Amanh\~a tem expediente normal."), vbInformation, "Hora Extra"
    End If
End Sub
