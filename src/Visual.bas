Attribute VB_Name = "Visual"
'==============================================================================
' Visual: decide a aparencia de cada casa (cor de fundo, simbolo e cor do
' simbolo) a partir do estado do jogo. Nao desenha nada: quem pinta as
' celulas e o modulo Tela. Separado assim, o visual tambem roda nos testes.
'
' Tres situacoes para uma casa:
'   iluminada agora  -> cor cheia, mais fraca conforme a distancia
'   ja vista         -> um resto de cor, quase preto (memoria do lugar)
'   nunca vista      -> preto
'==============================================================================
Option Explicit

Public Function CorPreta() As Long
    CorPreta = RGB(6, 6, 9)
End Function

' Devolve a aparencia da casa (x, y).
Public Sub Aparencia(ByVal x As Long, ByVal y As Long, ByRef fundo As Long, ByRef simbolo As String, ByRef corSimbolo As Long)
    Dim nivel As Long
    Dim fator As Double

    simbolo = ""
    corSimbolo = RGB(235, 235, 235)
    nivel = Luz(x, y)

    If nivel = 0 Then
        If Visto(x, y) Then
            fundo = Escurecer(CorDoPiso(Piso(x, y)), 0.16)
        Else
            fundo = CorPreta
        End If
        Exit Sub
    End If

    fator = Choose(nivel, 0.38, 0.66, 1#)
    fundo = Escurecer(CorDoPiso(Piso(x, y)), fator)

    ' Simbolo da casa e do que esta no chao.
    Select Case Piso(x, y)
        Case P_ARMARIO: simbolo = ChrW(9618): corSimbolo = Escurecer(RGB(150, 170, 215), fator)
        Case P_QUADRO:  simbolo = ChrW(9889): corSimbolo = RGB(255, 220, 90)
        Case P_COFRE:   simbolo = ChrW(9632): corSimbolo = Escurecer(RGB(210, 210, 220), fator)
        Case P_SAIDA:   simbolo = ChrW(9650): corSimbolo = RGB(170, 255, 190)
    End Select
    Select Case Objeto(x, y)
        Case O_CHAVE:   simbolo = "K": corSimbolo = RGB(255, 215, 60)
        Case O_FUSIVEL: simbolo = "F": corSimbolo = RGB(255, 150, 60)
        Case O_PILHA:   simbolo = "+": corSimbolo = RGB(120, 230, 130)
        Case O_BILHETE_1, O_BILHETE_2, O_BILHETE_3: simbolo = ChrW(8801): corSimbolo = RGB(245, 240, 215)
    End Select

    ' Quem esta na casa.
    If CriaX = x And CriaY = y Then
        simbolo = ChrW(9760)
        corSimbolo = RGB(255, 40, 40)
        If CriaturaAcordada Then fundo = Escurecer(RGB(110, 0, 0), fator)
    End If
    If JogadorX = x And JogadorY = y Then
        simbolo = ChrW(9787)
        If Escondido Then corSimbolo = RGB(120, 140, 190) Else corSimbolo = RGB(255, 245, 200)
    End If
End Sub

Private Function CorDoPiso(ByVal tipo As Long) As Long
    Select Case tipo
        Case P_PISO:           CorDoPiso = RGB(128, 108, 78)
        Case P_PAREDE:         CorDoPiso = RGB(78, 84, 98)
        Case P_PORTA:          CorDoPiso = RGB(150, 96, 44)
        Case P_PORTA_AMARELA:  CorDoPiso = RGB(205, 172, 40)
        Case P_PORTA_ELETRICA: CorDoPiso = RGB(50, 135, 170)
        Case P_SAIDA:          CorDoPiso = RGB(40, 150, 80)
        Case P_MOVEL:          CorDoPiso = RGB(92, 64, 38)
        Case P_ARMARIO:        CorDoPiso = RGB(62, 74, 108)
        Case P_QUADRO:         CorDoPiso = RGB(160, 50, 42)
        Case P_COFRE:          CorDoPiso = RGB(110, 112, 124)
        Case Else:             CorDoPiso = RGB(6, 6, 9)
    End Select
End Function

Private Function Escurecer(ByVal cor As Long, ByVal fator As Double) As Long
    Escurecer = RGB(Int((cor And 255) * fator), Int(((cor \ 256) And 255) * fator), Int(((cor \ 65536) And 255) * fator))
End Function
