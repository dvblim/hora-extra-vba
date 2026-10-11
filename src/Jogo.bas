Attribute VB_Name = "Jogo"
'==============================================================================
' Jogo: as regras. Nao toca em planilha: so guarda o estado e avanca turnos.
'
' O jogo e em turnos. Cada passo, espera ou tentativa de abrir algo gasta um
' turno; a pilha da lanterna cai e, depois que a luz volta, a criatura anda.
'
' Ordem da fuga:
'   bilhete na sua sala -> chave amarela no RH -> fusivel na copa ->
'   quadro de luz no corredor (a criatura acorda) -> arquivo ->
'   codigo do cofre (soma do relatorio) -> cracha -> catraca da recepcao
'==============================================================================
Option Explicit

Public Const CODIGO_COFRE As String = "1266"   ' 312 + 458 + 207 + 289, do bilhete 3
Public Const RAIO_LANTERNA As Long = 6
Public Const PILHA_INICIAL As Long = 110
Public Const PILHA_CARGA As Long = 70
Public Const PILHA_MAXIMA As Long = 200

' Fim de jogo.
Public Const JOGANDO As Long = 0
Public Const PERDEU As Long = 1
Public Const VENCEU As Long = 2

' Estado da criatura.
Public Const C_RONDA As Long = 0
Public Const C_PERSEGUE As Long = 1
Public Const C_PROCURA As Long = 2

Public JogadorX As Long, JogadorY As Long
Public Pilha As Long
Public Turnos As Long
Public TemChave As Boolean
Public TemFusivel As Boolean
Public TemCracha As Boolean
Public Energia As Boolean
Public CofreAberto As Boolean
Public Escondido As Boolean
Public Fim As Long

Public CriaX As Long, CriaY As Long
Public CriaturaAcordada As Boolean
Public CriaturaEstado As Long
Public CriaturaLigada As Boolean          ' os testes desligam para checar so os enigmas

Public Mensagem As String                 ' o que acabou de acontecer
Public Leitura As String                  ' texto do bilhete sob o jogador
Public Aviso As String                    ' sons da criatura
Public PedirCodigo As Boolean             ' a tela deve perguntar o codigo do cofre

Public Luz(0 To 43, 0 To 23) As Long      ' 0 = escuro, 1 a 3 = intensidade
Public Visto(0 To 43, 0 To 23) As Boolean ' ja foi iluminado alguma vez

Private mUltimoX As Long, mUltimoY As Long     ' onde a criatura viu o jogador pela ultima vez
Private mProcura As Long                       ' turnos que ainda fica procurando
Private mPonto As Long                         ' ponto atual da ronda
Private mFila(0 To 1055) As Long               ' fila da busca em largura (44 x 24 casas)
Private mOrigem(0 To 43, 0 To 23) As Long

Public Sub NovoJogo()
    Dim x As Long, y As Long

    Mapa.Carregar
    JogadorX = InicioX: JogadorY = InicioY
    CriaX = CriaturaX: CriaY = CriaturaY
    Pilha = PILHA_INICIAL
    Turnos = 0
    TemChave = False: TemFusivel = False: TemCracha = False
    Energia = False: CofreAberto = False: Escondido = False
    Fim = JOGANDO
    CriaturaAcordada = False
    CriaturaEstado = C_RONDA
    CriaturaLigada = True
    mProcura = 0: mPonto = 0
    PedirCodigo = False
    Leitura = "": Aviso = ""
    Mensagem = Ac("A luz caiu. Voc\^e est\'a sozinho no andar, com uma lanterna e pouca pilha.")

    For y = 0 To MAPA_H - 1
        For x = 0 To MAPA_W - 1
            Luz(x, y) = 0
            Visto(x, y) = False
        Next x
    Next y
    Iluminar
End Sub

' O que o jogador deveria estar tentando fazer agora.
Public Function Objetivo() As String
    If TemCracha Then
        Objetivo = Ac("Chegue \`a catraca da recep\,c\~ao.")
    ElseIf Energia Then
        Objetivo = Ac("Abra o cofre do arquivo. N\~ao deixe aquilo te alcan\,car.")
    ElseIf TemFusivel Then
        Objetivo = Ac("Leve o fus\'ivel ao quadro de luz do corredor.")
    ElseIf TemChave Then
        Objetivo = Ac("Abra a copa com a chave amarela.")
    Else
        Objetivo = Ac("Descubra como religar a luz.")
    End If
End Function

' --- Acoes do jogador --------------------------------------------------------------

Public Sub Mover(ByVal dx As Long, ByVal dy As Long)
    Dim nx As Long, ny As Long

    If Fim <> JOGANDO Then Exit Sub
    PedirCodigo = False
    nx = JogadorX + dx
    ny = JogadorY + dy
    If Not Dentro(nx, ny) Then Exit Sub

    Select Case Piso(nx, ny)
        Case P_PAREDE, P_VAZIO, P_MOVEL
            Exit Sub                                  ' bateu: nao gasta turno

        Case P_PORTA_AMARELA
            If TemChave Then
                Piso(nx, ny) = P_PORTA
                Mensagem = Ac("A chave amarela gira. A porta da copa abre.")
            Else
                Mensagem = Ac("Trancada. A fechadura \'e amarela.")
            End If

        Case P_PORTA_ELETRICA
            Mensagem = Ac("Porta eletr\^onica. Sem energia, n\~ao abre.")

        Case P_QUADRO
            If Energia Then
                Mensagem = Ac("O quadro de luz zumbe baixinho.")
            ElseIf TemFusivel Then
                LigarEnergia
            Else
                Mensagem = Ac("Quadro de luz. Falta um fus\'ivel.")
            End If

        Case P_COFRE
            If CofreAberto Then
                Mensagem = Ac("O cofre est\'a aberto e vazio.")
                Exit Sub
            End If
            Mensagem = Ac("O cofre pede um c\'odigo de 4 d\'igitos.")
            PedirCodigo = True
            Exit Sub                                  ' o turno passa em DigitarCodigo

        Case P_SAIDA
            If TemCracha Then
                Fim = VENCEU
                Mensagem = Ac("A catraca libera. Voc\^e saiu.")
                Exit Sub
            End If
            Mensagem = Ac("A catraca pede um crach\'a.")

        Case Else                                     ' piso, porta aberta ou armario
            JogadorX = nx
            JogadorY = ny
            Mensagem = ""
            Leitura = ""
            Entrar
    End Select

    PassarTurno
End Sub

' Fica parado um turno (por exemplo, escondido no armario).
Public Sub Esperar()
    If Fim <> JOGANDO Then Exit Sub
    PedirCodigo = False
    Mensagem = ""
    PassarTurno
End Sub

Public Sub DigitarCodigo(ByVal codigo As String)
    If Fim <> JOGANDO Then Exit Sub
    PedirCodigo = False
    If Trim$(codigo) = CODIGO_COFRE Then
        CofreAberto = True
        TemCracha = True
        Mensagem = Ac("O cofre abre. Dentro, um crach\'a de visitante.")
    Else
        Mensagem = Ac("C\'odigo errado. O teclado apita alto demais.")
    End If
    PassarTurno
End Sub

' O jogador acabou de pisar em uma casa: esconde, pega ou le o que houver.
Private Sub Entrar()
    Escondido = False
    If Piso(JogadorX, JogadorY) = P_ARMARIO Then
        If CriaturaEstado = C_PERSEGUE Then
            Mensagem = Ac("Voc\^e entra no arm\'ario, mas aquilo viu.")
        Else
            Escondido = True
            Mensagem = Ac("Voc\^e se fecha no arm\'ario e prende a respira\,c\~ao.")
        End If
    End If

    Select Case Objeto(JogadorX, JogadorY)
        Case O_CHAVE
            TemChave = True
            Objeto(JogadorX, JogadorY) = O_NADA
            Mensagem = Ac("Voc\^e pegou uma chave amarela.")
        Case O_FUSIVEL
            TemFusivel = True
            Objeto(JogadorX, JogadorY) = O_NADA
            Mensagem = Ac("Voc\^e pegou um fus\'ivel.")
        Case O_PILHA
            Pilha = Pilha + PILHA_CARGA
            If Pilha > PILHA_MAXIMA Then Pilha = PILHA_MAXIMA
            Objeto(JogadorX, JogadorY) = O_NADA
            Mensagem = Ac("Pilhas novas. A lanterna firma.")
        Case O_BILHETE_1, O_BILHETE_2, O_BILHETE_3
            Leitura = TextoDoBilhete(Objeto(JogadorX, JogadorY))
            Mensagem = Ac("H\'a um papel aqui.")
    End Select
End Sub

Public Function TextoDoBilhete(ByVal qual As Long) As String
    Select Case qual
        Case O_BILHETE_1
            TextoDoBilhete = Ac("RECADO DA MANUTEN\,C\~AO" & vbLf & vbLf & _
                "O fus\'ivel do andar queimou de novo. Deixei um reserva na copa, na prateleira." & vbLf & _
                "A copa fica trancada depois das 18h. A chave amarela est\'a com o RH." & vbLf & vbLf & _
                "N\~ao fique at\'e tarde.")
        Case O_BILHETE_2
            TextoDoBilhete = Ac("ATA DA REUNI\~AO" & vbLf & vbLf & _
                "1. O crach\'a de visitante fica guardado no cofre do arquivo." & vbLf & _
                "2. O c\'odigo do cofre passa a ser o TOTAL de horas extras do m\^es." & vbLf & _
                "3. O relat\'orio impresso est\'a no arquivo." & vbLf & vbLf & _
                "4. Ningu\'em soube explicar os passos no corredor.")
        Case O_BILHETE_3
            TextoDoBilhete = Ac("HORAS EXTRAS \- OUTUBRO" & vbLf & vbLf & _
                "Ana ........ 312" & vbLf & _
                "Bruno ...... 458" & vbLf & _
                "Carla ...... 207" & vbLf & _
                "Davi ....... 289" & vbLf & _
                "TOTAL ...... #REF!" & vbLf & vbLf & _
                "(a f\'ormula quebrou; some na m\~ao)")
    End Select
End Function

Private Sub LigarEnergia()
    Dim x As Long, y As Long

    Energia = True
    TemFusivel = False
    For y = 0 To MAPA_H - 1
        For x = 0 To MAPA_W - 1
            If Piso(x, y) = P_PORTA_ELETRICA Then Piso(x, y) = P_PORTA
        Next x
    Next y
    CriaturaAcordada = True
    Mensagem = Ac("As luzes piscam e a porta do arquivo destrava. Em algum lugar do andar, algo se levanta.")
End Sub

Private Sub PassarTurno()
    Turnos = Turnos + 1
    If Pilha > 0 Then Pilha = Pilha - 1
    If Pilha = 0 And Mensagem = "" Then Mensagem = Ac("A lanterna morreu. Voc\^e mal enxerga a pr\'opria m\~ao.")
    MoverCriatura
    Iluminar
    Escutar
End Sub

' --- Luz e visao ---------------------------------------------------------------------

' Recalcula o que a lanterna alcanca. Parede e porta trancada barram a luz.
Public Sub Iluminar()
    Dim x As Long, y As Long
    Dim raio As Long
    Dim d2 As Long

    For y = 0 To MAPA_H - 1
        For x = 0 To MAPA_W - 1
            Luz(x, y) = 0
        Next x
    Next y

    If Pilha > 0 Then raio = RAIO_LANTERNA Else raio = 1
    If Escondido Then raio = 1

    For y = JogadorY - raio To JogadorY + raio
        For x = JogadorX - raio To JogadorX + raio
            If Dentro(x, y) Then
                d2 = (x - JogadorX) * (x - JogadorX) + (y - JogadorY) * (y - JogadorY)
                If d2 <= raio * raio + 1 Then
                    If LinhaLivre(JogadorX, JogadorY, x, y) Then
                        If d2 <= 5 Then
                            Luz(x, y) = 3
                        ElseIf d2 <= 17 Then
                            Luz(x, y) = 2
                        Else
                            Luz(x, y) = 1
                        End If
                        Visto(x, y) = True
                    End If
                End If
            End If
        Next x
    Next y
End Sub

' Ha linha reta sem obstaculo entre as duas casas? (algoritmo de Bresenham)
' As pontas nao contam: uma parede pode ser vista, so nao se ve atraves dela.
Public Function LinhaLivre(ByVal x0 As Long, ByVal y0 As Long, ByVal x1 As Long, ByVal y1 As Long) As Boolean
    Dim dx As Long, dy As Long, sx As Long, sy As Long
    Dim erro As Long, e2 As Long

    dx = Abs(x1 - x0): dy = -Abs(y1 - y0)
    If x0 < x1 Then sx = 1 Else sx = -1
    If y0 < y1 Then sy = 1 Else sy = -1
    erro = dx + dy

    Do While Not (x0 = x1 And y0 = y1)
        e2 = 2 * erro
        If e2 >= dy Then
            erro = erro + dy
            x0 = x0 + sx
        End If
        If e2 <= dx Then
            erro = erro + dx
            y0 = y0 + sy
        End If
        If x0 = x1 And y0 = y1 Then Exit Do
        If Opaco(x0, y0) Then Exit Function
    Loop
    LinhaLivre = True
End Function

' --- A criatura ------------------------------------------------------------------------

Private Sub MoverCriatura()
    Dim passos As Long
    Dim i As Long
    Dim ax As Long, ay As Long

    If Not CriaturaLigada Or Not CriaturaAcordada Or Fim <> JOGANDO Then Exit Sub

    If VeJogador() Then
        CriaturaEstado = C_PERSEGUE
        mUltimoX = JogadorX: mUltimoY = JogadorY
    End If

    ' Em ronda anda devagar; perseguindo acompanha o jogador e as vezes da um bote.
    Select Case CriaturaEstado
        Case C_RONDA:    If Turnos Mod 2 = 0 Then passos = 1
        Case C_PROCURA:  passos = 1
        Case C_PERSEGUE: If Turnos Mod 4 = 0 Then passos = 2 Else passos = 1
    End Select

    For i = 1 To passos
        If Pegou() Then Exit For

        Select Case CriaturaEstado
            Case C_PERSEGUE
                ax = mUltimoX: ay = mUltimoY
            Case C_PROCURA
                ax = mUltimoX: ay = mUltimoY
                mProcura = mProcura - 1
                If mProcura <= 0 Then CriaturaEstado = C_RONDA
            Case Else
                PontoDeRonda mPonto, ax, ay
                If CriaX = ax And CriaY = ay Then
                    mPonto = (mPonto + 1) Mod 6
                    PontoDeRonda mPonto, ax, ay
                End If
        End Select

        PassoPara ax, ay

        If VeJogador() Then
            CriaturaEstado = C_PERSEGUE
            mUltimoX = JogadorX: mUltimoY = JogadorY
        ElseIf CriaturaEstado = C_PERSEGUE Then
            ' Chegou onde viu o jogador pela ultima vez, ou nao tem como chegar: procura por perto.
            If Abs(CriaX - mUltimoX) + Abs(CriaY - mUltimoY) <= 1 Or Not LivreParaCriatura(mUltimoX, mUltimoY) Then
                CriaturaEstado = C_PROCURA
                mProcura = 8
            End If
        End If
    Next i

    If Pegou() Then
        Fim = PERDEU
        Mensagem = Ac("Uma m\~ao fria fecha no seu ombro.")
    End If
End Sub

' A criatura enxerga o jogador? Precisa de linha reta livre, ate 9 casas.
Private Function VeJogador() As Boolean
    Dim d2 As Long

    If Escondido Then Exit Function
    d2 = (CriaX - JogadorX) * (CriaX - JogadorX) + (CriaY - JogadorY) * (CriaY - JogadorY)
    If d2 <= 81 Then VeJogador = LinhaLivre(CriaX, CriaY, JogadorX, JogadorY)
End Function

' Encostou no jogador (na mesma casa ou ao lado)? Escondido no armario nao conta.
Private Function Pegou() As Boolean
    If Escondido Then Exit Function
    Pegou = (Abs(CriaX - JogadorX) + Abs(CriaY - JogadorY) <= 1)
End Function

Private Sub PontoDeRonda(ByVal qual As Long, ByRef x As Long, ByRef y As Long)
    Select Case qual
        Case 0: x = 40: y = 9
        Case 1: x = 36: y = 15
        Case 2: x = 21: y = 9
        Case 3: x = 14: y = 13
        Case 4: x = 3: y = 9
        Case Else: x = 25: y = 14
    End Select
End Sub

' Da um passo na direcao do alvo, pelo caminho mais curto (busca em largura).
Private Sub PassoPara(ByVal ax As Long, ByVal ay As Long)
    Dim nx As Long, ny As Long

    If ProximoPasso(CriaX, CriaY, ax, ay, True, nx, ny) Then
        CriaX = nx
        CriaY = ny
    End If
End Sub

' Primeira casa do caminho mais curto de (sx, sy) ate (tx, ty).
' paraCriatura decide por onde se pode passar. Devolve False se nao ha caminho.
' Se o alvo nao e pisavel (movel, armario), o caminho vai ate uma casa vizinha.
Public Function ProximoPasso(ByVal sx As Long, ByVal sy As Long, ByVal tx As Long, ByVal ty As Long, _
                             ByVal paraCriatura As Boolean, ByRef nx As Long, ByRef ny As Long) As Boolean
    Dim cabeca As Long, cauda As Long
    Dim atual As Long, x As Long, y As Long, vx As Long, vy As Long
    Dim k As Long
    Dim achou As Boolean

    If sx = tx And sy = ty Then Exit Function

    For y = 0 To MAPA_H - 1
        For x = 0 To MAPA_W - 1
            mOrigem(x, y) = -1
        Next x
    Next y

    mFila(0) = sy * MAPA_W + sx
    mOrigem(sx, sy) = sy * MAPA_W + sx
    cauda = 1

    Do While cabeca < cauda And Not achou
        atual = mFila(cabeca)
        cabeca = cabeca + 1
        x = atual Mod MAPA_W
        y = atual \ MAPA_W
        For k = 0 To 3
            vx = x + Choose(k + 1, 1, -1, 0, 0)
            vy = y + Choose(k + 1, 0, 0, 1, -1)
            If Dentro(vx, vy) Then
                If mOrigem(vx, vy) = -1 Then
                    If vx = tx And vy = ty Then
                        mOrigem(vx, vy) = atual
                        achou = True
                        Exit For
                    ElseIf Passavel(vx, vy, paraCriatura) Then
                        mOrigem(vx, vy) = atual
                        mFila(cauda) = vy * MAPA_W + vx
                        cauda = cauda + 1
                    End If
                End If
            End If
        Next k
    Loop
    If Not achou Then Exit Function

    ' Volta do alvo ate a casa seguinte a origem.
    x = tx: y = ty
    Do While mOrigem(x, y) <> sy * MAPA_W + sx
        atual = mOrigem(x, y)
        x = atual Mod MAPA_W
        y = atual \ MAPA_W
    Loop
    If Not Passavel(x, y, paraCriatura) Then Exit Function    ' o alvo esta colado e nao e pisavel
    nx = x: ny = y
    ProximoPasso = True
End Function

Private Function Passavel(ByVal x As Long, ByVal y As Long, ByVal paraCriatura As Boolean) As Boolean
    If paraCriatura Then
        Passavel = LivreParaCriatura(x, y)
    Else
        Passavel = (Piso(x, y) = P_PISO Or Piso(x, y) = P_PORTA Or Piso(x, y) = P_ARMARIO)
    End If
End Function

' Sons: avisa quando a criatura esta perto mas fora da luz.
Private Sub Escutar()
    Dim d As Long

    Aviso = ""
    If Not CriaturaLigada Or Not CriaturaAcordada Or Fim <> JOGANDO Then Exit Sub
    If Luz(CriaX, CriaY) > 0 Then
        Aviso = Ac("EST\'A NA SUA FRENTE.")
        Exit Sub
    End If
    d = Abs(CriaX - JogadorX) + Abs(CriaY - JogadorY)
    If d <= 5 Then
        Aviso = Ac("Respira\,c\~ao. Muito perto.")
    ElseIf d <= 10 Then
        Aviso = Ac("Passos arrastados, perto.")
    ElseIf d <= 18 Then
        Aviso = Ac("Passos ao longe.")
    End If
End Sub
