Attribute VB_Name = "Mapa"
'==============================================================================
' Mapa: a planta do escritorio e o que ha em cada casa.
'
' Duas camadas: Piso(x, y) diz o tipo da casa (parede, porta, armario...) e
' Objeto(x, y) diz o que esta largado no chao (chave, pilha, bilhete...).
' A planta e desenhada em texto, uma linha por fileira:
'
'   #  parede         .  piso            +  porta comum
'   Y  porta amarela  E  porta eletrica  S  saida (catraca)
'   m  movel          A  armario (esconderijo)
'   Q  quadro de luz  C  cofre
'   k  chave          f  fusivel         p  pilha        1 2 3  bilhetes
'   @  inicio do jogador                 X  inicio da criatura
'==============================================================================
Option Explicit

Public Const MAPA_W As Long = 44
Public Const MAPA_H As Long = 24

' Tipos de casa.
Public Const P_VAZIO As Long = 0
Public Const P_PISO As Long = 1
Public Const P_PAREDE As Long = 2
Public Const P_PORTA As Long = 3
Public Const P_PORTA_AMARELA As Long = 4
Public Const P_PORTA_ELETRICA As Long = 5
Public Const P_SAIDA As Long = 6
Public Const P_MOVEL As Long = 7
Public Const P_ARMARIO As Long = 8
Public Const P_QUADRO As Long = 9
Public Const P_COFRE As Long = 10

' Objetos no chao.
Public Const O_NADA As Long = 0
Public Const O_CHAVE As Long = 1
Public Const O_FUSIVEL As Long = 2
Public Const O_PILHA As Long = 3
Public Const O_BILHETE_1 As Long = 11
Public Const O_BILHETE_2 As Long = 12
Public Const O_BILHETE_3 As Long = 13

Public Piso(0 To 43, 0 To 23) As Long
Public Objeto(0 To 43, 0 To 23) As Long

Public InicioX As Long, InicioY As Long            ' jogador
Public CriaturaX As Long, CriaturaY As Long        ' criatura

Private mFileira As Long

Public Sub Carregar()
    mFileira = 0
    Linha "############################################"
    Linha "#......#.......#........#.........#........#"
    Linha "#.@....#.....k.#..mm....#.A.A.A.A.#...C....#"
    Linha "#..mm..#..mm...#.....f..#.........#........#"
    Linha "#..1...#.......#........#.........#..mm....#"
    Linha "#......#...p...#......p.#.........#.....3..#"
    Linha "#......#.......#........#.........#........#"
    Linha "###+######+########Y########+########E######"
    Linha "#..........................................#"
    Linha "#..........................................#"
    Linha "#..........................................#"
    Linha "####+#########+#######Q##+##########+#######"
    Linha "#......#............#.........#............#"
    Linha "#A.....#.........2..#.......p.#..........A.#"
    Linha "#......#............#.........#............#"
    Linha "#......#...mmmmmm...#.........#............#"
    Linha "#......#...mmmmmm...#...mmm...#............#"
    Linha "#.m....#...mmmmmm...#.........#..mmmmmm....#"
    Linha "#.m....#............#.........#............#"
    Linha "#......#............#.........#............#"
    Linha "#....p.#............#....X....#............#"
    Linha "#......#.A..........#.........#.p..........#"
    Linha "#......#............#.........#............#"
    Linha "######################################S#####"
End Sub

Private Sub Linha(ByVal desenho As String)
    Dim x As Long
    Dim c As String

    For x = 0 To MAPA_W - 1
        c = Mid$(desenho, x + 1, 1)
        Piso(x, mFileira) = P_PISO
        Objeto(x, mFileira) = O_NADA
        Select Case c
            Case " ": Piso(x, mFileira) = P_VAZIO
            Case "#": Piso(x, mFileira) = P_PAREDE
            Case "+": Piso(x, mFileira) = P_PORTA
            Case "Y": Piso(x, mFileira) = P_PORTA_AMARELA
            Case "E": Piso(x, mFileira) = P_PORTA_ELETRICA
            Case "S": Piso(x, mFileira) = P_SAIDA
            Case "m": Piso(x, mFileira) = P_MOVEL
            Case "A": Piso(x, mFileira) = P_ARMARIO
            Case "Q": Piso(x, mFileira) = P_QUADRO
            Case "C": Piso(x, mFileira) = P_COFRE
            Case "k": Objeto(x, mFileira) = O_CHAVE
            Case "f": Objeto(x, mFileira) = O_FUSIVEL
            Case "p": Objeto(x, mFileira) = O_PILHA
            Case "1": Objeto(x, mFileira) = O_BILHETE_1
            Case "2": Objeto(x, mFileira) = O_BILHETE_2
            Case "3": Objeto(x, mFileira) = O_BILHETE_3
            Case "@": InicioX = x: InicioY = mFileira
            Case "X": CriaturaX = x: CriaturaY = mFileira
        End Select
    Next x
    mFileira = mFileira + 1
End Sub

Public Function Dentro(ByVal x As Long, ByVal y As Long) As Boolean
    Dentro = (x >= 0 And y >= 0 And x < MAPA_W And y < MAPA_H)
End Function

' A casa bloqueia a visao e a luz?
Public Function Opaco(ByVal x As Long, ByVal y As Long) As Boolean
    If Not Dentro(x, y) Then
        Opaco = True
    Else
        Select Case Piso(x, y)
            Case P_PAREDE, P_VAZIO, P_PORTA_AMARELA, P_PORTA_ELETRICA, P_SAIDA, P_QUADRO, P_COFRE
                Opaco = True
        End Select
    End If
End Function

' A criatura consegue pisar aqui? Ela nao entra em armario nem passa por porta trancada.
Public Function LivreParaCriatura(ByVal x As Long, ByVal y As Long) As Boolean
    If Dentro(x, y) Then LivreParaCriatura = (Piso(x, y) = P_PISO Or Piso(x, y) = P_PORTA)
End Function

' Primeira casa onde o piso (ou o objeto) tem esse codigo. Usada pelos testes.
Public Function Achar(ByVal codigo As Long, ByVal ehObjeto As Boolean, ByRef x As Long, ByRef y As Long) As Boolean
    Dim i As Long, j As Long

    For j = 0 To MAPA_H - 1
        For i = 0 To MAPA_W - 1
            If (ehObjeto And Objeto(i, j) = codigo) Or (Not ehObjeto And Piso(i, j) = codigo) Then
                x = i: y = j
                Achar = True
                Exit Function
            End If
        Next i
    Next j
End Function
