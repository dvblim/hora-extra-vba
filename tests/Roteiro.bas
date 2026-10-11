Option VBASupport 1
Option Explicit
' Roteiro de teste, executado no LibreOffice (modo VBA) por tools/testar.py.
' Joga a fuga inteira com passos automaticos, confere cada enigma, testa a
' criatura e grava quadros para a imagem de previa.

Private mArq As Integer
Private mFalhas As Long

Private Sub Nota(ByVal s As String)
    Print #mArq, "#" & Chr$(9) & s
End Sub

Private Sub Conferir(ByVal nome As String, ByVal ok As Boolean)
    If ok Then
        Nota "ok " & nome
    Else
        Nota "FALHA " & nome
        mFalhas = mFalhas + 1
    End If
End Sub

' Anda ate a casa (ou ate encostar nela, se nao for pisavel) e entao esbarra.
Private Function IrAte(ByVal tx As Long, ByVal ty As Long) As Boolean
    Dim nx As Long, ny As Long, n As Long

    Do While n < 500 And Fim = JOGANDO
        If JogadorX = tx And JogadorY = ty Then
            IrAte = True
            Exit Function
        End If
        If Not ProximoPasso(JogadorX, JogadorY, tx, ty, False, nx, ny) Then Exit Do
        Mover nx - JogadorX, ny - JogadorY
        n = n + 1
    Loop
    If Abs(JogadorX - tx) + Abs(JogadorY - ty) = 1 Then
        Mover tx - JogadorX, ty - JogadorY
        IrAte = True
    End If
End Function

Private Function IrAoObjeto(ByVal codigo As Long) As Boolean
    Dim x As Long, y As Long
    If Achar(codigo, True, x, y) Then IrAoObjeto = IrAte(x, y)
End Function

Private Function IrAoPiso(ByVal codigo As Long) As Boolean
    Dim x As Long, y As Long
    If Achar(codigo, False, x, y) Then IrAoPiso = IrAte(x, y)
End Function

Private Sub Quadro(ByVal nome As String)
    Dim x As Long, y As Long
    Dim fundo As Long, simbolo As String, corSimbolo As Long, codigo As Long

    Print #mArq, "Q" & Chr$(9) & nome
    For y = 0 To MAPA_H - 1
        For x = 0 To MAPA_W - 1
            Aparencia x, y, fundo, simbolo, corSimbolo
            If Len(simbolo) > 0 Then codigo = AscW(simbolo) Else codigo = 0
            Print #mArq, "C" & Chr$(9) & x & Chr$(9) & y & Chr$(9) & fundo & Chr$(9) & codigo & Chr$(9) & corSimbolo
        Next x
    Next y
End Sub

Sub Principal()
    Dim x As Long, y As Long, i As Long

    mArq = FreeFile
    Open Environ("JOGO_SAIDA") For Output As #mArq
    On Error GoTo Falha

    ' --- 1. A fuga inteira, sem a criatura ---
    NovoJogo
    CriaturaLigada = False
    Quadro "inicio"

    Conferir "bilhete 1 fala do fusivel", IrAoObjeto(O_BILHETE_1) And InStr(Leitura, "fus") > 0
    IrAoPiso P_PORTA_AMARELA
    Achar P_PORTA_AMARELA, False, x, y
    Conferir "copa trancada sem a chave", Piso(x, y) = P_PORTA_AMARELA
    Conferir "pega a chave amarela", IrAoObjeto(O_CHAVE) And TemChave
    IrAte x, y
    Conferir "chave abre a copa", Piso(x, y) = P_PORTA
    Conferir "pega o fusivel", IrAoObjeto(O_FUSIVEL) And TemFusivel
    IrAoPiso P_PORTA_ELETRICA
    Conferir "arquivo fechado sem energia", Achar(P_PORTA_ELETRICA, False, x, y)
    Conferir "bilhete 2 fala do cofre", IrAoObjeto(O_BILHETE_2) And InStr(Leitura, "cofre") > 0
    IrAoPiso P_QUADRO
    Conferir "fusivel religa a energia", Energia And CriaturaAcordada And Not TemFusivel
    Conferir "energia abre o arquivo", Not Achar(P_PORTA_ELETRICA, False, x, y)
    Conferir "bilhete 3 tem o relatorio", IrAoObjeto(O_BILHETE_3) And InStr(Leitura, "#REF!") > 0
    IrAoPiso P_COFRE
    Conferir "cofre pede o codigo", PedirCodigo
    DigitarCodigo "0000"
    Conferir "codigo errado nao abre", Not CofreAberto And Not TemCracha
    IrAoPiso P_SAIDA
    Conferir "catraca barra sem cracha", Fim = JOGANDO
    IrAoPiso P_COFRE
    DigitarCodigo CODIGO_COFRE
    Conferir "codigo certo entrega o cracha", CofreAberto And TemCracha
    Quadro "arquivo"
    IrAoPiso P_SAIDA
    Conferir "catraca libera com o cracha", Fim = VENCEU
    Nota "turnos=" & Turnos & ";pilha=" & Pilha
    Conferir "a pilha inicial e as do caminho bastam", Turnos < 400

    ' --- 2. A criatura alcanca quem fica parado no corredor ---
    NovoJogo
    IrAoObjeto O_CHAVE
    IrAoPiso P_PORTA_AMARELA
    IrAoObjeto O_FUSIVEL
    IrAoPiso P_QUADRO
    Conferir "criatura dormindo ate a luz voltar", CriaturaAcordada And Fim = JOGANDO
    For i = 1 To 400
        If Fim <> JOGANDO Then Exit For
        Esperar
    Next i
    Conferir "criatura pega quem fica parado", Fim = PERDEU
    Nota "pego_em=" & i

    ' --- 3. Escondido no armario, a criatura passa direto ---
    NovoJogo
    CriaturaAcordada = True
    JogadorX = 2: JogadorY = 13
    Mover -1, 0
    Conferir "entra no armario e fica escondido", Escondido And JogadorX = 1
    For i = 1 To 300
        Esperar
    Next i
    Conferir "criatura nao acha quem esta escondido", Fim = JOGANDO

    ' --- 4. Quadro para a previa: a criatura no fim do facho da lanterna ---
    NovoJogo
    CriaturaLigada = False
    IrAoObjeto O_CHAVE
    IrAoPiso P_PORTA_AMARELA
    IrAoObjeto O_FUSIVEL
    IrAoPiso P_QUADRO
    JogadorX = 26: JogadorY = 9
    CriaX = 31: CriaY = 9
    Iluminar
    Quadro "corredor"

    Nota "falhas=" & mFalhas
    Nota "FIM"
    Close #mArq
    Exit Sub
Falha:
    Nota "ERRO " & Err.Number & " " & Err.Description
    Close #mArq
End Sub
