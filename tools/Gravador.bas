Attribute VB_Name = "Gravador"
'==============================================================================
' Gravador: joga uma sequencia de movimentos e salva uma imagem por turno.
' Nao faz parte do jogo; serve para montar videos e GIFs do README.
'
' Uso: ponha na pasta da planilha um arquivo replay.txt com uma linha de
' letras e execute GravarReplay. As imagens saem na subpasta "quadros".
'
'   U D L R  andar (cima, baixo, esquerda, direita)
'   E        esperar um turno
'   C        digitar o codigo certo no cofre (logo apos esbarrar nele)
'
' Depois, por exemplo:  ffmpeg -framerate 5 -i quadros/%04d.png video.mp4
'==============================================================================
Option Explicit

Public Sub GravarReplay()
    Dim movimentos As String
    Dim pasta As String
    Dim arq As Integer
    Dim i As Long
    Dim feitos As Long

    arq = FreeFile
    Open ThisWorkbook.Path & "\replay.txt" For Input As #arq
    Line Input #arq, movimentos
    Close #arq

    pasta = ThisWorkbook.Path & "\quadros"
    If Len(Dir$(pasta, vbDirectory)) = 0 Then MkDir pasta
    On Error Resume Next
    Kill pasta & "\*.png"
    Kill pasta & "\fim.txt"
    On Error GoTo 0

    Tela.Preparar
    Jogo.NovoJogo
    Tela.Desenhar
    SalvarQuadro pasta, 0

    For i = 1 To Len(movimentos)
        Select Case Mid$(movimentos, i, 1)
            Case "U": Jogo.Mover 0, -1
            Case "D": Jogo.Mover 0, 1
            Case "L": Jogo.Mover -1, 0
            Case "R": Jogo.Mover 1, 0
            Case "E": Jogo.Esperar
            Case "C": Jogo.DigitarCodigo CODIGO_COFRE
        End Select
        Tela.Desenhar
        SalvarQuadro pasta, i
        feitos = i
        DoEvents
        If Jogo.Fim <> JOGANDO Then Exit For
    Next i

    ' Marca o fim, para quem esta esperando as imagens saber que acabou.
    arq = FreeFile
    Open pasta & "\fim.txt" For Output As #arq
    Print #arq, "quadros=" & (feitos + 1) & ";fim=" & Jogo.Fim & ";turnos=" & Jogo.Turnos
    Close #arq
End Sub

' Copia a area do jogo como imagem e exporta por um grafico temporario.
Private Sub SalvarQuadro(ByVal pasta As String, ByVal numero As Long)
    Dim area As Range
    Dim grafico As ChartObject
    Dim tentativa As Long

    Set area = Tela.AreaVisivel()
    If area Is Nothing Then Exit Sub

    ' A area de transferencia as vezes esta ocupada; tenta de novo.
    On Error Resume Next
    For tentativa = 1 To 5
        Err.Clear
        area.CopyPicture xlScreen, xlBitmap
        If Err.Number = 0 Then Exit For
        DoEvents
    Next tentativa
    On Error GoTo 0

    Set grafico = area.Worksheet.ChartObjects.Add(area.Left, area.Top + area.Height + 60, area.Width, area.Height)
    grafico.Activate
    grafico.Chart.Paste
    grafico.Chart.Export pasta & "\" & Format$(numero, "0000") & ".png", "PNG"
    grafico.Delete
End Sub
