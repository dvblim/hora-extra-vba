Attribute VB_Name = "Tela"
'==============================================================================
' Tela: desenha o jogo na planilha. Cada casa do mapa e uma celula quadrada.
'
' Para nao piscar nem ficar lento, a tela guarda o que ja esta pintado e so
' mexe nas celulas que mudaram desde o turno anterior (em geral, so as que
' estao em volta da lanterna).
'
' Disposicao: o mapa ocupa B2 em diante; o painel fica a direita do mapa e o
' papel que o jogador estiver lendo aparece embaixo.
'==============================================================================
Option Explicit

Public Const ABA_JOGO As String = "Hora Extra"

Private Const LIN_MAPA As Long = 2
Private Const COL_MAPA As Long = 2
Private Const LADO_PT As Double = 18            ' lado da celula, em pontos
Private Const COL_PAINEL As Long = 48           ' COL_MAPA + MAPA_W + 2
Private Const LIN_PAPEL As Long = 28            ' LIN_MAPA + MAPA_H + 2
Private Const LINHAS_PAPEL As Long = 10

Private mAba As Worksheet
Private mFundo(0 To 43, 0 To 23) As Long
Private mSimbolo(0 To 43, 0 To 23) As String
Private mCorSimbolo(0 To 43, 0 To 23) As Long
Private mPapel As String

' Monta (ou remonta) a aba do jogo do zero.
Public Sub Preparar()
    Dim area As Range
    Dim botao As Shape
    Dim x As Long, y As Long, i As Long
    Dim larguraCol As Double, medida As Double

    Application.ScreenUpdating = False

    Set mAba = Nothing
    On Error Resume Next
    Set mAba = ThisWorkbook.Worksheets(ABA_JOGO)
    On Error GoTo 0
    If mAba Is Nothing Then
        Set mAba = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        mAba.Name = ABA_JOGO
    End If
    ThisWorkbook.Activate
    mAba.Activate

    For i = mAba.Shapes.Count To 1 Step -1
        mAba.Shapes(i).Delete
    Next i
    mAba.Cells.Clear
    mAba.Cells.Interior.Color = CorPreta
    mAba.Cells.Font.Color = RGB(200, 200, 205)
    mAba.Cells.Font.Name = "Consolas"
    mAba.Cells.Font.Size = 10
    mAba.Cells.VerticalAlignment = xlVAlignCenter

    ' Celulas quadradas: a altura e em pontos, a largura e em caracteres,
    ' entao a largura e corrigida ate a celula medir o mesmo que a altura.
    Set area = mAba.Cells(LIN_MAPA, COL_MAPA).Resize(MAPA_H, MAPA_W)
    area.RowHeight = LADO_PT
    larguraCol = 2
    For i = 1 To 6
        area.ColumnWidth = larguraCol
        medida = area.Columns(1).Width
        If medida <= 0 Then
            larguraCol = larguraCol * 2
        ElseIf Abs(medida - LADO_PT) < 0.3 Then
            Exit For
        Else
            larguraCol = larguraCol * LADO_PT / medida
        End If
    Next i
    area.HorizontalAlignment = xlHAlignCenter
    area.Font.Name = "Segoe UI Symbol"
    area.Font.Size = 12
    area.Font.Bold = True
    mAba.Columns(1).ColumnWidth = 2
    mAba.Columns(COL_MAPA + MAPA_W).ColumnWidth = 2
    mAba.Columns(COL_MAPA + MAPA_W + 1).ColumnWidth = 2
    mAba.Columns(COL_PAINEL).ColumnWidth = 62
    mAba.Columns(COL_PAINEL).HorizontalAlignment = xlHAlignLeft     ' senao o numero do turno vai para a direita

    ' Painel.
    With mAba.Cells(1, COL_PAINEL)
        .Value2 = "HORA EXTRA"
        .Font.Name = "Georgia"
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = RGB(200, 40, 40)
    End With
    mAba.Rows(1).RowHeight = 30
    mAba.Cells(LIN_MAPA, COL_PAINEL).Value2 = Ac("Voc\^e ficou at\'e tarde. N\~ao foi o \'unico.")
    Rotulo LIN_MAPA + 3, "PILHA DA LANTERNA"
    Rotulo LIN_MAPA + 6, "TURNO"
    Rotulo LIN_MAPA + 9, "NO BOLSO"
    Rotulo LIN_MAPA + 12, "OBJETIVO"
    Rotulo LIN_MAPA + 15, "AGORA"
    Rotulo LIN_MAPA + 20, "TECLAS"
    mAba.Cells(LIN_MAPA + 21, COL_PAINEL).Value2 = Ac("Setas ou W A S D: andar e esbarrar (portas, quadro, cofre)")
    mAba.Cells(LIN_MAPA + 22, COL_PAINEL).Value2 = Ac("E: esperar um turno     R: recome\,car     Esc: sair do jogo")
    mAba.Cells(LIN_MAPA + 23, COL_PAINEL).Value2 = Ac("Arm\'ario (" & ChrW(9618) & "): entre para se esconder, se aquilo n\~ao estiver olhando")

    Set botao = mAba.Shapes.AddShape(5, mAba.Cells(LIN_PAPEL + LINHAS_PAPEL + 1, COL_MAPA).Left, _
                                     mAba.Cells(LIN_PAPEL + LINHAS_PAPEL + 1, COL_MAPA).Top, 150, 30)   ' 5 = retangulo arredondado
    botao.Fill.ForeColor.RGB = RGB(150, 30, 30)
    botao.Line.Visible = False
    botao.TextFrame.Characters.Text = "JOGAR DE NOVO"
    botao.TextFrame.Characters.Font.Size = 11
    botao.TextFrame.Characters.Font.Bold = True
    botao.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    botao.TextFrame.HorizontalAlignment = xlHAlignCenter
    botao.TextFrame.VerticalAlignment = xlVAlignCenter
    botao.OnAction = "Jogar"

    ' Esvazia a memoria do que esta pintado, para o primeiro quadro pintar tudo.
    For y = 0 To MAPA_H - 1
        For x = 0 To MAPA_W - 1
            mFundo(x, y) = -1
            mSimbolo(x, y) = ""
            mCorSimbolo(x, y) = -1
        Next x
    Next y
    mPapel = Chr$(0)

    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    mAba.Range(mAba.Cells(1, 1), mAba.Cells(LIN_PAPEL + LINHAS_PAPEL + 3, COL_PAINEL)).Select
    ActiveWindow.Zoom = True
    If ActiveWindow.Zoom > 150 Then ActiveWindow.Zoom = 150
    mAba.Cells(1, 1).Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    On Error GoTo 0

    Application.ScreenUpdating = True
End Sub

Private Sub Rotulo(ByVal lin As Long, ByVal legenda As String)
    With mAba.Cells(lin, COL_PAINEL)
        .Value2 = legenda
        .Font.Size = 8
        .Font.Bold = True
        .Font.Color = RGB(120, 120, 130)
    End With
End Sub

' A aba do jogo existe e e a que esta na frente?
Public Function EmFoco() As Boolean
    On Error Resume Next
    EmFoco = (ActiveSheet Is mAba)
    On Error GoTo 0
End Function

' Pinta o estado atual: mapa, painel e papel.
Public Sub Desenhar()
    Dim x As Long, y As Long
    Dim fundo As Long, simbolo As String, corSimbolo As Long

    If mAba Is Nothing Then Exit Sub
    Application.ScreenUpdating = False

    For y = 0 To MAPA_H - 1
        For x = 0 To MAPA_W - 1
            Aparencia x, y, fundo, simbolo, corSimbolo
            If fundo <> mFundo(x, y) Or simbolo <> mSimbolo(x, y) Or (Len(simbolo) > 0 And corSimbolo <> mCorSimbolo(x, y)) Then
                With mAba.Cells(LIN_MAPA + y, COL_MAPA + x)
                    If fundo <> mFundo(x, y) Then .Interior.Color = fundo
                    If simbolo <> mSimbolo(x, y) Then
                        ' O apostrofo obriga o Excel a tratar como texto ("+" sozinho viraria formula).
                        If Len(simbolo) > 0 Then .Value2 = "'" & simbolo Else .ClearContents
                    End If
                    If Len(simbolo) > 0 Then .Font.Color = corSimbolo
                End With
                mFundo(x, y) = fundo
                mSimbolo(x, y) = simbolo
                mCorSimbolo(x, y) = corSimbolo
            End If
        Next x
    Next y

    DesenharPainel
    DesenharPapel
    Application.ScreenUpdating = True
End Sub

Private Sub DesenharPainel()
    Dim barras As Long
    Dim bolso As String

    barras = (Pilha + 9) \ 10
    With mAba.Cells(LIN_MAPA + 4, COL_PAINEL)
        .Value2 = String$(barras, ChrW(9608)) & "  " & Pilha
        If Pilha > 60 Then
            .Font.Color = RGB(120, 220, 130)
        ElseIf Pilha > 20 Then
            .Font.Color = RGB(235, 200, 70)
        Else
            .Font.Color = RGB(230, 70, 60)
        End If
    End With

    mAba.Cells(LIN_MAPA + 7, COL_PAINEL).Value2 = Turnos

    If TemChave Then bolso = bolso & "chave amarela   "
    If TemFusivel Then bolso = bolso & Ac("fus\'ivel   ")
    If TemCracha Then bolso = bolso & Ac("crach\'a   ")
    If Len(bolso) = 0 Then bolso = "lanterna"
    mAba.Cells(LIN_MAPA + 10, COL_PAINEL).Value2 = bolso

    mAba.Cells(LIN_MAPA + 13, COL_PAINEL).Value2 = Objetivo()
    mAba.Cells(LIN_MAPA + 16, COL_PAINEL).Value2 = Mensagem
    With mAba.Cells(LIN_MAPA + 17, COL_PAINEL)
        .Value2 = Aviso
        .Font.Color = RGB(235, 60, 50)
        .Font.Bold = True
    End With
End Sub

' O papel sob o jogador, uma linha de texto por linha da planilha.
Private Sub DesenharPapel()
    Dim partes() As String
    Dim i As Long

    If Leitura = mPapel Then Exit Sub
    mPapel = Leitura

    mAba.Cells(LIN_PAPEL, COL_MAPA).Resize(LINHAS_PAPEL, 1).ClearContents
    If Len(Leitura) = 0 Then Exit Sub

    partes = Split(Leitura, vbLf)
    For i = 0 To UBound(partes)
        If i < LINHAS_PAPEL Then
            With mAba.Cells(LIN_PAPEL + i, COL_MAPA)
                .HorizontalAlignment = xlHAlignLeft
                .Font.Name = "Consolas"
                .Font.Size = 11
                .Font.Color = RGB(235, 228, 200)
                .Value2 = "'" & partes(i)
            End With
        End If
    Next i
End Sub
