Attribute VB_Name = "Texto"
'==============================================================================
' Texto: acentos sem depender da codificacao do arquivo.
'
' O editor do VBA importa arquivos em ANSI, nao em UTF-8, entao um acento
' digitado direto no codigo pode chegar quebrado a tela. Os textos do jogo
' sao escritos em ASCII com marcas no estilo LaTeX e convertidos por Ac():
'
'     \'a  a agudo      \~a  a til      \^a  a circunflexo      \`a  a crase
'     \,c  c cedilha    \o   ordinal    \-   travessao curto
'==============================================================================
Option Explicit

Public Function Ac(ByVal s As String) As String
    Dim i As Long
    Dim n As Long
    Dim ch As String
    Dim marca As String
    Dim codigo As Long
    Dim r As String

    i = 1
    n = Len(s)
    Do While i <= n
        ch = Mid$(s, i, 1)
        codigo = 0
        If ch = "\" And i < n Then
            marca = Mid$(s, i + 1, 1)
            If marca = "o" Then
                r = r & ChrW(186)
                i = i + 2
                codigo = -1
            ElseIf marca = "-" Then
                r = r & ChrW(8211)
                i = i + 2
                codigo = -1
            ElseIf i + 1 < n Then
                codigo = Acento(marca, Mid$(s, i + 2, 1))
                If codigo > 0 Then
                    r = r & ChrW(codigo)
                    i = i + 3
                End If
            End If
        End If
        If codigo = 0 Then
            r = r & ch
            i = i + 1
        End If
    Loop
    Ac = r
End Function

Private Function Acento(ByVal marca As String, ByVal letra As String) As Long
    Select Case marca
        Case "'"
            Acento = Escolha(letra, "aeiouAEIOU", Array(225, 233, 237, 243, 250, 193, 201, 205, 211, 218))
        Case "~"
            Acento = Escolha(letra, "aoAO", Array(227, 245, 195, 213))
        Case "^"
            Acento = Escolha(letra, "aeoAEO", Array(226, 234, 244, 194, 202, 212))
        Case "`"
            Acento = Escolha(letra, "aA", Array(224, 192))
        Case ","
            Acento = Escolha(letra, "cC", Array(231, 199))
    End Select
End Function

' Posicao exata da letra (maiuscula e minuscula sao diferentes) -> codigo.
Private Function Escolha(ByVal letra As String, ByVal letras As String, ByVal codigos As Variant) As Long
    Dim i As Long

    For i = 1 To Len(letras)
        If Asc(Mid$(letras, i, 1)) = Asc(letra) Then
            Escolha = codigos(i - 1)
            Exit Function
        End If
    Next i
End Function
