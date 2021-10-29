'******************************
'* Cria estrutura para coleção
'******************************
Public Structure Letras
    Public Letra As String
    Public Quantidade As Integer
End Structure

'********************************************
'* Cria estrutura controle de forma quadrada
'********************************************
Public Structure Quadrada
    Public Letra As String
    Public Posicao1 As Point
    Public Posicao2 As Point
End Structure

'**********************************************
'* Cria estrutura controle de forma triangular
'**********************************************
Public Structure Triangular
    Public Letra As String
    Public Posicao1 As Point
End Structure

Module Comum

    '*****************************************
    '* Cria coleção para letras e quantidades
    '*****************************************
    Dim ColecaoLetras As New Collection

    Public Class Engrenagem

        Shared Function RetiraAcentos(Texto As String) As String

            '*****************************************************************
            '* Declara variáveis de controle e listas de acentos e reposições
            '*****************************************************************
            Dim SemAcento As String = "cCaaaaaäAAAAAeeeeEEEEiiiiiIIIIooooooOOOOOuuuuUUUU"
            Dim Acentuado As String = "çÇáàâãaªÁÀÂÃÄëéèêËÉÈÊiíïîìÍÏÌÎöóòôõºÖÓÒÔÕüúùûÜÚÙÛ"
            Dim ContadorAcentos As Integer = 0
            Dim ContadorFrase As Integer = 0
            Dim NovoTexto As String = Texto

            '***************************
            '* Percorre texto informado
            '***************************
            For ContadorFrase = 1 To Texto.Length

                '*******************************
                '* Percorre lista de acentuação
                '*******************************
                For ContadorAcentos = 1 To Acentuado.Length

                    '***********************************************
                    '* Deve realiza substituição no caracter atual?
                    '***********************************************
                    If (Mid(Texto, ContadorFrase, 1) = Mid(Acentuado, ContadorAcentos, 1)) Then

                        '****************************************
                        '* Substitui acento por texto sem acento
                        '****************************************
                        NovoTexto = Replace(NovoTexto, Mid(Acentuado, ContadorAcentos, 1), Mid(SemAcento, ContadorAcentos, 1))
                    End If
                Next
            Next

            '***************************
            '* Devolve texto sem acento
            '***************************
            Return NovoTexto
        End Function

        Public Shared Sub InicializaArea(oFormulario As Form)

            '*********************************************************
            '* Gera novo gráfico com mesma cor de fundo do formulário
            '*********************************************************
            Dim oGraphic As Graphics = oFormulario.CreateGraphics()
            oGraphic.Clear(oFormulario.BackColor)

            '*****************
            '* Libera objetos
            '*****************
            oGraphic.Dispose()
        End Sub

        Public Shared Sub DesenharQuadrado(oFormulario As Form)

            '**********************************
            '* Inicializa controles de posição
            '**********************************
            Dim Contador As Integer = 0
            Dim Formas(25) As Quadrada
            Formas(0).Letra = "A" : Formas(0).Posicao1 = New Point(332, 323) : Formas(0).Posicao2 = New Point(332 + 64, 323)
            Formas(1).Letra = "B" : Formas(1).Posicao1 = New Point(460, 323) : Formas(1).Posicao2 = New Point(460 + 64, 323)
            Formas(2).Letra = "C" : Formas(2).Posicao1 = New Point(588, 323) : Formas(2).Posicao2 = New Point(588 + 64, 323)
            Formas(3).Letra = "D" : Formas(3).Posicao1 = New Point(268, 259) : Formas(3).Posicao2 = New Point(268 + 64, 259)
            Formas(4).Letra = "E" : Formas(4).Posicao1 = New Point(396, 259) : Formas(4).Posicao2 = New Point(396 + 64, 259)
            Formas(5).Letra = "F" : Formas(5).Posicao1 = New Point(524, 259) : Formas(5).Posicao2 = New Point(524 + 64, 259)
            Formas(6).Letra = "G" : Formas(6).Posicao1 = New Point(652, 259) : Formas(6).Posicao2 = New Point(652 + 64, 259)
            Formas(7).Letra = "H" : Formas(7).Posicao1 = New Point(332, 195) : Formas(7).Posicao2 = New Point(332 + 64, 195)
            Formas(8).Letra = "I" : Formas(8).Posicao1 = New Point(460, 195) : Formas(8).Posicao2 = New Point(460 + 64, 195)
            Formas(9).Letra = "J" : Formas(9).Posicao1 = New Point(588, 195) : Formas(9).Posicao2 = New Point(588 + 64, 195)
            Formas(10).Letra = "K" : Formas(10).Posicao1 = New Point(268, 131) : Formas(10).Posicao2 = New Point(268 + 64, 131)
            Formas(11).Letra = "L" : Formas(11).Posicao1 = New Point(396, 131) : Formas(11).Posicao2 = New Point(396 + 64, 131)
            Formas(12).Letra = "M" : Formas(12).Posicao1 = New Point(524, 131) : Formas(12).Posicao2 = New Point(524 + 64, 131)
            Formas(13).Letra = "N" : Formas(13).Posicao1 = New Point(652, 131) : Formas(13).Posicao2 = New Point(652 + 64, 131)
            Formas(14).Letra = "O" : Formas(14).Posicao1 = New Point(268, 387) : Formas(14).Posicao2 = New Point(268 + 64, 387)
            Formas(15).Letra = "P" : Formas(15).Posicao1 = New Point(396, 387) : Formas(15).Posicao2 = New Point(396 + 64, 387)
            Formas(16).Letra = "Q" : Formas(16).Posicao1 = New Point(524, 387) : Formas(16).Posicao2 = New Point(524 + 64, 387)
            Formas(17).Letra = "R" : Formas(17).Posicao1 = New Point(652, 387) : Formas(17).Posicao2 = New Point(652 + 64, 387)
            Formas(18).Letra = "S" : Formas(18).Posicao1 = New Point(332, 451) : Formas(18).Posicao2 = New Point(332 + 64, 451)
            Formas(19).Letra = "T" : Formas(19).Posicao1 = New Point(460, 451) : Formas(19).Posicao2 = New Point(460 + 64, 451)
            Formas(20).Letra = "U" : Formas(20).Posicao1 = New Point(588, 451) : Formas(20).Posicao2 = New Point(588 + 64, 451)
            Formas(21).Letra = "V" : Formas(21).Posicao1 = New Point(716, 451) : Formas(21).Posicao2 = New Point(716 + 64, 451)
            Formas(22).Letra = "W" : Formas(22).Posicao1 = New Point(268, 515) : Formas(22).Posicao2 = New Point(268 + 64, 515)
            Formas(23).Letra = "X" : Formas(23).Posicao1 = New Point(396, 515) : Formas(23).Posicao2 = New Point(396 + 64, 515)
            Formas(24).Letra = "Y" : Formas(24).Posicao1 = New Point(524, 515) : Formas(24).Posicao2 = New Point(524 + 64, 515)
            Formas(25).Letra = "Z" : Formas(25).Posicao1 = New Point(652, 515) : Formas(25).Posicao2 = New Point(652 + 64, 515)

            '*********************************************************
            '* Gera novo gráfico com mesma cor de fundo do formulário
            '*********************************************************
            Dim oGraphic As Graphics = oFormulario.CreateGraphics()

            '***************************************
            '* Cria caneta com transparência de 50%
            '***************************************
            Dim oPen As New Pen(Color.FromArgb(128, &H6F, &H94, &H7A), 64)

            '******************
            '* Desenha quadros
            '******************
            For Each Letra As Letras In ColecaoLetras

                '******************************
                '* Varre estrutura de posições
                '******************************
                For Contador = 0 To Formas.Length - 1

                    '*******************************************
                    '* A letra atual corresponde à sua posição?
                    '*******************************************
                    If Formas(Contador).Letra = Letra.Letra Then

                        '****************************************
                        '* Desenha múltiplas vezes se necessário
                        '****************************************
                        For Quantidade = 1 To Letra.Quantidade
                            oGraphic.DrawLine(oPen, Formas(Contador).Posicao1, Formas(Contador).Posicao2)
                        Next

                        '*****************
                        '* Finaliza busca
                        '*****************
                        Exit For
                    End If
                Next
            Next

            '*****************
            '* Libera objetos
            '*****************
            oPen.Dispose()
            oGraphic.Dispose()
        End Sub

        Public Shared Sub DesenharTriangulo(oFormulario As Form)

            '**********************************
            '* Inicializa controles de posição
            '**********************************
            Dim Contador As Integer = 0
            Dim Formas(25) As Triangular
            Formas(0).Letra = "A" : Formas(0).Posicao1 = New Point(457, 92)
            Formas(1).Letra = "B" : Formas(1).Posicao1 = New Point(425, 155)
            Formas(2).Letra = "C" : Formas(2).Posicao1 = New Point(489, 155)
            Formas(3).Letra = "D" : Formas(3).Posicao1 = New Point(394, 217)
            Formas(4).Letra = "E" : Formas(4).Posicao1 = New Point(457, 217)
            Formas(5).Letra = "F" : Formas(5).Posicao1 = New Point(519, 217)
            Formas(6).Letra = "G" : Formas(6).Posicao1 = New Point(363, 279)
            Formas(7).Letra = "H" : Formas(7).Posicao1 = New Point(426, 279)
            Formas(8).Letra = "I" : Formas(8).Posicao1 = New Point(488, 279)
            Formas(9).Letra = "J" : Formas(9).Posicao1 = New Point(549, 279)
            Formas(10).Letra = "K" : Formas(10).Posicao1 = New Point(331, 341)
            Formas(11).Letra = "L" : Formas(11).Posicao1 = New Point(394, 341)
            Formas(12).Letra = "M" : Formas(12).Posicao1 = New Point(456, 341)
            Formas(13).Letra = "N" : Formas(13).Posicao1 = New Point(517, 341)
            Formas(14).Letra = "O" : Formas(14).Posicao1 = New Point(579, 341)
            Formas(15).Letra = "P" : Formas(15).Posicao1 = New Point(300, 404)
            Formas(16).Letra = "Q" : Formas(16).Posicao1 = New Point(363, 404)
            Formas(17).Letra = "R" : Formas(17).Posicao1 = New Point(425, 404)
            Formas(18).Letra = "S" : Formas(18).Posicao1 = New Point(486, 404)
            Formas(19).Letra = "T" : Formas(19).Posicao1 = New Point(548, 404)
            Formas(20).Letra = "U" : Formas(20).Posicao1 = New Point(610, 404)
            Formas(21).Letra = "V" : Formas(21).Posicao1 = New Point(331, 469)
            Formas(22).Letra = "W" : Formas(22).Posicao1 = New Point(394, 469)
            Formas(23).Letra = "X" : Formas(23).Posicao1 = New Point(456, 469)
            Formas(24).Letra = "Y" : Formas(24).Posicao1 = New Point(517, 469)
            Formas(25).Letra = "Z" : Formas(25).Posicao1 = New Point(579, 469)

            '*********************************************************
            '* Gera novo gráfico com mesma cor de fundo do formulário
            '*********************************************************
            Dim oGraphic As Graphics = oFormulario.CreateGraphics()

            '**************************************
            '* Cria brush com transparência de 50%
            '**************************************
            Dim oBrush As New SolidBrush(Color.FromArgb(128, &HCE, &HE0, &HE9))

            '******************
            '* Desenha quadros
            '******************
            For Each Letra As Letras In ColecaoLetras

                '******************************
                '* Varre estrutura de posições
                '******************************
                For Contador = 0 To Formas.Length - 1

                    '*******************************************
                    '* A letra atual corresponde à sua posição?
                    '*******************************************
                    If Formas(Contador).Letra = Letra.Letra Then

                        '****************************************
                        '* Desenha múltiplas vezes se necessário
                        '****************************************
                        For Quantidade = 1 To Letra.Quantidade
                            oGraphic.FillPolygon(oBrush, New PointF() {
                                                 New Point(31 + Formas(Contador).Posicao1.X, 0 + Formas(Contador).Posicao1.Y), _
                                                 New Point(0 + Formas(Contador).Posicao1.X, 63 + Formas(Contador).Posicao1.Y), _
                                                 New Point(63 + Formas(Contador).Posicao1.X, 63 + Formas(Contador).Posicao1.Y), _
                                                 New Point(31 + Formas(Contador).Posicao1.X, 0 + Formas(Contador).Posicao1.Y)})
                        Next

                        '*****************
                        '* Finaliza busca
                        '*****************
                        Exit For
                    End If
                Next
            Next

            '*****************
            '* Libera objetos
            '*****************
            oBrush.Dispose()
            oGraphic.Dispose()
        End Sub

        Public Shared Sub DesenharCirculo(oFormulario As Form)

            '**********************************
            '* Inicializa controles de posição
            '**********************************
            Dim oABC As New Point(470, 219)
            Dim oDEF As New Point(530, 230)
            Dim oGHI As New Point(537, 285)
            Dim oJKL As New Point(529, 343)
            Dim oMNO As New Point(470, 350)
            Dim oPQR As New Point(411, 343)
            Dim oSTU As New Point(403, 286)
            Dim oVWXYZ As New Point(411, 229)
            Dim oTamanho As New Size(60, 60)
            Dim Quantidade As Integer = 0
            Dim DesX As Integer = 30
            Dim DesY As Integer = 30

            '*********************************************************
            '* Gera novo gráfico com mesma cor de fundo do formulário
            '*********************************************************
            Dim oGraphic As Graphics = oFormulario.CreateGraphics()

            '**************************************
            '* Cria brush com transparência de 50%
            '**************************************
            Dim oBrush As New SolidBrush(Color.FromArgb(128, 0, 192, 0))

            '*****************
            '* Desenha elipse
            '*****************
            For Each Letra As Letras In ColecaoLetras

                '*************************************************
                '* Define rotina de movimentação para letra atual
                '*************************************************
                Select Case Letra.Letra

                    Case "A", "B", "C"

                        '****************************************
                        '* Redefine posição até atingir o centro
                        '****************************************
                        For Quantidade = 1 To Letra.Quantidade
                            oGraphic.FillEllipse(oBrush, New Rectangle(oABC, oTamanho))
                            If oABC.Y > 10 Then
                                oABC.Y -= DesY
                            End If
                        Next

                    Case "D", "E", "F"

                        '****************************************
                        '* Redefine posição até atingir o centro
                        '****************************************
                        For Quantidade = 1 To Letra.Quantidade
                            oGraphic.FillEllipse(oBrush, New Rectangle(oDEF, oTamanho))
                            If oDEF.Y > 30 Then
                                oDEF.X += DesX
                                oDEF.Y -= DesY
                            End If
                        Next

                    Case "G", "H", "I"

                        '****************************************
                        '* Redefine posição até atingir o centro
                        '****************************************
                        For Quantidade = 1 To Letra.Quantidade
                            oGraphic.FillEllipse(oBrush, New Rectangle(oGHI, oTamanho))
                            If oGHI.X < 820 Then
                                oGHI.X += DesX
                            End If
                        Next

                    Case "J", "K", "L"

                        '****************************************
                        '* Redefine posição até atingir o centro
                        '****************************************
                        For Quantidade = 1 To Letra.Quantidade
                            oGraphic.FillEllipse(oBrush, New Rectangle(oJKL, oTamanho))
                            If oJKL.Y < 540 Then
                                oJKL.X += DesX
                                oJKL.Y += DesY
                            End If
                        Next

                    Case "M", "N", "O"

                        '****************************************
                        '* Redefine posição até atingir o centro
                        '****************************************
                        For Quantidade = 1 To Letra.Quantidade
                            oGraphic.FillEllipse(oBrush, New Rectangle(oMNO, oTamanho))
                            If oMNO.Y <= 559 Then
                                oMNO.Y += DesY
                            End If
                        Next

                    Case "P", "Q", "R"

                        '****************************************
                        '* Redefine posição até atingir o centro
                        '****************************************
                        For Quantidade = 1 To Letra.Quantidade
                            oGraphic.FillEllipse(oBrush, New Rectangle(oPQR, oTamanho))
                            If oPQR.Y < 530 Then
                                oPQR.X -= DesX
                                oPQR.Y += DesY
                            End If
                        Next

                    Case "S", "T", "U"

                        '****************************************
                        '* Redefine posição até atingir o centro
                        '****************************************
                        For Quantidade = 1 To Letra.Quantidade
                            oGraphic.FillEllipse(oBrush, New Rectangle(oSTU, oTamanho))
                            If oSTU.X > 90 Then
                                oSTU.X -= DesX
                            End If
                        Next

                    Case "V", "W", "X", "Y", "Z"

                        '****************************************
                        '* Redefine posição até atingir o centro
                        '****************************************
                        For Quantidade = 1 To Letra.Quantidade
                            oGraphic.FillEllipse(oBrush, New Rectangle(oVWXYZ, oTamanho))
                            If oVWXYZ.Y > 40 Then
                                oVWXYZ.X -= DesX
                                oVWXYZ.Y -= DesY
                            End If
                        Next
                End Select
            Next

            '*****************
            '* Libera objetos
            '*****************
            oBrush.Dispose()
            oGraphic.Dispose()
        End Sub

        Public Shared Sub AnalisarFrase(Texto As String)

            '********************
            '* Declara variáveis
            '********************
            Dim ContadorLetra As Integer = 0
            Dim ContadorFrase As Integer = 0
            Dim Quantidade As Integer = 0

            '*************************
            '* Anula coleçao anterior
            '*************************
            ColecaoLetras.Clear()

            '****************************************
            '* Trata texto: maiúsculas e sem acentos
            '****************************************
            Texto = RetiraAcentos(Texto.ToUpper())

            '********************
            '* Percorre alfabeto
            '********************
            For ContadorLetra = 65 To 90

                '************************
                '* Inicializa quantidade
                '************************
                Quantidade = 0

                '******************************
                '* Varre frase contando letras
                '******************************
                For Contador = 0 To Texto.Length - 1

                    '*****************************************
                    '* A letra atual existe na posição atual?
                    '*****************************************
                    If Texto.Substring(Contador, 1) = Chr(ContadorLetra) Then

                        '************************
                        '* Incrementa quantidade
                        '************************
                        Quantidade += 1
                    End If
                Next

                '***************************************
                '* A letra atual está contida na frase?
                '***************************************
                If Quantidade Then

                    '************************************
                    '* Cria novo objeto para letra atual
                    '************************************
                    Dim Letra As New Letras
                    Letra.Letra = Chr(ContadorLetra)
                    Letra.Quantidade = Quantidade

                    '*****************************************
                    '* Adiciona letra e quantidade na coleção
                    '*****************************************
                    ColecaoLetras.Add(Letra)
                End If
            Next
        End Sub

    End Class

End Module
