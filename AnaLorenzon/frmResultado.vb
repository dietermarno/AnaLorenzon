Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Public Class frmResultado

    Private Sub frmResultado_Paint(sender As Object, e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint

        '***************************************************
        '* Analisa frase, totalizando quantidades de letras
        '***************************************************
        Engrenagem.AnalisarFrase(frmDados.txtTexto.Text)

        '*****************************
        '* Inicializa área do desenho
        '*****************************
        Engrenagem.InicializaArea(Me)

        '************************
        '* Exibe tipo de geração
        '************************
        Me.lblTipo.Text = Me.Tag

        '***************************
        '* Desenha conforme seleção
        '***************************
        Select Case Me.Tag

            Case "quadrado"
                Engrenagem.DesenharQuadrado(Me)

            Case "circulo"
                Engrenagem.DesenharCirculo(Me)

            Case "triangulo"
                Engrenagem.DesenharTriangulo(Me)

            Case Else
                Engrenagem.DesenharQuadrado(Me)
                Engrenagem.DesenharCirculo(Me)
                Engrenagem.DesenharTriangulo(Me)
        End Select
    End Sub
End Class
