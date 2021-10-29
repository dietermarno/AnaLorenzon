Public Class frmDados

    Sub DesligaBotoes()

        '************************
        '* Altera cores de fundo
        '************************
        Me.lblCirculo.ForeColor = Color.FromArgb(255, 180, 180, 180)
        Me.pnlCirculo.BackColor = Color.FromArgb(255, 180, 180, 180)
        Me.lblTriangulo.ForeColor = Color.FromArgb(255, 180, 180, 180)
        Me.pnlTriangulo.BackColor = Color.FromArgb(255, 180, 180, 180)
        Me.lblQuadrado.ForeColor = Color.FromArgb(255, 180, 180, 180)
        Me.pnlQuadrado.BackColor = Color.FromArgb(255, 180, 180, 180)
        Me.lblMix.ForeColor = Color.FromArgb(255, 180, 180, 180)
        Me.pnlMix.BackColor = Color.FromArgb(255, 180, 180, 180)
    End Sub

    Private Sub picGerarFormas_Click(sender As System.Object, e As System.EventArgs)

        '********************
        '* Declara variáveis
        '********************
        Dim Mensagem As String = ""

        '*************************
        '* O texto foi informado?
        '*************************
        If Me.txtTexto.Text.Trim() = "" Then

            '**********************
            '* Notifica e finaliza
            '**********************
            Mensagem = "Por favor, informe o texto."
            MsgBox(Mensagem, vbExclamation, "Aviso")
            Exit Sub
        End If

        '*****************************************
        '* Passa texto via TAG e ativa formulário
        '*****************************************
        frmResultado.Tag = Me.txtTexto.Text
        frmResultado.ShowDialog()
    End Sub

    Private Sub lblQuadrado_Click(sender As System.Object, e As System.EventArgs) Handles lblQuadrado.Click

        '*****************
        '* Define desenho
        '*****************
        frmResultado.Tag = "quadrado"

        '*****************
        '* Desliga botões
        '*****************
        DesligaBotoes()

        '***************************
        '* Altera botão selecionado
        '***************************
        Me.lblQuadrado.ForeColor = Color.FromArgb(255, 255, 255, 255)
        Me.pnlQuadrado.BackColor = Color.FromArgb(255, 255, 255, 255)
    End Sub

    Private Sub lblTriangulo_Click(sender As System.Object, e As System.EventArgs) Handles lblTriangulo.Click

        '*****************
        '* Define desenho
        '*****************
        frmResultado.Tag = "triangulo"

        '*****************
        '* Desliga botões
        '*****************
        DesligaBotoes()

        '***************************
        '* Altera botão selecionado
        '***************************
        Me.lblTriangulo.ForeColor = Color.FromArgb(255, 255, 255, 255)
        Me.pnlTriangulo.BackColor = Color.FromArgb(255, 255, 255, 255)
    End Sub

    Private Sub lblCirculo_Click(sender As System.Object, e As System.EventArgs) Handles lblCirculo.Click

        '*****************
        '* Define desenho
        '*****************
        frmResultado.Tag = "circulo"

        '*****************
        '* Desliga botões
        '*****************
        DesligaBotoes()

        '***************************
        '* Altera botão selecionado
        '***************************
        Me.lblCirculo.ForeColor = Color.FromArgb(255, 255, 255, 255)
        Me.pnlCirculo.BackColor = Color.FromArgb(255, 255, 255, 255)
    End Sub

    Private Sub lblMix_Click(sender As System.Object, e As System.EventArgs) Handles lblMix.Click

        '*****************
        '* Define desenho
        '*****************
        frmResultado.Tag = "mix"

        '*****************
        '* Desliga botões
        '*****************
        DesligaBotoes()

        '***************************
        '* Altera botão selecionado
        '***************************
        Me.lblMix.ForeColor = Color.FromArgb(255, 255, 255, 255)
        Me.pnlMix.BackColor = Color.FromArgb(255, 255, 255, 255)
    End Sub

    Private Sub lblGerarFormas_Click(sender As System.Object, e As System.EventArgs) Handles lblGerarFormas.Click

        '******************
        '* Executa gráfico
        '******************
        frmResultado.ShowDialog()
    End Sub

    Private Sub lblC_Click(sender As System.Object, e As System.EventArgs) Handles lblC.Click

        '******************
        '* Limpa digitação
        '******************
        Me.txtTexto.Text = ""
        Me.txtTexto.Focus()
    End Sub

    Private Sub frmDados_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        '*****************
        '* Desliga botões
        '*****************
        DesligaBotoes()

        '**********************
        '* Liga primeira opção
        '**********************
        lblQuadrado_Click(Nothing, Nothing)
    End Sub
End Class
