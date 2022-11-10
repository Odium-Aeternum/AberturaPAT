
Imports System.Buffers
Imports System.Diagnostics.Eventing.Reader
Imports System.IO
Imports System.Net
Imports System.Security
Imports System.Text

Public Class Form1
    Dim textoFinalRMA, AnomaliaRMA, NFaturaFornecedor, DataFornecedor, Fornecedor, ProdutoRMA, TesteRMA, ObservacoesRMA As String
    Dim Transformador, MarcasUsoTorre, ExtrasTorre, TransformadorSN, Extras, PasswordTorre, DanosFisicos, PasswordAtual, SN, Anomalia, Password, Ligou, Bateria, MarcasUso, Observacoes, TextoFinal As String
    Dim TextoFinalTorre, PasswordAtualTorre, DanosFisicosTorre, AnomaliaTorre, MarcaTorre, ObservacoesTorre, corTorre As String
    Dim TextoFinalCliente, ObservacoesCliente, ProdutoCliente, NFatura, DataFatura, AnomaliaCliente, DanosFisicosCliente, EmbalagemCliente, MarcasUsoCliente As String

    Dim myCredentials As New NetworkCredential("Nanochip", "Nanosystem123", "\\nanoChipNAS\")


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load



        fsTorre = My.Computer.FileSystem.OpenTextFileReader(CaminhoFicheiroTorre)
        DATAHardWareTorre = fsTorre.ReadLine
        fsTorre.Close()

        fsTorreSoft = My.Computer.FileSystem.OpenTextFileReader(CaminhoFicheiroTorreSoftware)
        DATASoftwareTorre = fsTorreSoft.ReadLine
        fsTorreSoft.Close()

        fsPortatil = My.Computer.FileSystem.OpenTextFileReader(CaminhoFicheiroPortatil)
        DATAHardwarePortatil = fsPortatil.ReadLine
        fsPortatil.Close()

        fsPortatilSoft = My.Computer.FileSystem.OpenTextFileReader(CaminhoFicheiroPortatilSoftware)
        DATASoftwarePortatil = fsPortatilSoft.ReadLine
        fsPortatilSoft.Close()

        Label16.Text = DATAHardwarePortatil
        Label17.Text = DATAHardWareTorre
        Label69.Text = DATASoftwarePortatil
        Label70.Text = DATASoftwareTorre

        Label29.Text = Label16.Text
        Label28.Text = Label17.Text
        Label78.Text = Label69.Text
        Label77.Text = Label70.Text

        Label34.Text = Label16.Text
        Label33.Text = Label17.Text
        Label79.Text = Label70.Text
        Label80.Text = Label69.Text

        Label52.Text = Label16.Text
        Label51.Text = Label17.Text
        Label82.Text = Label69.Text
        Label81.Text = Label70.Text

        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        ComboBox3.SelectedIndex = 0
        ComboBox4.SelectedIndex = 0
    End Sub

    Dim CaminhoFicheiroTorre As String = "\\nanoChipNAS\Sam Alves Rafa\AberturaPat\DiasPrevisaoTorre.txt"
    Dim CaminhoFicheiroTorreSoftware As String = "\\nanoChipNAS\Sam Alves Rafa\AberturaPat\DiasPrevisaoTorreSoftware.txt"
    Dim CaminhoFicheiroPortatil As String = "\\nanoChipNAS\Sam Alves Rafa\AberturaPat\DiasPrevisaoPortatil.txt"
    Dim CaminhoFicheiroPortatilSoftware As String = "\\nanoChipNAS\Sam Alves Rafa\AberturaPat\DiasPrevisaoPortatilSoftware.txt"

    Dim DATAHardWareTorre, DATAHardwarePortatil, DATASoftwareTorre, DATASoftwarePortatil As String
    Dim fsTorre, fsTorreSoft, fsPortatil, fsPortatilSoft As System.IO.StreamReader

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click

        Dim indexTorre As Integer = ComboBox1.SelectedIndex
        indexTorre = indexTorre + 3

        Dim indexTorreSoftware As Integer = ComboBox2.SelectedIndex
        indexTorreSoftware = indexTorreSoftware + 3

        Dim indexPortatil As Integer = ComboBox3.SelectedIndex
        indexPortatil = indexPortatil + 3

        Dim indexPortatilSoftware As Integer = ComboBox4.SelectedIndex
        indexPortatilSoftware = indexPortatilSoftware + 3



        Dim DataHoje As Date = Now
        Dim dataPrevisãoTorre As Date = DataHoje.AddDays(indexTorre)
        Dim dataPrevisãoTorreSoftware As Date = DataHoje.AddDays(indexTorreSoftware)
        Dim dataPrevisãoPortatil As Date = DataHoje.AddDays(indexPortatil)
        Dim dataPrevisãoPortatilSoftware As Date = DataHoje.AddDays(indexPortatilSoftware)


        If ComboBox1.SelectedIndex = -1 Then
            MessageBox.Show("É necessario colocar a previsão para prosseguir!!!")
            Exit Sub
        Else



            Dim fsTorre, fsTorreSoft, fsPortatil, fsPortatilSoft As System.IO.StreamWriter
            fsTorre = My.Computer.FileSystem.OpenTextFileWriter(CaminhoFicheiroTorre, False)
            fsTorreSoft = My.Computer.FileSystem.OpenTextFileWriter(CaminhoFicheiroTorreSoftware, False)
            fsPortatil = My.Computer.FileSystem.OpenTextFileWriter(CaminhoFicheiroPortatil, False)
            fsPortatilSoft = My.Computer.FileSystem.OpenTextFileWriter(CaminhoFicheiroPortatilSoftware, False)

            fsTorre.WriteLine(dataPrevisãoTorre.ToShortDateString)
            fsTorre.Close()

            fsTorreSoft.WriteLine(dataPrevisãoTorreSoftware.ToShortDateString)
            fsTorreSoft.Close()

            fsPortatil.WriteLine(dataPrevisãoPortatil.ToShortDateString)
            fsPortatil.Close()

            fsPortatilSoft.WriteLine(dataPrevisãoPortatilSoftware.ToShortDateString)
            fsPortatilSoft.Close()



        End If

        fsTorre = My.Computer.FileSystem.OpenTextFileReader(CaminhoFicheiroTorre)
        DATAHardWareTorre = fsTorre.ReadLine
        fsTorre.Close()

        fsTorreSoft = My.Computer.FileSystem.OpenTextFileReader(CaminhoFicheiroTorreSoftware)
        DATASoftwareTorre = fsTorreSoft.ReadLine
        fsTorreSoft.Close()

        fsPortatil = My.Computer.FileSystem.OpenTextFileReader(CaminhoFicheiroPortatil)
        DATAHardwarePortatil = fsPortatil.ReadLine
        fsPortatil.Close()

        fsPortatilSoft = My.Computer.FileSystem.OpenTextFileReader(CaminhoFicheiroPortatilSoftware)
        DATASoftwarePortatil = fsPortatilSoft.ReadLine
        fsPortatilSoft.Close()

        Label16.Text = DATAHardwarePortatil
        Label17.Text = DATAHardWareTorre
        Label69.Text = DATASoftwarePortatil
        Label70.Text = DATASoftwareTorre

        Label29.Text = Label16.Text
        Label28.Text = Label17.Text
        Label78.Text = Label69.Text
        Label77.Text = Label70.Text

        Label34.Text = Label16.Text
        Label33.Text = Label17.Text
        Label79.Text = Label70.Text
        Label80.Text = Label69.Text

        Label52.Text = Label16.Text
        Label51.Text = Label17.Text
        Label82.Text = Label69.Text
        Label81.Text = Label70.Text

        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        ComboBox3.SelectedIndex = 0
        ComboBox4.SelectedIndex = 0

    End Sub
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click


        textoFinalRMA = ""
        AnomaliaRMA = RichTextBox8.Text
        DataFornecedor = TextBox19.Text
        Fornecedor = TextBox22.Text
        ObservacoesRMA = RichTextBox7.Text
        ProdutoRMA = TextBox18.Text
        NFaturaFornecedor = TextBox21.Text

        If AnomaliaRMA = Nothing Then
            MessageBox.Show("É necessario colocar a anomalia para prosseguir!!!")
            Exit Sub
        ElseIf NFatura = Nothing Then
            MessageBox.Show("É necessario colocar o Nº da fatura para prosseguir!!!")
            Exit Sub
        ElseIf DataFornecedor = Nothing Then
            MessageBox.Show("É necessario colocar a data da fatura para prosseguir!!!")
            Exit Sub
        ElseIf Fornecedor = Nothing Then
            MessageBox.Show("É necessario colocar o fornecedor para prosseguir!!!")
            Exit Sub
        ElseIf produtoRMA = Nothing Then
            MessageBox.Show("É necessario colocar o produto para prosseguir!!!")
            Exit Sub
        ElseIf CheckedListBox10.SelectedIndex = -1 Then
            MessageBox.Show("É necessario indicar se a anomalia está confirmada!!!")
            Exit Sub
        Else
            If ObservacoesRMA = Nothing Then
                ObservacoesRMA = Nothing
            Else
                ObservacoesRMA = "• Observações: " & vbCrLf & "    • " & RichTextBox7.Text
            End If


        End If


        textoFinalRMA = "• " & AnomaliaRMA & vbCrLf & vbCrLf & "• Fatura Fornecedor:" & vbCrLf & "    • Nº: " & NFaturaFornecedor & vbCrLf & "    • Data: " & DataFornecedor & vbCrLf & "    • Fornecedor: " & Fornecedor & vbCrLf & vbCrLf & "• Produto: " & ProdutoRMA & vbCrLf & vbCrLf & "• Teste:" & vbCrLf & "    • " & TesteRMA & vbCrLf & vbCrLf & ObservacoesRMA
        TextBox20.Text = textoFinalRMA


    End Sub
    Private Sub CheckedListBox10_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox10.SelectedIndexChanged
        Dim idx, sidx As Integer
        sidx = CheckedListBox10.SelectedIndex
        For idx = 0 To CheckedListBox10.Items.Count - 1
            If idx <> sidx Then
                CheckedListBox10.SetItemChecked(idx, False)
            Else
                CheckedListBox10.SetItemChecked(sidx, True)
                If sidx = 0 Then
                    TesteRMA = "Anomalia ainda por confirmar"
                ElseIf sidx = 1 Then
                    TesteRMA = "Anomalia já confirmada"
                End If
            End If
        Next
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Clipboard.SetText(TextBox13.Text)
    End Sub
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        TextoFinalCliente = ""
        DanosFisicosCliente = Nothing
        AnomaliaCliente = RichTextBox6.Text

        If AnomaliaCliente = Nothing Then
            MessageBox.Show("É necessario colocar a anomalia para prosseguir!!!")
            Exit Sub
        Else
            NFatura = TextBox15.Text
            DataFatura = TextBox16.Text
            ProdutoCliente = TextBox17.Text

            If NFatura = Nothing Then
                MessageBox.Show("É necessario colocar o nº da fatura para prosseguir!!!")
                Exit Sub
            ElseIf DataFatura = Nothing Then
                MessageBox.Show("É necessario colocar a data da fatura para prosseguir!!!")
                Exit Sub
            ElseIf ProdutoCliente = Nothing Then
                MessageBox.Show("É necessario colocar o produto para prosseguir!!!")
                Exit Sub
            End If

            If CheckedListBox9.SelectedIndex = -1 Then
                MessageBox.Show("É necessario indicar o estado da embalagem!!!")
                Exit Sub
            ElseIf CheckedListBox8.SelectedIndex = -1 Then
                MessageBox.Show("É necessario indicar o estado do equipamento!!!")
                Exit Sub
            ElseIf CheckedListBox8.SelectedIndex = 1 And TextBox14.Text = Nothing Then
                MessageBox.Show("É necessario indicar o dano fisico!!!")
                Exit Sub
            ElseIf CheckedListBox8.SelectedIndex = 1 And TextBox14.Text <> Nothing Then
                DanosFisicosCliente = ":(" & TextBox14.Text & ")"
            ElseIf CheckedListBox8.SelectedIndex = 0 Then
                TextBox14.Text = Nothing
            End If

            If RichTextBox5.Text = Nothing Then
                ObservacoesCliente = ""
            Else
                ObservacoesCliente = vbCrLf & vbCrLf & "• Observações: " & vbCrLf & "    • " & RichTextBox5.Text
            End If
        End If



        TextoFinalCliente = "• " & AnomaliaCliente & vbCrLf & vbCrLf & "• Fatura:" & vbCrLf & "    • Nº da Fatura: " & NFatura & vbCrLf & "    • Data da Fatura: " & DataFatura & vbCrLf & vbCrLf & "• Produto: " & ProdutoCliente & vbCrLf & EmbalagemCliente & vbCrLf & vbCrLf & "• Estado do Equipamento:" & vbCrLf & "    • " & MarcasUsoCliente & DanosFisicosCliente & vbCrLf & ObservacoesCliente
        TextBox13.Text = TextoFinalCliente



    End Sub
    Private Sub CheckedListBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox8.SelectedIndexChanged
        Dim idx, sidx As Integer
        sidx = CheckedListBox8.SelectedIndex
        For idx = 0 To CheckedListBox8.Items.Count - 1
            If idx <> sidx Then
                CheckedListBox8.SetItemChecked(idx, False)
            Else
                CheckedListBox8.SetItemChecked(sidx, True)
                If sidx = 0 Then
                    MarcasUsoCliente = "Normal"
                    TextBox14.Visible = False
                ElseIf sidx = 1 Then
                    MarcasUsoCliente = "Danos Físicos Visíveis"
                    TextBox14.Visible = True
                End If
            End If
        Next
    End Sub
    Private Sub CheckedListBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox9.SelectedIndexChanged
        Dim idx, sidx As Integer
        sidx = CheckedListBox9.SelectedIndex
        For idx = 0 To CheckedListBox9.Items.Count - 1
            If idx <> sidx Then
                CheckedListBox9.SetItemChecked(idx, False)
            Else
                CheckedListBox9.SetItemChecked(sidx, True)
                If sidx = 0 Then
                    EmbalagemCliente = "• Sem embalagem"

                ElseIf sidx = 1 Then
                    EmbalagemCliente = "• Embalagem completa"

                ElseIf sidx = 2 Then
                    EmbalagemCliente = "• Embalagem danificada"

                ElseIf sidx = 3 Then
                    EmbalagemCliente = "• Embalagem com conteudos em falta"

                End If
            End If
        Next
    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Close()
    End Sub
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Clipboard.SetText(TextBox20.Text)
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Close()
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Clipboard.SetText(TextBox12.Text)
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Close()
    End Sub
    Private Sub CheckedListBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox7.SelectedIndexChanged
        Dim idx, sidx As Integer
        sidx = CheckedListBox7.SelectedIndex
        For idx = 0 To CheckedListBox7.Items.Count - 1
            If idx <> sidx Then
                CheckedListBox7.SetItemChecked(idx, False)
            Else
                CheckedListBox7.SetItemChecked(sidx, True)
                If sidx = 0 Then
                    MarcasUsoTorre = "Normal"
                    TextBox11.Visible = False
                ElseIf sidx = 1 Then
                    MarcasUsoTorre = "Normal com alguns riscos"
                    TextBox11.Visible = False
                ElseIf sidx = 2 Then
                    MarcasUsoTorre = "Normal com sujidade"
                    TextBox11.Visible = False
                ElseIf sidx = 3 Then
                    MarcasUsoTorre = "Normal com sujidade e riscos"
                    TextBox11.Visible = False
                Else
                    MarcasUsoTorre = "Danos Físicos Visíveis"
                    TextBox11.Visible = True
                End If
            End If
        Next
    End Sub
    Private Sub CheckedListBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox6.SelectedIndexChanged
        Dim idx, sidx As Integer
        sidx = CheckedListBox6.SelectedIndex
        For idx = 0 To CheckedListBox6.Items.Count - 1
            If idx <> sidx Then
                CheckedListBox6.SetItemChecked(idx, False)
            Else
                CheckedListBox6.SetItemChecked(sidx, True)
                If sidx = 0 Then
                    PasswordTorre = ""
                    TextBox9.Visible = True
                    Label22.Visible = True
                ElseIf sidx = 1 Then
                    PasswordTorre = "Não"
                    TextBox9.Visible = False
                    Label22.Visible = False
                Else
                    PasswordTorre = "Não Sabe"
                    TextBox9.Visible = False
                    Label22.Visible = False
                End If
            End If
        Next
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click


        TextoFinalTorre = ""
        PasswordAtualTorre = Nothing
        DanosFisicosTorre = Nothing
        AnomaliaTorre = RichTextBox2.Text
        If AnomaliaTorre = Nothing Then
            MessageBox.Show("É necessario colocar a anomalia para prosseguir!!!")
            Exit Sub
        Else
            MarcaTorre = TextBox7.Text

            If MarcaTorre = Nothing Then
                MessageBox.Show("É necessario colocar a marca da caixa para prosseguir!!!")
                Exit Sub
            ElseIf TextBox8.Text = Nothing Then
                MessageBox.Show("É necessario colocar a cor da caixa para prosseguir!!!")
                Exit Sub
            Else
                corTorre = TextBox8.Text


                If CheckedListBox6.SelectedIndex = -1 Then
                    MessageBox.Show("É necessario indicar se a torre contem password!!!")
                    Exit Sub
                ElseIf CheckedListBox6.SelectedIndex = 1 Or CheckedListBox6.SelectedIndex = 2 Then
                    TextBox9.Text = Nothing
                ElseIf CheckedListBox6.SelectedIndex = 0 And TextBox9.Text = Nothing Then
                    MessageBox.Show("É necessario indicar a password!!!")
                    Exit Sub
                ElseIf CheckedListBox6.SelectedIndex = 0 And TextBox9.Text <> Nothing Then
                    PasswordAtualTorre = TextBox9.Text
                End If

                If CheckedListBox7.SelectedIndex = -1 Then
                    MessageBox.Show("É necessario indicar o estado do equipamento!!!")
                    Exit Sub
                ElseIf CheckedListBox7.SelectedIndex = 4 And TextBox11.Text = Nothing Then
                    MessageBox.Show("É necessario indicar o dano fisico!!!")
                    Exit Sub
                ElseIf CheckedListBox7.SelectedIndex = 4 And TextBox11.Text <> Nothing Then
                    DanosFisicosTorre = ":(" & TextBox11.Text & ")"
                ElseIf CheckedListBox7.SelectedIndex <> 4 Then
                    TextBox11.Text = Nothing
                End If

                If RichTextBox4.Text = Nothing Then
                    ObservacoesTorre = ""
                Else
                    ObservacoesTorre = vbCrLf & vbCrLf & "• Observações: " & vbCrLf & "    • " & RichTextBox4.Text
                End If
            End If
        End If
        If TextBox10.Text = Nothing Then
            Extras = ""
        ElseIf TextBox10.Text <> Nothing Then
            Extras = "    • Extras:" & TextBox10.Text
        End If

        TextoFinalTorre = "• " & AnomaliaTorre & vbCrLf & vbCrLf & "• Conteúdo:" & vbCrLf & "    • Marca da Caixa: " & MarcaTorre & vbCrLf & "    • Cor da Caixa: " & corTorre & vbCrLf & ExtrasTorre & vbCrLf & vbCrLf & "• Password: " & PasswordTorre & PasswordAtualTorre & vbCrLf & vbCrLf & "• Estado do Equipamento:" & vbCrLf & "    • " & MarcasUsoTorre & DanosFisicosTorre & vbCrLf & ObservacoesTorre
        TextBox12.Text = TextoFinalTorre







    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Close()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Clipboard.SetText(TextBox6.Text)
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextoFinal = ""
        PasswordAtual = Nothing
        DanosFisicos = Nothing
        Anomalia = RichTextBox1.Text
        If Anomalia = Nothing Then
            MessageBox.Show("É necessario colocar a anomalia para prosseguir!!!")
            Exit Sub
        Else
            SN = TextBox1.Text

            If SN = Nothing Then
                MessageBox.Show("É necessario colocar o numero de serie do portátil para prosseguir!!!")
                Exit Sub
            ElseIf CheckedListBox2.SelectedIndex = -1 Then
                MessageBox.Show("É necessario indicar se o portátil ligou ao balcão!!!")
                Exit Sub
            ElseIf CheckedListBox1.SelectedIndex = -1 Then
                MessageBox.Show("É necessario indicar se o cliente deixou transformador!!!")
                Exit Sub
            ElseIf CheckedListBox1.SelectedIndex = 0 And TextBox2.Text = Nothing Then
                MessageBox.Show("É necessario indicar o numero de serie do transformador!!!")
                Exit Sub
            ElseIf TextBox2.Text = Nothing Then
                TransformadorSN = ""
            End If
            If CheckedListBox1.SelectedIndex = 0 And TextBox2.Text <> Nothing Then
                TransformadorSN = "  • SN: " & TextBox2.Text
                If CheckedListBox3.SelectedIndex = -1 Then
                    MessageBox.Show("É necessario indicar se o portátil contem bateria!!!")
                    Exit Sub
                ElseIf CheckedListBox4.SelectedIndex = -1 Then
                    MessageBox.Show("É necessario indicar se o portátil contem password!!!")
                    Exit Sub
                ElseIf CheckedListBox4.SelectedIndex = 1 Or CheckedListBox4.SelectedIndex = 2 Then
                    TextBox4.Text = Nothing
                ElseIf CheckedListBox4.SelectedIndex = 0 And TextBox4.Text = Nothing Then
                    MessageBox.Show("É necessario indicar a password!!!")
                    Exit Sub
                ElseIf CheckedListBox4.SelectedIndex = 0 And TextBox4.Text <> Nothing Then
                    PasswordAtual = TextBox4.Text
                End If
            End If

            If CheckedListBox5.SelectedIndex = -1 Then
                MessageBox.Show("É necessario indicar o estado do equipamento!!!")
                Exit Sub
            ElseIf CheckedListBox5.SelectedIndex = 4 And TextBox5.Text = Nothing Then
                MessageBox.Show("É necessario indicar o dano fisico!!!")
                Exit Sub
            ElseIf CheckedListBox5.SelectedIndex = 4 And TextBox5.Text <> Nothing Then
                DanosFisicos = ":(" & TextBox5.Text & ")"
            ElseIf CheckedListBox5.SelectedIndex <> 4 Then
                TextBox5.Text = Nothing
            End If

            If RichTextBox3.Text = Nothing Then
                Observacoes = ""
            Else
                Observacoes = vbCrLf & vbCrLf & "• Observações: " & vbCrLf & "    • " & RichTextBox3.Text
            End If

        End If
        If TextBox3.Text = Nothing Then
            Extras = ""
        ElseIf TextBox3.Text <> Nothing Then
            Extras = "    • Extras:" & TextBox3.Text
        End If

        TextoFinal = "• " & Anomalia & vbCrLf & vbCrLf & "• Conteúdo:" & vbCrLf & "    • Nº de Serie: " & SN & vbCrLf & "    • Portatil Ligou? " & Ligou & vbCrLf & "    • Transformador: " & Transformador & TransformadorSN & vbCrLf & "    • Bateria: " & Bateria & vbCrLf & Extras & vbCrLf & vbCrLf & "• Password: " & Password & PasswordAtual & vbCrLf & vbCrLf & "• Estado do Equipamento:" & vbCrLf & "    • " & MarcasUso & DanosFisicos & vbCrLf & Observacoes
        TextBox6.Text = TextoFinal








    End Sub
    Private Sub CheckedListBox1_mouseclick(sender As Object, e As EventArgs) Handles CheckedListBox1.SelectedIndexChanged
        Dim idx, sidx As Integer


        sidx = CheckedListBox1.SelectedIndex
        For idx = 0 To CheckedListBox1.Items.Count - 1
            If idx <> sidx Then
                CheckedListBox1.SetItemChecked(idx, False)
            Else
                CheckedListBox1.SetItemChecked(sidx, True)
                If sidx = 0 Then
                    Transformador = "Sim"
                    TextBox2.Visible = True
                Else
                    Transformador = "Não"
                    TextBox2.Visible = False
                    TextBox2.Text = Nothing
                End If
            End If
        Next
    End Sub
    Private Sub CheckedListBox2_mouseclick(sender As Object, e As EventArgs) Handles CheckedListBox2.SelectedIndexChanged
        Dim idx, sidx As Integer
        sidx = CheckedListBox2.SelectedIndex
        For idx = 0 To CheckedListBox2.Items.Count - 1
            If idx <> sidx Then
                CheckedListBox2.SetItemChecked(idx, False)
            Else
                CheckedListBox2.SetItemChecked(sidx, True)

                If sidx = 0 Then
                    Ligou = "Sim"
                ElseIf sidx = 1 Then
                    Ligou = "Não"
                Else
                    Ligou = "Não foi possível testar ao balcão"
                End If
            End If
        Next

    End Sub
    Private Sub CheckedListBox3_mouseclick(sender As Object, e As EventArgs) Handles CheckedListBox3.SelectedIndexChanged
        Dim idx, sidx As Integer
        sidx = CheckedListBox3.SelectedIndex
        For idx = 0 To CheckedListBox3.Items.Count - 1
            If idx <> sidx Then
                CheckedListBox3.SetItemChecked(idx, False)
            Else
                CheckedListBox3.SetItemChecked(sidx, True)
                If sidx = 0 Then
                    Bateria = "Sim"
                ElseIf sidx = 1 Then
                    Bateria = "Não"
                Else
                    Bateria = "Interna"
                End If
            End If
        Next
    End Sub
    Private Sub CheckedListBox4_mouseclick(sender As Object, e As EventArgs) Handles CheckedListBox4.SelectedIndexChanged
        Dim idx, sidx As Integer
        sidx = CheckedListBox4.SelectedIndex
        For idx = 0 To CheckedListBox4.Items.Count - 1
            If idx <> sidx Then
                CheckedListBox4.SetItemChecked(idx, False)
            Else
                CheckedListBox4.SetItemChecked(sidx, True)
                If sidx = 0 Then
                    Password = ""
                    TextBox4.Visible = True
                    Label12.Visible = True
                ElseIf sidx = 1 Then
                    Password = "Não"
                    TextBox4.Visible = False
                    Label12.Visible = False
                Else
                    Password = "Não Sabe"
                    TextBox4.Visible = False
                    Label12.Visible = False
                End If
            End If
        Next
    End Sub
    Private Sub CheckedListBox5_mouseclick(sender As Object, e As EventArgs) Handles CheckedListBox5.SelectedIndexChanged
        Dim idx, sidx As Integer
        sidx = CheckedListBox5.SelectedIndex
        For idx = 0 To CheckedListBox5.Items.Count - 1
            If idx <> sidx Then
                CheckedListBox5.SetItemChecked(idx, False)
            Else
                CheckedListBox5.SetItemChecked(sidx, True)
                If sidx = 0 Then
                    MarcasUso = "Normal"
                    TextBox5.Visible = False
                ElseIf sidx = 1 Then
                    MarcasUso = "Normal com alguns riscos"
                    TextBox5.Visible = False
                ElseIf sidx = 2 Then
                    MarcasUso = "Normal com sujidade"
                    TextBox5.Visible = False
                ElseIf sidx = 3 Then
                    MarcasUso = "Normal com sujidade e riscos"
                    TextBox5.Visible = False
                Else
                    MarcasUso = "Danos Físicos Visíveis"
                    TextBox5.Visible = True
                End If
            End If
        Next
    End Sub

End Class
