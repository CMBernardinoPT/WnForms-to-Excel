
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Public Class Form1
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked Then
            TextBoxValor.Text = "14"
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked Then
            TextBoxValor.Text = "9"
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked Then
            TextBoxValor.Text = "4"
        End If
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked Then
            TextBoxValor.Text = "1"
        End If
    End Sub




    Private Sub RadioButton8_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton8.CheckedChanged
        If RadioButton8.Checked Then
            TextBoxValor1.Text = "20"
        End If
    End Sub

    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton7.CheckedChanged
        If RadioButton7.Checked Then
            TextBoxValor1.Text = "9"
        End If
    End Sub

    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.CheckedChanged
        If RadioButton6.Checked Then
            TextBoxValor1.Text = "4"
        End If
    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        If RadioButton5.Checked Then
            TextBoxValor1.Text = "1"
        End If
    End Sub

    Private Sub TextBox1_TextChanged_1(sender As Object, e As EventArgs)

    End Sub


    Private Sub RadioButton12_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton12.CheckedChanged
        If RadioButton12.Checked Then
            TextBoxValor2.Text = "20"
        End If
    End Sub

    Private Sub RadioButton11_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton11.CheckedChanged
        If RadioButton11.Checked Then
            TextBoxValor2.Text = "13"
        End If
    End Sub

    Private Sub RadioButton10_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton10.CheckedChanged
        If RadioButton10.Checked Then
            TextBoxValor2.Text = "8"
        End If
    End Sub

    Private Sub RadioButton9_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton9.CheckedChanged
        If RadioButton9.Checked Then
            TextBoxValor2.Text = "3"
        End If
    End Sub


    Private Sub TextBoxScore_TextChanged(sender As Object, e As EventArgs) Handles TextBoxScore.TextChanged

    End Sub

    Private Sub RadioButton16_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton16.CheckedChanged
        If RadioButton16.Checked Then
            TextBoxValor3.Text = "20"
        End If
    End Sub

    Private Sub RadioButton15_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton15.CheckedChanged
        If RadioButton15.Checked Then
            TextBoxValor3.Text = "13"
        End If
    End Sub

    Private Sub RadioButton14_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton14.CheckedChanged
        If RadioButton14.Checked Then
            TextBoxValor3.Text = "7"
        End If
    End Sub

    Private Sub RadioButton13_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton13.CheckedChanged
        If RadioButton13.Checked Then
            TextBoxValor3.Text = "1"
        End If
    End Sub

    Private Sub RadioButton20_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton20.CheckedChanged
        If RadioButton20.Checked Then
            TextBoxValor4.Text = "20"
        End If
    End Sub

    Private Sub RadioButton19_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton19.CheckedChanged
        If RadioButton19.Checked Then
            TextBoxValor4.Text = "13"
        End If
    End Sub

    Private Sub RadioButton18_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton18.CheckedChanged
        If RadioButton18.Checked Then
            TextBoxValor4.Text = "6"
        End If
    End Sub

    Private Sub RadioButton17_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton17.CheckedChanged
        If RadioButton17.Checked Then
            TextBoxValor4.Text = "1"
        End If
    End Sub

    Private Sub RadioButton24_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton24.CheckedChanged
        If RadioButton24.Checked Then
            TextBoxValor5.Text = "20"
        End If
    End Sub

    Private Sub RadioButton23_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton23.CheckedChanged
        If RadioButton23.Checked Then
            TextBoxValor5.Text = "13"
        End If
    End Sub

    Private Sub RadioButton22_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton22.CheckedChanged
        If RadioButton22.Checked Then
            TextBoxValor5.Text = "8"
        End If
    End Sub

    Private Sub RadioButton21_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton21.CheckedChanged
        If RadioButton21.Checked Then
            TextBoxValor5.Text = "3"
        End If
    End Sub


    Private Sub RadioButton28_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton28.CheckedChanged
        If RadioButton28.Checked Then
            TextBoxValor6.Text = "20"
        End If
    End Sub

    Private Sub RadioButton27_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton27.CheckedChanged
        If RadioButton27.Checked Then
            TextBoxValor6.Text = "13"
        End If
    End Sub

    Private Sub RadioButton26_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton26.CheckedChanged
        If RadioButton26.Checked Then
            TextBoxValor6.Text = "6"
        End If
    End Sub

    Private Sub RadioButton25_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton25.CheckedChanged
        If RadioButton25.Checked Then
            TextBoxValor6.Text = "2"
        End If
    End Sub

    Private Sub TextBoxValor6_TextChanged(sender As Object, e As EventArgs) Handles TextBoxValor6.TextChanged

    End Sub

    Private Sub RadioButton32_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton32.CheckedChanged
        If RadioButton32.Checked Then
            TextBoxValor7.Text = "14"
        End If
    End Sub

    Private Sub RadioButton31_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton31.CheckedChanged
        If RadioButton31.Checked Then
            TextBoxValor7.Text = "9"
        End If
    End Sub

    Private Sub RadioButton30_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton30.CheckedChanged
        If RadioButton30.Checked Then
            TextBoxValor7.Text = "4"
        End If
    End Sub

    Private Sub RadioButton29_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton29.CheckedChanged
        If RadioButton29.Checked Then
            TextBoxValor7.Text = "1"
        End If
    End Sub

    Private Sub RadioButton36_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton36.CheckedChanged
        If RadioButton36.Checked Then
            TextBoxValor8.Text = "20"
        End If
    End Sub

    Private Sub RadioButton35_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton35.CheckedChanged
        If RadioButton35.Checked Then
            TextBoxValor8.Text = "13"
        End If
    End Sub

    Private Sub RadioButton34_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton34.CheckedChanged
        If RadioButton34.Checked Then
            TextBoxValor8.Text = "8"
        End If
    End Sub

    Private Sub RadioButton33_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton33.CheckedChanged
        If RadioButton33.Checked Then
            TextBoxValor8.Text = "3"
        End If
    End Sub




    Private Sub AtualizarMedia()
        Dim soma As Double = 0
        Dim contador As Integer = 0

        ' Lista de todas as TextBoxes que contêm os valores a serem calculados
        Dim textBoxes As TextBox() = {TextBoxValor, TextBoxValor1, TextBoxValor2, TextBoxValor3, TextBoxValor4, TextBoxValor5, TextBoxValor6, TextBoxValor7, TextBoxValor8}

        For Each tb As TextBox In textBoxes
            Dim valor As Double
            ' Verifica se o conteúdo da TextBox pode ser convertido para Double
            If Double.TryParse(tb.Text, valor) Then
                soma += valor
                contador += 1
            End If
        Next

        ' Calcular a média
        Dim media As Double = 0
        If contador > 0 Then
            media = soma / contador
        End If

        ' Atualizar TextBoxScore
        TextBoxScore.Text = media.ToString("0.##") ' Formata a média para ter até duas casas decimais
    End Sub


    Private Sub TextBoxValor_TextChanged(sender As Object, e As EventArgs) Handles TextBoxValor.TextChanged, TextBoxValor1.TextChanged, TextBoxValor2.TextChanged, TextBoxValor3.TextChanged, TextBoxValor4.TextChanged, TextBoxValor5.TextChanged, TextBoxValor6.TextChanged, TextBoxValor7.TextChanged, TextBoxValor8.TextChanged
        AtualizarMedia()
    End Sub





    Private Sub btnEnviar_Click(sender As Object, e As EventArgs) Handles btnEnviar.Click
        ' Criar uma nova instância do Excel
        Dim excelApp As New Excel.Application

        ' Tornar o Excel invisível (opcional)
        excelApp.Visible = False

        ' Abrir o workbook existente ou criar um novo
        Dim workbook As Excel.Workbook
        Dim filePath As String = "C:/Users/beny4/Desktop/testevisualstudio.xlsx"

        ' Verifique se o arquivo existe para decidir se deve criar um novo ou abrir o existente
        If System.IO.File.Exists(filePath) Then
            workbook = excelApp.Workbooks.Open(filePath)
        Else
            workbook = excelApp.Workbooks.Add()
        End If

        Dim sheet As Excel.Worksheet = workbook.Sheets(1)

        ' Encontrar a última linha preenchida na coluna A e adicionar os dados na próxima linha
        Dim lastRow As Integer = sheet.Cells(sheet.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
        Dim newEntryRow As Integer = lastRow + 1

        ' Escreva o valor da textBox na próxima linha vazia
        sheet.Cells(newEntryRow, 1).Value = TextBoxNome.Text
        sheet.Cells(newEntryRow, 2).Value = TextBoxValor.Text
        sheet.Cells(newEntryRow, 3).Value = TextBoxValor1.Text
        sheet.Cells(newEntryRow, 4).Value = TextBoxValor2.Text
        sheet.Cells(newEntryRow, 5).Value = TextBoxValor3.Text
        sheet.Cells(newEntryRow, 6).Value = TextBoxValor4.Text
        sheet.Cells(newEntryRow, 7).Value = TextBoxValor5.Text
        sheet.Cells(newEntryRow, 8).Value = TextBoxValor6.Text
        sheet.Cells(newEntryRow, 9).Value = TextBoxValor7.Text
        sheet.Cells(newEntryRow, 10).Value = TextBoxValor8.Text
        sheet.Cells(newEntryRow, 11).Value = TextBoxValor5.Text
        sheet.Cells(newEntryRow, 12).Value = TextBoxScore.Text


        ' Salvar e fechar o workbook
        If System.IO.File.Exists(filePath) Then
            workbook.Save() ' Salva o arquivo existente
        Else
            workbook.SaveAs(filePath) ' Salva como um novo arquivo se não existir
        End If

        workbook.Close(False)

        ' Limpar os objetos COM
        Marshal.ReleaseComObject(sheet)
        Marshal.ReleaseComObject(workbook)
        excelApp.Quit()
        Marshal.ReleaseComObject(excelApp)

        ' Mensagem para o usuário
        MessageBox.Show("Dados enviados para o Excel com sucesso!")
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

End Class
