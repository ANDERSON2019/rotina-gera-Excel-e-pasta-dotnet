Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel

Module Module1
    Public xlApp As New Excel.Application
    Public xlWorkBook As Excel.Workbook
    Public xlworkSheet As Excel.Worksheet
    Public localDestino As String

    Sub Main()


        Criar_Excel()


    End Sub



    Public Sub Criar_Excel()

        Dim Diretorio_Original As String = "caminho de onde esta o modelo do excel ja formatado"
        Dim NomeDiretorio As String = "caminho e nome do arquivo que sera gerado"
        Dim ExcelSheet As Object
        Dim Dbl_Linha As Double

        'ABRIR PLANILHA 
        xlWorkBook = xlApp.Workbooks.Open(Diretorio_Original)

        'SELECIONAR ABA
        xlworkSheet = xlWorkBook.Sheets("Plan1")

        'CONFIGURAÇÕES PARA VISUALIZAR PLANILHA
        ExcelSheet = xlWorkBook
        ExcelSheet.Application.Visible = False
        ExcelSheet.Windows(1).Visible = True

        Dbl_Linha = "2"

        'LOCAL PASTA
        localDestino = "C:\Users\wallace.costa\Desktop\Excel"
        Criar_Pastas()

        NomeDiretorio = localDestino & "\Planilha_Cancelados_.xlsx"

        'adicionar um for

        With xlworkSheet

            .Range("A" & Dbl_Linha).Value = "NUM_CPF"
            .Range("B" & Dbl_Linha).Value = "DES_NOME"
            .Range("C" & Dbl_Linha).Value = "ID_TRANSACAO"
            .Range("D" & Dbl_Linha).Value = "AMOUNT"
            .Range("E" & Dbl_Linha).Value = "RESPOSTA"
            .Range("F" & Dbl_Linha).Value = "AUT"

            'PULAR LINHA
            Dbl_Linha = Dbl_Linha + 1

        End With

        'SALVAR PLANILHA NO DIRETORIO 
        xlworkSheet.SaveAs(NomeDiretorio)

        'LIMPAR MEMORIA
        xlworkSheet.ClearArrows()
        xlWorkBook.Close()
        xlApp.Quit()




    End Sub

    Public Sub Criar_Pastas()

        Dim Ano As String = Format$(Now, "yyyy")
        Dim Mes As String = Format$(Now, "MM")
        Dim Dia As String = Format$(Now, "dd")
        Dim MesExt As String = MesExtenso(Val(Mes))

        localDestino = localDestino & Ano
        If Not My.Computer.FileSystem.DirectoryExists(localDestino) Then
            My.Computer.FileSystem.CreateDirectory(localDestino)
        End If

        '------------------------------------------------------------------------
        'CRIAR PASTA MES
        localDestino = localDestino & "\" & Mes & " - " & MesExt
        If Not My.Computer.FileSystem.DirectoryExists(localDestino) Then
            My.Computer.FileSystem.CreateDirectory(localDestino)
        End If

        '------------------------------------------------------------------------
        'CRIAR PASTA DIA
        localDestino = localDestino & "\" & Dia
        If Not My.Computer.FileSystem.DirectoryExists(localDestino) Then
            My.Computer.FileSystem.CreateDirectory(localDestino)
        End If

    End Sub

    Function MesExtenso(ByVal Mes As Integer) As String
        MesExtenso = ""
        Select Case Mes
            Case 1
                MesExtenso = "Janeiro"
            Case 2
                MesExtenso = "Fevereiro"
            Case 3
                MesExtenso = "Março"
            Case 4
                MesExtenso = "Abril"
            Case 5
                MesExtenso = "Maio"
            Case 6
                MesExtenso = "Junho"
            Case 7
                MesExtenso = "Julho"
            Case 8
                MesExtenso = "Agosto"
            Case 9
                MesExtenso = "Setembro"
            Case 10
                MesExtenso = "Outubro"
            Case 11
                MesExtenso = "Novembro"
            Case 12
                MesExtenso = "Dezembro"
        End Select

    End Function

End Module
