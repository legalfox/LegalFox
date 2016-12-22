Imports Microsoft.VisualBasic

Public Class cMenus

    Public Shared Function sql_ItensDoMenu(menu As String) As String
        Return "SELECT PG.DESCR, PG.LINK FROM PAGINAS PG, MENUS_PAGINAS MP 	WHERE PG.PAGINA_ID = MP.PAGINA_ID 	AND MP.MENU_ID = '" & menu.ToUpper & "'	ORDER BY PG.DESCR"
    End Function


    Public Shared Function MontaMenu(_pai As Integer, Optional nivel As String = "") As String
        Try

   
        Dim obj As New cControle
        Dim dt_menu As Data.DataTable
        Dim ln_menu As Data.DataRow
        Dim Conteudo As String = ""

        'Seleciona todos os menus
        dt_menu = obj.ExecutarBusca(sql_menu(_pai))

        For Each ln_menu In dt_menu.Rows
            Conteudo += "<li class='list-group-item'><a href='/categoria/" & ln_menu.Item("URL").ToString & "'><svg class='icon " & ln_menu.Item("ICONE").ToString & " glyph fs1'><use xlink:href='#" & ln_menu.Item("ICONE").ToString & "'></use></svg>" & ln_menu.Item("DESCRICAO").ToString & "</a></li>"
        Next


            Return Conteudo


        Catch ex As Exception

        End Try
    End Function


    Public Shared Function MontaMenuRuas() As String
        Dim obj As New cControle
        Dim dt_menu As Data.DataTable
        Dim ln_menu As Data.DataRow
        Dim Conteudo As String = ""

        'Seleciona todos os menus
        dt_menu = obj.ExecutarBusca(sql_menu_ruas())

        For Each ln_menu In dt_menu.Rows
            Conteudo += "<li class='list-group-item'><a href='/rua/" & cFuncao.FormatarURL(ln_menu.Item("ENDERECO").ToString) & "'>" & ln_menu.Item("ENDERECO").ToString & "</a></li>"
        Next


        Return Conteudo
    End Function




    Public Shared Function sql_menu(_pai As Integer) As String
        Return "SELECT * FROM ruaseloj_usr1.CATEGORIAS WHERE CATEGORIA_PAI = " & _pai & " AND STATUS = 'A' ORDER BY SEQ"
    End Function

    Public Shared Function sql_menu_ruas() As String
        Return "SELECT COUNT(LOJA_ID) TOTAL, ENDERECO FROM LOJAS WHERE STATUS = 'A' GROUP BY ENDERECO ORDER BY 1 DESC"
    End Function

    Public Shared Sub CabecalhoDetalhe(pTipo As String, idTipo As String, ByRef tituloPagina As Label, ByRef Bread As Label)
        Dim obj As New cControle
        Dim dt_result As Data.DataTable
        Dim ln_result As Data.DataRow
        Dim Conteudo As String = ""
        Dim _sql As String = ""

        Try

            Select Case pTipo
                Case "Categoria"
                    _sql = "SELECT DESCRICAO, ICONE FROM ruaseloj_usr1.CATEGORIAS WHERE CATEGORIA_ID = " & idTipo & " AND STATUS = 'A' ORDER BY SEQ "

                    'Seleciona todos os menus
                    dt_result = obj.ExecutarBusca(_sql)
                    For Each ln_result In dt_result.Rows
                        tituloPagina.Text = ln_result.Item("DESCRICAO").ToString
                        Bread.Text = tituloPagina.Text
                    Next

                Case "Ruas"
                    _sql = "SELECT RUA FROM ruaseloj_usr1.RUAS WHERE URL = '" & idTipo & "'"

                    'Seleciona todos os menus
                    dt_result = obj.ExecutarBusca(_sql)
                    For Each ln_result In dt_result.Rows
                        tituloPagina.Text = ln_result.Item("RUA").ToString
                        Bread.Text = tituloPagina.Text
                    Next
                Case "Pesquisa"

                    tituloPagina.Text = "Pesquisa"
                    Bread.Text = "Resultado para a palavra: " & idTipo

                Case "Contato"

                    tituloPagina.Text = "Contato"
                    Bread.Text = tituloPagina.Text

                Case "anuncie-gratis"

                    tituloPagina.Text = "Anuncie grátis"
                    Bread.Text = tituloPagina.Text


            End Select




        Catch ex As Exception

        End Try

    End Sub
End Class
