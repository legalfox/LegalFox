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
            dt_menu = obj.ExecutarBusca(sql_menu())

            Conteudo += "<div class='hor-menu'> "
            Conteudo += "<ul class='nav navbar-nav'> "

            For Each ln_menu In dt_menu.Rows
                Conteudo += "<li aria-haspopup='true' class='menu-dropdown classic-menu-dropdown active'> "
                Conteudo += "   <a href='javascript:;'> " & ln_menu.Item("menu_label") & " <span class='arrow'></span> </a> "
                Conteudo += "   <ul class='dropdown-menu pull-left'> "
                Conteudo += "       <li aria-haspopup='true' class=' active'> "
                Conteudo += "           <a href='index.html' class='nav-link  active'> "
                Conteudo += "               <i class='icon-bar-chart'></i> Item 1 "
                Conteudo += "           </a> "
                Conteudo += "       </li> "
                Conteudo += "</ul>"
                'Conteudo += " < li Class='list-group-item'><a href='/categoria/" & ln_menu.Item("URL").ToString & "'><svg class='icon " & ln_menu.Item("ICONE").ToString & " glyph fs1'><use xlink:href='#" & ln_menu.Item("ICONE").ToString & "'></use></svg>" & ln_menu.Item("DESCRICAO").ToString & "</a></li>"
            Next


            Return Conteudo


        Catch ex As Exception

        End Try
    End Function


    Public Shared Function sql_menu(_escritorio_id As String) As String
        Return "SELECT escritorio_id, menu_id, menu_label FROM [legalfox_prd].[dbo].[lf_menus] WHERE escritorio_id = " & _escritorio_id & " AND STATUS = 'A' ORDER BY SEQ"
    End Function

End Class
