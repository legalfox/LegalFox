Imports Microsoft.VisualBasic

Public Class cMenus

    Public Shared Function sql_ItensDoMenu(menu As String) As String
        Return "SELECT PG.DESCR, PG.LINK FROM PAGINAS PG, MENUS_PAGINAS MP 	WHERE PG.PAGINA_ID = MP.PAGINA_ID 	AND MP.MENU_ID = '" & menu.ToUpper & "'	ORDER BY PG.DESCR"
    End Function


    Public Shared Function MontaMenu() As String
        Try


            Dim obj As New cControle
            Dim dt_menu As Data.DataTable
            Dim ln_menu As Data.DataRow
            Dim dt_pagina As Data.DataTable
            Dim ln_pagina As Data.DataRow
            Dim Conteudo As String = ""

            'Seleciona todos os menus
            dt_menu = obj.ExecutarBusca(sql_menu(HttpContext.Current.Session("ESCRITORIO_ID")))

            Conteudo += "<div class='hor-menu'> "
            Conteudo += "<ul class='nav navbar-nav'> "

            For Each ln_menu In dt_menu.Rows
                Conteudo += "<li aria-haspopup='true' class='menu-dropdown classic-menu-dropdown active'> "
                Conteudo += "   <a href='javascript:;'> " & ln_menu.Item("menu_label") & " <span class='arrow'></span> </a> "
                Conteudo += "   <ul class='dropdown-menu pull-left'> "

                'Seleciona as paginas do menu
                dt_pagina = obj.ExecutarBusca(sql_pagina(HttpContext.Current.Session("ESCRITORIO_ID"), ln_menu.Item("menu_id")))

                For Each ln_pagina In dt_pagina.Rows
                    Conteudo += "       <li aria-haspopup='true' class=' '> "
                    Conteudo += "           <a href='" & ln_pagina.Item("pagina_id") & "' class='nav-link  '> "
                    Conteudo += "               <i class='" & ln_pagina.Item("icone") & "'></i> " & ln_pagina.Item("pagina_label") & " "
                    Conteudo += "           </a> "
                    Conteudo += "       </li> "
                Next

                Conteudo += "   </ul>"
                Conteudo += "</li>"

            Next

            Conteudo += "</ul>"
            Conteudo += "</div>"


            Return Conteudo


        Catch ex As Exception

        End Try
    End Function


    Public Shared Function sql_menu(_escritorio_id As String) As String
        Return "SELECT menu_id, menu_label, icone FROM [legalfox_prd].[dbo].[lf_menus] WHERE escritorio_id = '" & _escritorio_id & "' AND status = 'A' ORDER BY sequencia"
    End Function

    Public Shared Function sql_pagina(_escritorio_id As String, _menu_id As String) As String
        Return "SELECT pagina_id, pagina_label, icone FROM [legalfox_prd].[dbo].[lf_paginas] where escritorio_id = '" & _escritorio_id & "' and menu_id = '" & _menu_id & "' AND status = 'A' ORDER BY sequencia"
    End Function


End Class
