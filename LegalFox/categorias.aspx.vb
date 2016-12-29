Public Class categorias
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        MenuLiteral.Text = cMenus.MontaMenu
    End Sub

End Class