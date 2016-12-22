Imports Microsoft.VisualBasic
Imports System.Text
Imports System

Public Class cUsuarios

    Public Shared Function ExisteUsuariosEnderecos(EMAIL_ID As String) As Boolean
        Dim _str As New StringBuilder
        Dim obj As New cControle
        Dim dt As Data.DataTable

        _str.Append("SELECT ISNULL(MAX(1),0) ")
        _str.Append("FROM USUARIOS_ENDERECOS ")
        _str.Append("WHERE EMAIL_ID = '" & EMAIL_ID & "' ")

        dt = obj.ExecutarBusca(_str.ToString)

        If dt.Rows.Count > 0 Then
            Return dt.Rows(0).Item(0).ToString
        Else
            Return 0
        End If

    End Function

    Public Shared Sub BuscarDadosUsuario(EMAIL_ID As String,
                                     ByRef NOME As String,
                                     ByRef ENDERECO As String,
                                     ByRef NUMERO As String,
                                     ByRef COMPLEMENTO As String,
                                     ByRef BAIRRO As String,
                                     ByRef CIDADE As String,
                                     ByRef ESTADO As String,
                                     ByRef PAIS As String,
                                     ByRef CEP As String,
                                     ByRef DDD As String,
                                     ByRef TELEFONE As String)

        Dim obj As New cControle
        Dim _str1 As New StringBuilder
        Dim dt1 As Data.DataTable
        Dim _str2 As New StringBuilder
        Dim dt2 As Data.DataTable
        Dim _str3 As New StringBuilder
        Dim dt3 As Data.DataTable

        Dim dr As Data.DataRow



        _str1.Append("SELECT ")
        _str1.Append("NOME ")
        _str1.Append("FROM USUARIOS ")
        _str1.Append("WHERE EMAIL_ID = '" & EMAIL_ID & "' ")

        dt1 = obj.ExecutarBusca(_str1.ToString)

        For Each dr In dt1.Rows
            NOME = dr.Item("NOME").ToString
        Next

        _str2.Append("SELECT ")
        _str2.Append("A.ENDERECO AS ENDERECO, ")
        _str2.Append("A.NUMERO AS NUMERO, ")
        _str2.Append("A.COMPLEMENTO AS COMPLEMENTO, ")
        _str2.Append("A.BAIRRO AS BAIRRO, ")
        _str2.Append("A.CIDADE AS CIDADE, ")
        _str2.Append("A.ESTADO AS ESTADO, ")
        _str2.Append("A.CEP AS CEP ")
        _str2.Append("FROM USUARIOS_ENDERECOS A ")
        _str2.Append("WHERE A.EMAIL_ID = '" & EMAIL_ID & "' ")
        _str2.Append("AND A.STATUS = 'A' ")
        _str2.Append("AND A.COBRANCA = 'S' ")

        dt2 = obj.ExecutarBusca(_str2.ToString)

        For Each dr In dt2.Rows
            ENDERECO = dr.Item("ENDERECO").ToString
            NUMERO = dr.Item("NUMERO").ToString
            COMPLEMENTO = dr.Item("COMPLEMENTO").ToString
            BAIRRO = dr.Item("BAIRRO").ToString
            CIDADE = dr.Item("CIDADE").ToString
            ESTADO = dr.Item("ESTADO").ToString
            CEP = dr.Item("CEP").ToString
        Next

        _str3.Append("SELECT ")
        _str3.Append("DDD, ")
        _str3.Append("CONTATO_CONTEUDO ")
        _str3.Append("FROM USUARIOS_CONTATOS ")
        _str3.Append("WHERE EMAIL_ID = '" & EMAIL_ID & "' ")
        _str3.Append("AND TIPO_CONTATO = 'T' ")
        '_str3.Append("AND COBRANCA = 'S' ")

        dt3 = obj.ExecutarBusca(_str3.ToString)

        For Each dr In dt3.Rows
            DDD = dr.Item("DDD").ToString
            TELEFONE = dr.Item("CONTATO_CONTEUDO").ToString
        Next

    End Sub

    Public Shared Function ValidarSenha(EMAIL_ID As String, SENHA As String) As Boolean
        Dim obj As New cControle
        Dim dt As Data.DataTable

        dt = obj.ExecutarBusca("SELECT SENHA FROM USUARIOS WHERE EMAIL_ID = '" & EMAIL_ID & "'")

        If dt.Rows.Count > 0 Then
            If dt.Rows(0).Item("SENHA").ToString = cFuncao.Encrypt(SENHA) Then
                Return True
            End If
        End If

        Return False

    End Function

End Class
