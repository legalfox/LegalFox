Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.Data
Imports System.Configuration.ConfigurationManager
Imports System.Web
Imports System.Text
Imports System

Public Class cControle
    Inherits CConexao

    Public Function ExecutarBusca(ByVal pSql As String) As DataTable
        Dim daBuscar As SqlDataAdapter
        Dim dtResultado As DataTable = New DataTable
        If Conectar() Then
            Try
                daBuscar = New SqlDataAdapter(pSql, objConn)
                'Executando a Query no Banco de Dados e Jogando o no DataTable
                daBuscar.Fill(dtResultado)

            Catch ex As Exception
                Dim stackframe As New Diagnostics.StackFrame

                cLogs.InserirLog(System.Web.HttpContext.Current.Session("EMAIL_ID"),
                                 System.IO.Path.GetFileName(HttpContext.Current.Request.FilePath),
                                 stackframe.GetMethod.ToString,
                                 "SQL: " & pSql & vbCrLf & _
                                 "Erro: " & ex.Message.ToString & vbCrLf & _
                                 ex.StackTrace.ToString)
                Return Nothing
            Finally
                Desconectar()
            End Try
        Else
            Return Nothing
        End If
        'Se não ocorrer erro, retorno o datatable
        Return dtResultado

    End Function

    Public Function Inserir(ByVal pSQLInsert As String) As Boolean

        Dim sqlInserir As String
        Dim cmdInserir As SqlCommand

        'Pegando o Comando de Insert
        sqlInserir = pSQLInsert

        If Conectar() Then
            Try

                'Pegando uma Instancia do obj DbCommand
                cmdInserir = objConn.CreateCommand()
                cmdInserir.CommandText = sqlInserir

                'Executando a Query no Banco de Dados
                cmdInserir.ExecuteNonQuery()

            Catch ex As Exception
                Dim stackframe As New Diagnostics.StackFrame

                cLogs.InserirLog(System.Web.HttpContext.Current.Session("EMAIL_ID"),
                                 System.IO.Path.GetFileName(HttpContext.Current.Request.FilePath),
                                 stackframe.GetMethod.ToString,
                                 "SQL: " & pSQLInsert & vbCrLf & _
                                 "Erro: " & ex.Message.ToString & vbCrLf & _
                                 ex.StackTrace.ToString)

                'Se ocorrer Erro, adiciono a Exception no obj de Erro e retono false
                Return False
            Finally
                Desconectar()
            End Try
        Else
            Return False
        End If
        'Se não ocorrer erro, retorno true
        Return True

    End Function

    Public Function Apagar(ByVal pSQLDelete As String) As Boolean

        Dim sqlApagar As String
        Dim cmdApagar As SqlCommand

        'Pegando o Comando de Insert
        sqlApagar = pSQLDelete

        If Conectar() Then
            Try

                'Pegando uma Instancia do obj DbCommand
                cmdApagar = objConn.CreateCommand()
                cmdApagar.CommandText = sqlApagar

                'Executando a Query no Banco de Dados
                cmdApagar.ExecuteNonQuery()

            Catch ex As Exception
                Dim stackframe As New Diagnostics.StackFrame

                cLogs.InserirLog(System.Web.HttpContext.Current.Session("EMAIL_ID"),
                                 System.IO.Path.GetFileName(HttpContext.Current.Request.FilePath),
                                 stackframe.GetMethod.ToString,
                                 "SQL: " & pSQLDelete & vbCrLf & _
                                 "Erro: " & ex.Message.ToString & vbCrLf & _
                                 ex.StackTrace.ToString)

                'Se ocorrer Erro, adiciono a Exception no obj de Erro e retono false
                Return False
            Finally
                Desconectar()
            End Try
        Else
            Return False
        End If
        'Se não ocorrer erro, retorno true
        Return True

    End Function

    Public Function Alterar(ByVal pSQLUpdate As String) As Boolean

        Dim sqlAlterar As String
        Dim cmdAlterar As SqlCommand

        'Pegando o Comando de Insert
        sqlAlterar = pSQLUpdate

        If Conectar() Then
            Try
                'Pegando uma Instancia do obj DbCommand
                cmdAlterar = objConn.CreateCommand()
                cmdAlterar.CommandText = pSQLUpdate

                'Executando a Query no Banco de Dados
                cmdAlterar.ExecuteNonQuery()

            Catch ex As Exception
                Dim stackframe As New Diagnostics.StackFrame

                cLogs.InserirLog(System.Web.HttpContext.Current.Session("EMAIL_ID"),
                                 System.IO.Path.GetFileName(HttpContext.Current.Request.FilePath),
                                 stackframe.GetMethod.ToString,
                                 "SQL: " & pSQLUpdate & vbCrLf & _
                                 "Erro: " & ex.Message.ToString & vbCrLf & _
                                 ex.StackTrace.ToString)

                'Se ocorrer Erro, adiciono a Exception no obj de Erro e retono false
                Return False
            Finally
                Desconectar()
            End Try
        Else
            Return False
        End If
        'Se não ocorrer erro, retorno true
        Return True

    End Function

End Class
