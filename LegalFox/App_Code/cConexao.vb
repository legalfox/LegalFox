Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration.ConfigurationManager
Imports System
Imports System.IO
Imports System.Configuration



Public Class CConexao
    Private _objConn As New SqlConnection   'OBJETO DE CONEXAO
    Public _Erro As Boolean                    'Objeto de erro

    Public Function Conectar() As Boolean

        'Abre a conexao
        Try
            _objConn = New SqlConnection
            _objConn.ConnectionString = ConfigurationManager.ConnectionStrings("conexao").ToString
            _objConn.Open()
        Catch ex As Exception
            Return False
        End Try

        Return True

    End Function

    Public Sub Desconectar() 'matando todos objetos usados
        _objConn.Close()
        _objConn = Nothing
    End Sub

    ReadOnly Property Status() As ConnectionState
        Get
            'Retorna o Status da conexao
            Return objConn.State
        End Get
    End Property

    Public ReadOnly Property objConn() As SqlConnection
        Get
            'Retorna o Status da conexao ativa
            Return _objConn
        End Get
    End Property

    Public Property Erro() As Boolean
        Get
            Return _Erro
        End Get
        Set(ByVal Value As Boolean)
            _Erro = Value
        End Set
    End Property

End Class

