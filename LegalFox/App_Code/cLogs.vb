Imports Microsoft.VisualBasic
Imports System.Text

Public Class cLogs
    Inherits cControle

    Public Shared Sub InserirLog(EMAIL_ID As String,
                                 PAGE_ID As String,
                                 FUNCSUB_ID As String,
                                 TEXT As String)
        Dim _str As New StringBuilder
        Dim obj As New cControle

        _str.Append("INSERT INTO LOGS (")
        _str.Append("EMAIL_ID, ")
        _str.Append("DATETIME, ")
        _str.Append("PAGE_ID, ")
        _str.Append("FUNCSUB_ID, ")
        _str.Append("TEXT ")
        _str.Append(") VALUES (")
        _str.Append("'" & EMAIL_ID & "', ")
        _str.Append("getdate(), ")
        _str.Append("'" & PAGE_ID & "', ")
        _str.Append("'" & FUNCSUB_ID & "', ")
        _str.Append("'" & TEXT.Replace("'", "") & "') ")

        obj.Inserir(_str.ToString)

    End Sub

End Class
