Imports Microsoft.VisualBasic
Imports System.Configuration.ConfigurationManager
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports System.Security.Cryptography
Imports System.IO
Imports System
Imports System.Web.UI.WebControls
Imports System.Net.Mail

Public Class cFuncao
    Inherits cControle

    Public Shared Function GravaArquivo(ByVal vfuTemp As FileUpload, ByVal vNome As String, ByVal vAppDiretorio As String, ByRef Mensagem As String, ByRef Arquivo As String, Optional ByVal Foto As Boolean = False, Optional ByVal FotoQualidade As Integer = 100) As Boolean
        Try
            Mensagem = ""
            'Verificamos se tem alguma coisa postada 
            If Not IsNothing(vfuTemp.PostedFile) Then
                'Pegamos as informacoes do arquivo postado 
                Dim infoarquivo As New IO.FileInfo(vfuTemp.PostedFile.FileName)
                ' Extensao do arquivo carregado
                Dim extensao As String = System.IO.Path.GetExtension(vfuTemp.PostedFile.FileName)
                'Definimos onde ele será salvo 
                Dim strCaminho As String = AppSettings(vAppDiretorio) & vNome & extensao

                Arquivo = vNome & extensao

                If Foto = True Then
                    If extensao.ToUpper = ".JPG" Or extensao = ".PNG" Then
                        'Pegando do Fileupload o arquivo para salvar direto no ReduzirImagem
                        Dim ImageFile As Bitmap = New Bitmap(vfuTemp.PostedFile.InputStream)

                        If FotoQualidade < 100 Then
                            'Reduz o tamanho da imagem
                            ReduzirImagem(strCaminho, ImageFile, FotoQualidade)
                            Return True
                        End If
                    Else
                        Mensagem = "Somente arquivos JPG e PNG são permitidos, por favor, verifique."
                        Return False
                    End If
                End If

                vfuTemp.PostedFile.SaveAs(System.Web.HttpContext.Current.Server.MapPath(strCaminho))

            End If

        Catch ex As Exception
            Mensagem = "Erro no sistema, contato o Administrador e informe o erro: " & ex.Message.ToString
            Return False
        End Try

        Return True

    End Function

    Public Shared Function enviar_email(ByVal email_id As String, ByVal email_subject As String, ByVal email_body As String, Optional ByVal caminho_arquivo As String = "", Optional ByVal email_id_lojista As String = "") As Boolean
        Try

            'create the mail message
            Dim mail As New MailMessage()

            'set the addresses
            mail.From = New MailAddress(AppSettings("emailfrom"), AppSettings("emaildisplay"))

            If email_id_lojista <> "" Then
                mail.To.Add(email_id_lojista)
                mail.CC.Add("contato@ruaselojas.com.br")
                mail.CC.Add(email_id)
            Else
                mail.To.Add(email_id)
                mail.CC.Add("contato@ruaselojas.com.br")
            End If



            'set the content
            mail.Subject = email_subject
            mail.Body = email_body
            mail.IsBodyHtml = True

            'add an attachment from the filesystem
            If caminho_arquivo <> "" Then mail.Attachments.Add(New Attachment(caminho_arquivo))

            'send the message
            Dim smtp As New SmtpClient()
            With smtp
                .EnableSsl = False
                .UseDefaultCredentials = False
                .Host = AppSettings("smtp")
                .Port = AppSettings("porta")
                .Credentials = New System.Net.NetworkCredential(AppSettings("emailfrom"), AppSettings("emailpass"))

                .Send(mail)
            End With
            Return True
        Catch ex As Exception
            cLogs.InserirLog(email_id, "cFuncao", "enviar_email", ex.Message)
            Return False
        End Try


    End Function

    Public Shared Function ReduzirImagem(ByVal pNomeArquivo As String, ByVal ImageTmp As Bitmap, ByVal pQuality As Long) As Integer
        Try


            'Dim img As Bitmap teste
            Dim h, w As Integer

            'img = CType(Bitmap.FromFile(pArquivoGrande), Bitmap)

            'Se img é mais larga do que alta
            If CDec(ImageTmp.Width) / CDec(ImageTmp.Height) > 1 Then
                'If ImageTmp.Height < 480 Then
                '    h = ImageTmp.Height
                'Else
                '    h = 480
                'End If

                'If ImageTmp.Width < 640 Then
                '    w = ImageTmp.Width
                'Else
                '    w = 640
                'End If
                w = 640
                h = w * ImageTmp.Height / ImageTmp.Width
            Else
                'Se a divisao for 1, significa que a foto é quadrada
                If CDec(ImageTmp.Width) / CDec(ImageTmp.Height) = 1 Then
                    'Se o quadrado for maior que 640px diminui pra 640, se for menor ou igual deixa o tamanho original
                    If ImageTmp.Width > 640 Then
                        w = 640
                        h = 640
                    Else
                        w = ImageTmp.Width
                        h = ImageTmp.Height
                    End If
                End If

                'Se a divisao for < 1, significa que a foto mais alta que larga
                If CDec(ImageTmp.Width) / CDec(ImageTmp.Height) < 1 Then
                    'If ImageTmp.Width < 480 Then
                    '    w = ImageTmp.Width
                    'Else
                    '    w = 480
                    'End If

                    'If ImageTmp.Height < 640 Then
                    '    h = ImageTmp.Height
                    'Else
                    '    h = 640
                    'End If
                    h = 480
                    w = h * ImageTmp.Width / ImageTmp.Height

                End If
            End If

            Dim oCanvas As Bitmap = New Bitmap(ImageTmp, w, h)

            Dim encoderParams As System.Drawing.Imaging.EncoderParameters = New System.Drawing.Imaging.EncoderParameters()
            Dim g As Graphics = Graphics.FromImage(oCanvas)
            g.SmoothingMode = SmoothingMode.HighQuality

            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
            g.SmoothingMode = SmoothingMode.HighQuality
            Dim quality As Long = pQuality
            Dim encoderParam As System.Drawing.Imaging.EncoderParameter = New System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality)
            encoderParams.Param(0) = encoderParam

            Dim arrayICI As ImageCodecInfo() = ImageCodecInfo.GetImageEncoders()
            Dim jpegICI As ImageCodecInfo
            Dim x As Integer
            For x = 0 To arrayICI.Length - 1

                If (arrayICI(x).FormatDescription.Equals("JPEG")) Then
                    jpegICI = arrayICI(x)
                    Exit For
                End If

            Next

            If Not jpegICI Is Nothing Then
                'Salvar a imagem com encoder para reduzir tamanho
                oCanvas.Save(System.Web.HttpContext.Current.Server.MapPath(pNomeArquivo), jpegICI, encoderParams)
            End If

            g.Dispose()
            oCanvas.Dispose()

        Catch ex As Exception

        End Try

        Return 0
    End Function

    '8 bytes randomly selected for both the Key and the Initialization Vector
    'the IV is used to encrypt the first block of text so that any repetitive 
    'patterns are not apparent 
    Private Shared KEY_64() As Byte = {42, 16, 93, 156, 78, 4, 218, 32}
    Private Shared IV_64() As Byte = {55, 103, 246, 79, 36, 99, 167, 3}

    '24 byte or 192 bit key and IV for TripleDES
    Private Shared KEY_192() As Byte = {42, 16, 93, 156, 78, 4, 218, 32, _
    15, 167, 44, 80, 26, 250, 155, 112, _
    2, 94, 11, 204, 119, 35, 184, 197}
    Private Shared IV_192() As Byte = {55, 103, 246, 79, 36, 99, 167, 3, _
    42, 5, 62, 83, 184, 7, 209, 13, _
    145, 23, 200, 58, 173, 10, 121, 222}

    'Standard DES encryption
    Public Shared Function Encrypt(ByVal value As String) As String

        If value <> "" Then
            Dim cryptoProvider As DESCryptoServiceProvider = _
            New DESCryptoServiceProvider()
            Dim ms As MemoryStream = New MemoryStream()
            Dim cs As CryptoStream = _
            New CryptoStream(ms, cryptoProvider.CreateEncryptor(KEY_64, IV_64), _
            CryptoStreamMode.Write)
            Dim sw As StreamWriter = New StreamWriter(cs)

            sw.Write(value)
            sw.Flush()
            cs.FlushFinalBlock()
            ms.Flush()

            'convert back to a string
            Return Convert.ToBase64String(ms.GetBuffer(), 0, CInt(ms.Length))
        End If
        Return ""
    End Function


    'Standard DES decryption
    Public Shared Function Decrypt(ByVal value As String) As String
        If value <> "" Then
            Dim cryptoProvider As DESCryptoServiceProvider = _
            New DESCryptoServiceProvider()

            'convert from string to byte array
            Dim buffer() As Byte = Convert.FromBase64String(value)
            Dim ms As MemoryStream = New MemoryStream(buffer)
            Dim cs As CryptoStream = _
            New CryptoStream(ms, cryptoProvider.CreateDecryptor(KEY_64, IV_64), _
            CryptoStreamMode.Read)
            Dim sr As StreamReader = New StreamReader(cs)

            Return sr.ReadToEnd()
        End If
        Return ""
    End Function

    Public Shared Function ValidaCPF(ByVal CPF As String) As Boolean
        Dim dadosArray() As String = {"11111111111",
                                      "22222222222",
                                      "33333333333",
                                      "44444444444",
                                      "55555555555",
                                      "66666666666",
                                      "77777777777",
                                      "88888888888",
                                      "99999999999"}
        Dim i, x, n1, n2 As Integer

        CPF = CPF.Trim

        For i = 0 To dadosArray.Length - 1

            If CPF.Length <> 11 Or dadosArray(i).Equals(CPF) Then
                Return False
            End If

        Next

        'remove a mascara
        'CPF = CPF.Substring(0, 3) + CPF.Substring(4, 3) + CPF.Substring(8, 3) + CPF.Substring(12)

        For x = 0 To 1

            n1 = 0

            For i = 0 To 8 + x
                n1 = n1 + Val(CPF.Substring(i, 1)) * (10 + x - i)
            Next

            n2 = 11 - (n1 - (Int(n1 / 11) * 11))

            If n2 = 10 Or n2 = 11 Then
                n2 = 0
            End If

            If n2 <> Val(CPF.Substring(9 + x, 1)) Then
                Return False
            End If

        Next

        Return True

    End Function

    Public Shared Function ValidaCNPJ(ByVal CNPJ As String) As Boolean

        'Dim i As Integer
        Dim valida As Boolean

        CNPJ = CNPJ.Trim

        'For i = 0 To dadosArray.Length - 1

        If CNPJ.Length <> 18 Then 'Or dadosArray(i).Equals(CNPJ) Then
            Return False
        End If

        'Next

        'remove a mascara
        CNPJ = CNPJ.Substring(0, 2) + CNPJ.Substring(3, 3) + CNPJ.Substring(7, 3) + CNPJ.Substring(11, 4) + CNPJ.Substring(16)

        valida = efetivaValidacao(CNPJ)

        If valida Then
            ValidaCNPJ = True
        Else
            ValidaCNPJ = False
        End If

    End Function

    Private Shared Function efetivaValidacao(ByVal cnpj As String) As Boolean

        Dim Numero(13) As Integer
        Dim soma As Integer
        Dim i As Integer
        Dim resultado1 As Integer
        Dim resultado2 As Integer

        For i = 0 To Numero.Length - 1
            Numero(i) = CInt(cnpj.Substring(i, 1))
        Next

        soma = Numero(0) * 5 + Numero(1) * 4 + Numero(2) * 3 + Numero(3) * 2 + Numero(4) * 9 + Numero(5) * 8 + Numero(6) * 7 + _
                   Numero(7) * 6 + Numero(8) * 5 + Numero(9) * 4 + Numero(10) * 3 + Numero(11) * 2

        soma = soma - (11 * (Int(soma / 11)))

        If soma = 0 Or soma = 1 Then
            resultado1 = 0
        Else
            resultado1 = 11 - soma
        End If

        If resultado1 = Numero(12) Then
            soma = Numero(0) * 6 + Numero(1) * 5 + Numero(2) * 4 + Numero(3) * 3 + Numero(4) * 2 + Numero(5) * 9 + Numero(6) * 8 + _
                         Numero(7) * 7 + Numero(8) * 6 + Numero(9) * 5 + Numero(10) * 4 + Numero(11) * 3 + Numero(12) * 2

            soma = soma - (11 * (Int(soma / 11)))

            If soma = 0 Or soma = 1 Then
                resultado2 = 0
            Else
                resultado2 = 11 - soma
            End If

            If resultado2 = Numero(13) Then
                Return True
            Else
                Return False
            End If

        Else

            Return False

        End If

    End Function

    Public Shared Sub BuscarEndereco(ByVal CEP As String,
                                     ByRef ENDERECO As String,
                                     ByRef BAIRRO As String,
                                     ByRef CIDADE As String,
                                     ByRef ESTADO As String)
        Dim obj As New cControle
        Dim _STR As New StringBuilder
        Dim dt As Data.DataTable
        Dim dr As Data.DataRow

        _STR.Append("SELECT ")
        _STR.Append("TIPO_LOGRADOURO + ' ' + ENDERECO AS DESCR, ")
        _STR.Append("BAIRRO, ")
        _STR.Append("CIDADE, ")
        _STR.Append("ESTADO ")
        _STR.Append("FROM CEP ")
        _STR.Append("WHERE CEP = '" & Convert.ToInt32(CEP) & "' ")

        dt = obj.ExecutarBusca(_STR.ToString)

        For Each dr In dt.Rows
            ENDERECO = dr.Item("DESCR")
            BAIRRO = dr.Item("BAIRRO")
            CIDADE = dr.Item("CIDADE")
            ESTADO = dr.Item("ESTADO")
        Next

    End Sub

    Public Shared Function FormatarURL(ByVal URL As String) As String
        'replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(lower(@DESCRICAO),' ','-'),'á','a'),'é','e'),'í','i'),'ó','o'),'ú','u'),'ã','a'),'õ','o'),'ê','e'),'ô','o'),'---','-'),'&','-'),'ç','c')
        URL = URL.ToLower.Replace(" ", "-").Replace("á", "a").Replace("é", "e").Replace("í", "i").Replace("ó", "o").Replace("ú", "u").Replace("ã", "a").Replace("õ", "o").Replace("ê", "e").Replace("ô", "o").Replace("---", "-").Replace("&", "-").Replace("ç", "c").Replace(".", "-")
        Return URL
    End Function



End Class
