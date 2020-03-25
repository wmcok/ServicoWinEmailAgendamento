Imports System.ComponentModel.DataAnnotations
Imports System.ComponentModel.DataAnnotations.Schema
Imports System.Net.Mail
<Table("Email")>
Public Class Email
    <Key>
    <DatabaseGenerated(DatabaseGeneratedOption.Identity)>
    Public Property ID As Integer
    <MaxLength(200)>
    Public Property Assunto As String
    Public Property Texto As String
    Public Property HTML As Boolean

    Public Property Arquivos As String
    Public Property Destinatarios As String
    Public Property Tela As String

    Public Property Enviado As Boolean
    Public Property Erro As String
    Public Property Data As DateTime
    Public Property DataEnvio As DateTime?

    Public Property DataProgramada As DateTime?

    Public Property Tentativas As Integer

    Sub New()
        HTML = True
        Enviado = False
        Erro = Nothing
        Tentativas = 0
    End Sub

End Class

Public Module EmailFuncoes
    Public Function Enviar_Email(assunto As String, texto As String, destinos As List(Of String), arquivos As List(Of String), Optional HTML As Boolean = True)

        Try
            Dim email As New MailMessage()
            Dim smtpClient As New SmtpClient()
            Dim credencial As New System.Net.NetworkCredential
            credencial.UserName = "relatorios@grupoelzacosmeticos.com.br"
            credencial.Password = "rela@2012"
            smtpClient.Credentials = credencial
            smtpClient.Host = "smtp.softhair.com.br"
            smtpClient.Port = "587"
            email.From = New MailAddress("relatorios@grupoelzacosmeticos.com.br")

            If destinos.Count = 0 Then
                Return "Nenhum destinatário no e-mail."
            End If
            For Each d In destinos
                If d <> "" Then
                    'email.[To].Add(d.ToLower.Replace("P:\", "D:\DADOS DA EMPRESA\DEPARTAMENTOS\PUBLICO\"))
                    email.[To].Add(d.ToLower)
                End If

            Next
            Dim Pastas As New List(Of String)

            Pastas.Add("P:\")
            Pastas.Add("C:\PUBLICO\")
            Pastas.Add("D:\DADOS DA EMPRESA\DEPARTAMENTOS\PUBLICO\")
            Pastas.Add("\\SERVIDOR\PUBLICO\")
            If arquivos.Count > 0 Then
                For Each s In arquivos
                    If s <> "" Then
                        Dim Pasta As String = ""
                        For Each P In Pastas

                            If My.Computer.FileSystem.FileExists(s.Replace("P:\", P.ToString)) Then
                                Pasta = s.Replace("P:\", P.ToString)

                            End If
                        Next
                        If Pasta <> "" Then
                            email.Attachments.Add(New Attachment(Pasta))
                        Else
                            Return "Arquivo não encontrado no caminho:" & Pasta
                        End If
                    End If

                Next
            End If


            email.Subject = assunto
            email.IsBodyHtml = HTML
            email.Body = texto
            smtpClient.Timeout = 999999
            smtpClient.Send(email)


            Return ""
        Catch ex As Exception
            Dim erro As String = ""
            Dim Exc As Exception = ex
            While Not IsNothing(Exc)
                erro &= Exc.Message & vbCrLf
                Exc = Exc.InnerException
            End While

            Return erro

        End Try
        Exit Function
    End Function

    Public Sub Enviar_Email_Erro(texto As String)

        Try
            Dim email As New MailMessage()
            Dim smtpClient As New SmtpClient()
            Dim credencial As New System.Net.NetworkCredential
            credencial.UserName = "relatorios@grupoelzacosmeticos.com.br"
            credencial.Password = "rela@2012"
            smtpClient.Credentials = credencial
            smtpClient.Host = "smtp.softhair.com.br"
            smtpClient.Port = "587"
            email.From = New MailAddress("relatorios@grupoelzacosmeticos.com.br")

            'destinos
            email.[To].Add("rodsousa@softhair.com.br")
            email.[To].Add("cpd@softhair.com.br")



            email.Subject = "Softhair - Erro no envio de Emails Agendados no Relatorios"
            email.IsBodyHtml = True
            email.Body = texto
            smtpClient.Timeout = 999999
            smtpClient.Send(email)



        Catch ex As Exception
            Exit Sub

        End Try

    End Sub
    Public Function Enviar_Email(email As Email)


        Return Enviar_Email(email.Assunto, email.Texto, If(email.Destinatarios, "").Split(";").ToList, If(email.Arquivos, "").Split(";").ToList, email.HTML)
    End Function
End Module
