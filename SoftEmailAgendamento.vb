Imports System.Runtime.InteropServices
Imports System.Timers
Imports Microsoft.VisualBasic.Logging

Public Class SoftEmailAgendamento
    Private eventId As Integer = 1
    Private TempoPadrao As Int32 = 600000
    Private TesteConfig As Boolean = False
    Private HabitarTestes As Boolean = False
    Dim pasta As String = My.Application.Info.DirectoryPath.ToString & "\Config.INI"
    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Adicione código aqui para iniciar seu serviço. Este método deve ajustar
        ' o que é necessário para que seu serviço possa executar seu trabalho.
        ' Update the service state to Start Pending.
        Dim serviceStatus As ServiceStatus = New ServiceStatus()
        serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING
        serviceStatus.dwWaitHint = 30000
        SetServiceStatus(Me.ServiceHandle, serviceStatus)

        AtualizarConfig()

        If TesteConfig = True Then
            EventLog1.WriteEntry("Iniciando o Serviço." & vbCrLf & "HABILITARTESTE: " & HabitarTestes.ToString & vbCrLf & "TEMPOATUALIZACAO: " & TempoPadrao.ToString)
        Else
            EventLog1.WriteEntry("Iniciando o Serviço." & vbCrLf & "Houve um Erro: Não Foi possível achar o arquivo de config.: " & pasta & vbCrLf & "HABILITARTESTE: " & HabitarTestes.ToString & vbCrLf & "TEMPOATUALIZACAO: " & TempoPadrao.ToString)
        End If




        '====================================================================================




        ' Set up a timer that triggers every minute.
        Dim timer As Timer = New Timer()
        timer.Interval = TempoPadrao '60000 = 60 seconds
        AddHandler timer.Elapsed, AddressOf Me.OnTimer
        timer.Start()


        '====================================================================================

        serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING
        SetServiceStatus(Me.ServiceHandle, serviceStatus)
    End Sub
    Public Sub AtualizarConfig()
        If My.Computer.FileSystem.FileExists(pasta) Then
            TesteConfig = True
            HabitarTestes = Convert.ToBoolean(lerINI(pasta, "GERAL", "HABILITARTESTE"))
            TempoPadrao = (Convert.ToInt32(lerINI(pasta, "GERAL", "TEMPOATUALIZACAO")) * 60000)
        End If
    End Sub

    'Public Sub LimparLog()
    '    For Each item In EventLog1.
    '        With item
    '            .clear
    '        End With



    '    Next
    'End Sub

    Protected Overrides Sub OnStop()
        ' Adicione código aqui para realizar qualquer limpeza necessária para parar seu serviço.
        Dim serviceStatus As ServiceStatus = New ServiceStatus()
        serviceStatus.dwCurrentState = ServiceState.SERVICE_STOP_PENDING
        serviceStatus.dwWaitHint = 30000
        SetServiceStatus(Me.ServiceHandle, serviceStatus)

        '====================================================================================

        EventLog1.WriteEntry("O serviço Foi Parado.")

        '====================================================================================

        serviceStatus.dwCurrentState = ServiceState.SERVICE_STOPPED
        SetServiceStatus(Me.ServiceHandle, serviceStatus)
    End Sub
    Protected Overrides Sub OnContinue()
        'EventLog1.WriteEntry("In OnContinue.")
        Envia_Emails()

        eventId = eventId + 1
    End Sub
    Public Sub New()
        MyBase.New()
        InitializeComponent()
        Me.EventLog1 = New System.Diagnostics.EventLog
        If Not System.Diagnostics.EventLog.SourceExists("SoftSource") Then
            System.Diagnostics.EventLog.CreateEventSource("SoftSource",
            "LogNovoSoft")
        End If
        EventLog1.Source = "SoftSource"
        EventLog1.Log = "LogNovoSoft"
    End Sub
    Private Sub OnTimer(sender As Object, e As Timers.ElapsedEventArgs)
        ' TODO: Insert monitoring activities here.
        'EventLog1.WriteEntry("Monitoring the System", EventLogEntryType.Information, eventId)
        AtualizarConfig()
        Envia_Emails()

        eventId = eventId + 1
    End Sub
    Public Enum ServiceState
        SERVICE_STOPPED = 1
        SERVICE_START_PENDING = 2
        SERVICE_STOP_PENDING = 3
        SERVICE_RUNNING = 4
        SERVICE_CONTINUE_PENDING = 5
        SERVICE_PAUSE_PENDING = 6
        SERVICE_PAUSED = 7
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure ServiceStatus
        Public dwServiceType As Long
        Public dwCurrentState As ServiceState
        Public dwControlsAccepted As Long
        Public dwWin32ExitCode As Long
        Public dwServiceSpecificExitCode As Long
        Public dwCheckPoint As Long
        Public dwWaitHint As Long
    End Structure
    Declare Auto Function SetServiceStatus Lib "advapi32.dll" (ByVal handle As IntPtr, ByRef serviceStatus As ServiceStatus) As Boolean

    Public Sub Envia_Emails()
        Dim ListaEmail As List(Of Email)
        Dim Envios As Integer = 0
        'EventLog1.WriteEntry("Iniciando Envio de Emails  " & Now.ToShortDateString & "  " & Now.ToShortTimeString, EventLogEntryType.Information, eventId)

        '==================================
        'POPULA LISTA DE EMAILS
        Using db As New DbRelatorios
            ListaEmail = db.Emails.Where(Function(m) Not m.Enviado And If(m.DataProgramada, Now) <= Now).ToList
        End Using


        Using db2 As New DbRelatorios
            For I = 0 To ListaEmail.Count - 1 Step 1
                Dim email As Email = db2.Emails.Find(CInt(ListaEmail(I).ID))
                If Not IsNothing(email) Then

                    Dim Erro As String = EmailFuncoes.Enviar_Email(email)
                    If Erro = "" Then
                        Envios += 1
                        email.DataEnvio = Now
                        email.Enviado = True

                    Else
                        email.Tentativas += 1
                        email.Erro = Erro
                    End If
                    If email.Tentativas >= 10 Then
                        EventLog1.WriteEntry("O e-mail já sofreu 10 tentativas de envio sem sucesso e será removido. Assunto:   " & email.Assunto, EventLogEntryType.Information, eventId)
                        ' MsgBox("O e-mail já sofreu 10 tentativas de envio sem sucesso e será removido. Assunto:" & vbCrLf & email.Assunto, vbCritical)
                        email.Enviado = True

                        Dim Texto As String
                        Texto = "Não foi Possivel Enviar o Email" & vbCrLf &
                                "ID: " & email.ID & vbCrLf &
                                "Destinatário:" & email.Destinatarios(0) & vbCrLf &
                                "Tentativas: " & email.Tentativas & vbCrLf &
                                "Data: " & email.Data & vbCrLf &
                                "Assunto: " & email.Assunto & vbCrLf &
                                "Erro: " & email.Erro


                        EmailFuncoes.Enviar_Email_Erro(Texto)
                    End If
                    db2.Entry(email).State = Entity.EntityState.Modified
                    db2.SaveChanges()

                Else
                    EventLog1.WriteEntry("Email não encontrado.", EventLogEntryType.Information, eventId)
                    'linha.Cells(7).Value = "Email não encontrado."
                End If
            Next
        End Using




        '==================================

        If HabitarTestes = True Then
            EventLog1.WriteEntry("Log de Envio de Emails " & vbCrLf & "Modo de Teste HABILITADO " & vbCrLf & "Enviados:  " & Envios & " Data: " & Now.ToShortDateString & "  " & Now.ToShortTimeString, EventLogEntryType.Information, eventId)
        ElseIf HabitarTestes = False And Envios > 0 Then
            EventLog1.WriteEntry("Log de Envio de Emails, Enviados:  " & Envios & " Data: " & Now.ToShortDateString & "  " & Now.ToShortTimeString, EventLogEntryType.Information, eventId)
        End If


    End Sub


End Class
