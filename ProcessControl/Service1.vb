Public Class Service1
    Public oTimer As New System.Timers.Timer(My.Settings.iTimerInterval * 60 * 1000)
    Public sExclusions As String = "Idle"


    Protected Overrides Sub OnStart(ByVal args() As String)
        CreateEvent(EventLogEntryType.Information, 1, "ProcessControl Service Started")
        AddHandler oTimer.Elapsed, AddressOf OnTimedEvent
        oTimer.Interval = 1
        oTimer.Enabled = True


    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        CreateEvent(EventLogEntryType.Information, 1, "ProcessControl Service Stopped")
    End Sub


    Sub OnTimedEvent()
        oTimer.Enabled = False
        Dim dsProcessControl As New DataSet()
        Dim dr As DataRow

        Try

            dsProcessControl.ReadXmlSchema(My.Settings.sProcessControlXMLPath)
            dsProcessControl.ReadXml(My.Settings.sProcessControlXMLPath)
        Catch ex As Exception
            CreateEvent(EventLogEntryType.Error, 2, "Exception finding process loading ProcessControl.xml " & ex.Message & ex.StackTrace)
            oTimer.Enabled = False
            Me.Stop()
            Exit Sub
        End Try

        For Each dr In dsProcessControl.Tables("Process").Rows
            sExclusions += " " & dr("Name")
        Next


        Try

            For Each dr In dsProcessControl.Tables("Process").Rows
                If dr("Name").ToString = "Default" Then
                    ResetAllProcessAffinity(dr("Affinity").ToString)
                Else
                    If dr("Priority").ToString <> "" Then
                        SetProcessPriority(dr("Name").ToString, dr("Priority").ToString)
                    End If

                    If dr("Affinity").ToString <> "" Then
                        SetProcessAffinity(dr("Name").ToString, dr("Affinity").ToString)
                    End If
                End If
            Next


        Catch ex As Exception
            CreateEvent(EventLogEntryType.Error, 2, "General Exception " & ex.Message & ex.StackTrace)
        End Try

        dsProcessControl.Dispose()

        oTimer.Interval = My.Settings.iTimerInterval * 60 * 1000
        oTimer.Enabled = True


    End Sub

    Sub SetProcessAffinity(ByVal sProcessName As String, ByVal iAfin As System.IntPtr)
        Dim oProcesses As Process()
        Try
            oProcesses = Process.GetProcessesByName(sProcessName)
        Catch ex As Exception
            CreateEvent(EventLogEntryType.Error, 2, sProcessName & " exception finding process " & ex.Message & ex.StackTrace)
            Exit Sub
        End Try


        For Each oProcess In oProcesses
            Try

                If oProcess.ProcessorAffinity <> iAfin Then
                    oProcess.ProcessorAffinity = iAfin
                    CreateEvent(EventLogEntryType.Information, 2, oProcess.ProcessName & " affinity set to " & iAfin.ToString)
                End If

            Catch ex As Exception
                CreateEvent(EventLogEntryType.Error, 2, oProcess.ProcessName & " exception setting Affinity " & ex.Message & ex.StackTrace)

            End Try
        Next


    End Sub

    Sub ResetAllProcessAffinity(ByVal iAfin As System.IntPtr)
        Dim oProcesses As Process()
        Try
            oProcesses = Process.GetProcesses()
        Catch ex As Exception
            CreateEvent(EventLogEntryType.Error, 2, "Exception finding process list" & ex.Message & ex.StackTrace)
            Exit Sub
        End Try

        For Each oProcess In oProcesses
            Try
                If sExclusions.Contains(oProcess.ProcessName) Then
                    CreateEvent(EventLogEntryType.Information, 99, oProcess.ProcessName & " skipping affinity change")
                Else
                    If oProcess.ProcessorAffinity <> iAfin Then
                        oProcess.ProcessorAffinity = iAfin
                        CreateEvent(EventLogEntryType.Information, 2, oProcess.ProcessName & " affinity set to " & iAfin.ToString)
                    End If
                End If

            Catch ex As Exception
                CreateEvent(EventLogEntryType.Error, 2, oProcess.ProcessName & " exception setting Affinity " & ex.Message & ex.StackTrace)

            End Try

Next

    End Sub

    Sub SetProcessPriority(ByVal sProcessName As String, ByVal sPriority As String)
        Dim oProcesses As Process()
        Try
            oProcesses = Process.GetProcessesByName(sProcessName)
        Catch ex As Exception
            CreateEvent(EventLogEntryType.Error, 2, sProcessName & " exception finding process " & ex.Message & ex.StackTrace)
            Exit Sub
        End Try

        Dim ProcessPriority As ProcessPriorityClass

        Select Case sPriority
            Case "RealTime"
                ProcessPriority = ProcessPriorityClass.RealTime
            Case "High"
                ProcessPriority = ProcessPriorityClass.High
            Case "AboveNormal"
                ProcessPriority = ProcessPriorityClass.AboveNormal
            Case "Normal"
                ProcessPriority = ProcessPriorityClass.Normal
            Case "BelowNormal"
                ProcessPriority = ProcessPriorityClass.BelowNormal
            Case "Idle"
                ProcessPriority = ProcessPriorityClass.Idle
            Case Else
                CreateEvent(EventLogEntryType.Warning, 2, sProcessName & " no priority specified")
                Exit Sub
        End Select


        For Each oProcess In oProcesses

            Try
                If oProcess.PriorityClass <> ProcessPriority Then
                    oProcess.PriorityClass = ProcessPriority
                    CreateEvent(EventLogEntryType.Information, 2, oProcess.ProcessName & " priority set to " & sPriority)
                End If
            Catch ex As Exception
                CreateEvent(EventLogEntryType.Error, 2, oProcess.ProcessName & " exception setting Priority " & ex.Message & ex.StackTrace)

            End Try

        Next

    End Sub

    Sub CreateEvent(ByVal EventType As EventLogEntryType, ByVal EventID As Int32, ByVal EventMessage As String)
        Dim xEventLog As EventLog = New EventLog("Application", ".", "ProcessControl")
        Dim xMyEvent As EventInstance = New EventInstance(EventID, 0, EventType)
        Dim xEventString() As String = {EventMessage}
        xEventLog.WriteEvent(xMyEvent, xEventString)

    End Sub

End Class


