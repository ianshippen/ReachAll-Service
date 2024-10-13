Imports System.ServiceProcess
Imports System.Configuration.Install

Public Class ReachAllService
    Inherits System.ServiceProcess.ServiceBase

    Public Sub New()
        MyBase.New()
        InitializeComponents()
        ' TODO: Add any further initialization code
    End Sub

    Private Sub InitializeComponents()
        ' This is called before MyInit()
        Me.ServiceName = "ReachAllService"
        Me.AutoLog = True
        Me.CanStop = True
        SetTarget(TargetType.SERVICE)
    End Sub

    Private Sub MyInit()
        LogInfo("ReachAllService", "MyInit() called - will show as ERROR as we have not initialised the logs settings yet")
        ApplicationAndServiceShared.CommonInit()
        LogInfo("ReachAllService", "CommonInit() completed")

        With Options.optionSettingsConfigDictionary
            ' Is CRM enabled in options ?
            If .GetBooleanItem([Enum].GetName(GetType(Options.OptionSettingsConfigItems), Options.OptionSettingsConfigItems.callRecordingManagerEnable)) Then
                LogInfo("ReachAllService", "CRM enabled")

                If CallRecordingManager.LoadCRMSettings() Then
                    LogAgentStatus("ReachAllService::MyInit()", "Call Recording Manager settings loaded OK with Call Recording Manager enabled")
                Else
                    .SetItem([Enum].GetName(GetType(Options.OptionSettingsConfigItems), Options.OptionSettingsConfigItems.callRecordingManagerEnable), False)
                    LogError("ReachAllService::MyInit()", "Failed to load Call Recording Manager settings with Call Recording Manager enabled")
                End If
            End If

            ' Is Agent Status enabled in options ?
            If .GetBooleanItem([Enum].GetName(GetType(Options.OptionSettingsConfigItems), Options.OptionSettingsConfigItems.agentStatusEnable)) Then
                LogInfo("ReachAllService", "Agent Status enabled")

                If AgentStatusSettings.LoadAgentStatusSettings() Then
                    LogAgentStatus("ReachAllService::MyInit()", "Agent Status settings loaded OK with Agent Status enabled")
                Else
                    .SetItem([Enum].GetName(GetType(Options.OptionSettingsConfigItems), Options.OptionSettingsConfigItems.agentStatusEnable), False)
                    LogError("ReachAllService::MyInit()", "Failed to load Agent Status settings with Agent Status enabled")
                End If
            End If

            ' Is Call Forward Manager enabled in options ?
            If .GetBooleanItem([Enum].GetName(GetType(Options.OptionSettingsConfigItems), Options.OptionSettingsConfigItems.callForwardingManagerEnable)) Then
                LogInfo("ReachAllService", "Call Forward Manager enabled")

                If LoadCallForwardSettings() Then
                    LogCallForward("ReachAllService::MyInit()", "Call Forward settings loaded OK with Call Forwarding Manager enabled")
                    CallForwardStart()
                Else
                    .SetItem([Enum].GetName(GetType(Options.OptionSettingsConfigItems), Options.OptionSettingsConfigItems.callForwardingManagerEnable), False)
                    LogError("ReachAllService::MyInit()", "Failed to load Call Forward settings with Call Forwarding Manager enabled")
                End If
            End If

            ' Is Group Info enabled in options ?
            If .GetBooleanItem([Enum].GetName(GetType(Options.OptionSettingsConfigItems), Options.OptionSettingsConfigItems.groupInfoEnable)) Then
                LogInfo("ReachAllService", "Group Info enabled")

                If LoadGroupInfoSettings() Then
                    LogGroupInfo("ReachAllService::MyInit()", "Group Info settings loaded OK with Group Info enabled")
                Else
                    .SetItem([Enum].GetName(GetType(Options.OptionSettingsConfigItems), Options.OptionSettingsConfigItems.groupInfoEnable), False)
                    LogError("ReachAllService::MyInit()", "Failed to load Group Info settings with Group Info enabled")
                End If
            End If
        End With

        ' Kick off the scheduler with just refresh in case the server has just restarted
        Jobs.ScheduleJobs(True)
    End Sub

    ' This method starts the service.
    <MTAThread()> Shared Sub Main()
        ' To run more than one service you have to add them to the array
        System.ServiceProcess.ServiceBase.Run(New System.ServiceProcess.ServiceBase() {New ReachAllService})
        '   Dim x As New ReachAllService

        ' System.Threading.Thread.Sleep(System.Threading.Timeout.Infinite)
    End Sub

    ' Clean up any resources being used.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        MyBase.Dispose(disposing)
        ' TODO: Add cleanup code here (if required)
    End Sub

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' TODO: Add start code here (if required)
        ' to start your service.
        ' Me.Timer1.Enabled = True
        MyInit()
    End Sub

    Protected Overrides Sub OnStop()
        ' TODO: Add tear-down code here (if required) 
        ' to stop your service.
        ' Me.Timer1.Enabled = False
    End Sub

    Private Sub InitializeComponent()
        Me.ServiceName = "ReachAllService"
    End Sub
End Class








