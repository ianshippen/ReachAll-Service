Module Local_Utilities
    Public Sub SetProgressBarMax(ByVal p As Integer)
    End Sub

    Public Sub SetProgressBarValue(ByVal p As Integer)
    End Sub

    Public Function GetDatabaseConfigStringForSecurity() As String
        Return CreateConnectionString(CDRSettings.cdrMultiSiteSettingsList(0).settingsConfigDictionary)
    End Function

    Public Sub EnableAgentStatusFormSaveButton(ByVal p As Boolean)
    End Sub
End Module
