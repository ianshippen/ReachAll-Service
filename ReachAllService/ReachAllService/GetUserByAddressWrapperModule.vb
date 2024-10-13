Module GetUserByAddressWrapperModule
    Private myPbxConfig As Object = Nothing

    Public Function GetUserByAddressWrapper(ByRef p_target As String, ByRef p_info As Object) As Boolean
        Dim result As Boolean = True

        result = InitialiseMyPBXConfig()

        If result Then
            Try
                p_info = myPbxConfig.GetUserByAddress(p_target)
            Catch ex As Exception
                result = False
            End Try

            If Not result Then
                Dim y As String = ""

                p_info = Nothing
                Randomize()

                For i = 0 To p_target.Length - 1
                    Dim a As Integer = Asc(p_target.Chars(i))

                    If (a >= 65 And a <= 90) Or (a >= 97 And a <= 122) Then
                        If Rnd() >= 0.5 Then
                            If a <= 90 Then
                                a = a + 32
                            Else
                                a = a - 32
                            End If

                            y = y & Chr(a)
                        Else
                            y = y & p_target.Chars(i)
                        End If
                    Else
                        y = y & p_target.Chars(i)
                    End If
                Next

                result = True

                Try
                    p_info = myPbxConfig.GetUserByAddress(y)
                Catch ex As Exception
                    result = False
                End Try
            End If
        End If

        Return result
    End Function

    Private Function InitialiseMyPBXConfig() As Boolean
        Dim result As Boolean = True

        myPbxConfig = Nothing

        Try
            myPbxConfig = CreateObject("IpPbxSrv.PBXConfig")
        Catch ex As Exception
            LogError("GetUserByAddressWrapperModule::InitialiseMyPBXConfig() - Could not recreate IpPbxSrv.PBXConfig object", ex.Message)
            myPbxConfig = Nothing
            result = False
        End Try

        Return result
    End Function
End Module
