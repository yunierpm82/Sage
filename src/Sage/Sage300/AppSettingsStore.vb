Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Text.Json

Namespace Sage300

    ''' Small persistent app setting (last column-mapping template used per module), stored per user.
    Public Class AppSettingsStore

        Private Shared Function GetSettingsFilePath() As String
            Dim folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Sage")
            Directory.CreateDirectory(folder)
            Return Path.Combine(folder, "settings.json")
        End Function

        ''' "category" is Nothing for GL Booking (kept as the original "LastTemplate" key so existing
        ''' installs don't lose their setting) or e.g. "AP" for other modules.
        Private Shared Function GetKey(category As String) As String
            Return If(String.IsNullOrEmpty(category), "LastTemplate", "LastTemplate_" & category)
        End Function

        Public Shared Function GetLastTemplate(Optional category As String = Nothing) As String
            Try
                Dim filePath = GetSettingsFilePath()
                If Not File.Exists(filePath) Then Return Nothing

                Dim json = File.ReadAllText(filePath)
                Dim data = JsonSerializer.Deserialize(Of Dictionary(Of String, String))(json)
                Dim key = GetKey(category)
                If data IsNot Nothing AndAlso data.ContainsKey(key) Then Return data(key)
                Return Nothing
            Catch
                Return Nothing
            End Try
        End Function

        Public Shared Sub SetLastTemplate(templateName As String, Optional category As String = Nothing)
            Try
                Dim filePath = GetSettingsFilePath()
                Dim data As Dictionary(Of String, String) = Nothing
                If File.Exists(filePath) Then
                    Try
                        data = JsonSerializer.Deserialize(Of Dictionary(Of String, String))(File.ReadAllText(filePath))
                    Catch
                    End Try
                End If
                If data Is Nothing Then data = New Dictionary(Of String, String)

                data(GetKey(category)) = templateName
                File.WriteAllText(filePath, JsonSerializer.Serialize(data))
            Catch
            End Try
        End Sub

    End Class

End Namespace
