Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Text.Json

Namespace Sage300

    ''' Small persistent app setting (last column-mapping template used), stored per user.
    Public Class AppSettingsStore

        Private Shared Function GetSettingsFilePath() As String
            Dim folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Sage")
            Directory.CreateDirectory(folder)
            Return Path.Combine(folder, "settings.json")
        End Function

        Public Shared Function GetLastTemplate() As String
            Try
                Dim filePath = GetSettingsFilePath()
                If Not File.Exists(filePath) Then Return Nothing

                Dim json = File.ReadAllText(filePath)
                Dim data = JsonSerializer.Deserialize(Of Dictionary(Of String, String))(json)
                If data IsNot Nothing AndAlso data.ContainsKey("LastTemplate") Then Return data("LastTemplate")
                Return Nothing
            Catch
                Return Nothing
            End Try
        End Function

        Public Shared Sub SetLastTemplate(templateName As String)
            Try
                Dim data As New Dictionary(Of String, String) From {{"LastTemplate", templateName}}
                File.WriteAllText(GetSettingsFilePath(), JsonSerializer.Serialize(data))
            Catch
            End Try
        End Sub

    End Class

End Namespace
