Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text.Json

Namespace Sage300

    ''' Saves/loads column-mapping templates (required column -> real column from a third-party
    ''' application's Excel), one JSON file per template, in the user's application data folder.
    ''' An optional "category" (e.g. "AP") keeps each module's templates in their own subfolder so
    ''' names don't collide; the default (Nothing) keeps the original GL Booking location.
    Public Class ColumnMappingStore

        Private Shared Function GetFolder(Optional category As String = Nothing) As String
            Dim baseFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Sage", "ColumnMappings")
            Dim folder = If(String.IsNullOrEmpty(category), baseFolder, Path.Combine(baseFolder, category))
            Directory.CreateDirectory(folder)
            Return folder
        End Function

        Public Shared Function ListProfiles(Optional category As String = Nothing) As List(Of String)
            Return Directory.GetFiles(GetFolder(category), "*.json").
                Select(Function(f) Path.GetFileNameWithoutExtension(f)).
                OrderBy(Function(n) n, StringComparer.OrdinalIgnoreCase).
                ToList()
        End Function

        Public Shared Function Load(profileName As String, Optional category As String = Nothing) As Dictionary(Of String, String)
            Dim filePath = Path.Combine(GetFolder(category), profileName & ".json")
            If Not File.Exists(filePath) Then Return New Dictionary(Of String, String)
            Dim json = File.ReadAllText(filePath)
            Return JsonSerializer.Deserialize(Of Dictionary(Of String, String))(json)
        End Function

        Public Shared Sub Save(profileName As String, mapping As Dictionary(Of String, String), Optional category As String = Nothing)
            Dim filePath = Path.Combine(GetFolder(category), profileName & ".json")
            Dim json = JsonSerializer.Serialize(mapping, New JsonSerializerOptions With {.WriteIndented = True})
            File.WriteAllText(filePath, json)
        End Sub

        Public Shared Sub Delete(profileName As String, Optional category As String = Nothing)
            Dim filePath = Path.Combine(GetFolder(category), profileName & ".json")
            If File.Exists(filePath) Then File.Delete(filePath)
        End Sub

    End Class

End Namespace
