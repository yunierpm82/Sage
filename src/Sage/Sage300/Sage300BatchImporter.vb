Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports AccpacCOMAPI

Namespace Sage300

    Public Class Sage300ImportResult
        Public Property Success As Boolean
        Public Property Message As String
        Public Property BatchNumber As Integer
    End Class

    ''' Crea un lote (batch) nuevo de Libro Mayor en Sage 300 a partir de un archivo Excel,
    ''' usando la API COM clasica de "Business Logic Views" (AccpacCOMAPI). El lote queda
    ''' sin contabilizar (no se llama a Post) para que se revise en Sage 300 antes de postear.
    '''
    ''' IMPORTANTE - pendiente de verificar contra un login real:
    ''' Los View ID (GL0016/GL0017/GL0018) y los nombres de campo de abajo son los valores
    ''' estandar mas usados en integraciones de Sage 300/Accpac para el modulo de G/L, pero
    ''' no se pudieron probar de punta a punta desde este entorno (no hay credenciales validas
    ''' de Sage 300 aqui). Si algo falla, el mensaje de error mostrado en la pantalla dira
    ''' exactamente que campo o vista no coincide (incluyendo la lista de campos reales de esa
    ''' vista), lo cual sirve para corregir el nombre correcto rapidamente.
    Public Class Sage300BatchImporter

        Private Const ViewBatch As String = "GL0016"
        Private Const ViewJournalEntry As String = "GL0017"
        Private Const ViewJournalEntryDetail As String = "GL0018"

        Private Const AppID As String = "SM"
        Private Const AppVersion As String = "70A"
        Private Const ProgramName As String = "SageBatchImport"

        Public Function ImportBatch(companyId As String, userId As String, password As String, excelPath As String) As Sage300ImportResult
            Dim session As AccpacSessionClass = Nothing

            Try
                Dim entriesByNumber = ExcelTransactionReader.ReadEntries(excelPath).
                    GroupBy(Function(l) l.EntryNumber).
                    OrderBy(Function(g) g.Key).
                    ToList()

                If entriesByNumber.Count = 0 Then
                    Return New Sage300ImportResult With {.Success = False, .Message = "El archivo Excel no contiene transacciones válidas."}
                End If

                session = New AccpacSessionClass()
                session.Init("", AppID, ProgramName, AppVersion)
                session.Open(userId, password, companyId, DateTime.Today, 0, "")

                Dim dbLink = session.OpenDBLink(tagDBLinkTypeEnum.DBLINK_COMPANY, tagDBLinkFlagsEnum.DBLINK_FLG_READWRITE)

                Dim batchView As AccpacView = Nothing
                Dim batchOpenCode = dbLink.OpenView(ViewBatch, batchView)
                If batchOpenCode <> 0 OrElse batchView Is Nothing Then
                    Return New Sage300ImportResult With {.Success = False, .Message = $"No se pudo abrir la vista de lote ({ViewBatch}). Código de retorno: {batchOpenCode}"}
                End If

                batchView.RecordClear()
                SetField(batchView, "DESCRIPTION", $"Importado desde Excel {Path.GetFileName(excelPath)}")
                SetField(batchView, "SOURCE", "GL")
                batchView.Insert()

                Dim batchNumber = CInt(GetField(batchView, "BATCH"))

                Dim entryView As AccpacView = Nothing
                Dim entryOpenCode = dbLink.OpenView(ViewJournalEntry, entryView)
                If entryOpenCode <> 0 OrElse entryView Is Nothing Then
                    Return New Sage300ImportResult With {.Success = False, .Message = $"No se pudo abrir la vista de asiento ({ViewJournalEntry}). Código de retorno: {entryOpenCode}"}
                End If

                Dim detailView As AccpacView = Nothing
                Dim detailOpenCode = dbLink.OpenView(ViewJournalEntryDetail, detailView)
                If detailOpenCode <> 0 OrElse detailView Is Nothing Then
                    Return New Sage300ImportResult With {.Success = False, .Message = $"No se pudo abrir la vista de detalle de asiento ({ViewJournalEntryDetail}). Código de retorno: {detailOpenCode}"}
                End If

                For Each entryGroup In entriesByNumber
                    Dim firstLine = entryGroup.First()

                    entryView.RecordClear()
                    SetField(entryView, "BATCH", batchNumber)
                    SetField(entryView, "DESCRIPTION", firstLine.Description)
                    SetField(entryView, "ENTRYDATE", firstLine.EntryDate)
                    SetField(entryView, "SOURCELEDGER", "GL")
                    SetField(entryView, "SOURCETYPE", "GL-GE")
                    entryView.Insert()

                    Dim entryNumber = CInt(GetField(entryView, "ENTRY"))

                    For Each line In entryGroup
                        detailView.RecordClear()
                        SetField(detailView, "BATCH", batchNumber)
                        SetField(detailView, "ENTRY", entryNumber)
                        SetField(detailView, "ACCTID", line.Account)
                        SetField(detailView, "DESCRIPTION", line.Description)
                        SetField(detailView, "REFERENCE", line.Reference)
                        SetField(detailView, "AMOUNT", If(line.Debit <> 0, line.Debit, -line.Credit))
                        detailView.Insert()
                    Next
                Next

                Return New Sage300ImportResult With {
                    .Success = True,
                    .Message = $"Importación exitosa. Se creó el lote (batch) número {batchNumber} con {entriesByNumber.Count} asiento(s), sin contabilizar. Revísalo en Sage 300 antes de postearlo.",
                    .BatchNumber = batchNumber
                }

            Catch ex As Exception
                Return New Sage300ImportResult With {.Success = False, .Message = ex.Message}
            Finally
                Try
                    If session IsNot Nothing Then session.Close()
                Catch
                End Try
            End Try
        End Function

        Private Sub SetField(view As AccpacView, fieldName As String, value As Object)
            Dim field As AccpacViewField
            Try
                field = view.Fields.FieldByName(fieldName)
            Catch
                Throw New Exception($"El campo '{fieldName}' no existe en la vista {view.ViewID}. Campos disponibles: {String.Join(", ", GetFieldNames(view))}")
            End Try
            field.Value = value
        End Sub

        Private Function GetField(view As AccpacView, fieldName As String) As Object
            Dim field As AccpacViewField
            Try
                field = view.Fields.FieldByName(fieldName)
            Catch
                Throw New Exception($"El campo '{fieldName}' no existe en la vista {view.ViewID}. Campos disponibles: {String.Join(", ", GetFieldNames(view))}")
            End Try
            Return field.Value
        End Function

        Private Function GetFieldNames(view As AccpacView) As List(Of String)
            Dim names As New List(Of String)
            For i = 0 To view.Fields.Count - 1
                names.Add(view.Fields.FieldByIndex(i).Name)
            Next
            Return names
        End Function

    End Class

End Namespace
