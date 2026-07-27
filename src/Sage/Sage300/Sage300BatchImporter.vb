Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq

Namespace Sage300

    Public Class Sage300ImportResult
        Public Property Success As Boolean
        Public Property Message As String
        Public Property BatchNumber As String
    End Class

    ''' Crea un lote (batch) nuevo de Libro Mayor en Sage 300 a partir de un archivo Excel,
    ''' usando el objeto COM "ACCPAC.xapiSession" (late binding, sin libreria de interop) --
    ''' el mismo objeto que usan las aplicaciones de Access del usuario y que SI funciona en
    ''' esta instalacion (a diferencia de AccpacCOMAPI.AccpacSessionClass, que fallaba con un
    ''' error de licencia).
    '''
    ''' Los View ID (GL0006 = asiento/"header", GL0008 = batch, GL0010 = detalle) y los nombres
    ''' de campo, asi como la secuencia Compose/Init/Update/Insert, fueron tomados directamente
    ''' de una macro de Access del usuario que ya funciona contra esta misma base de datos.
    ''' El lote queda sin contabilizar (no se llama Post) para revisarlo en Sage 300 antes de postear.
    Public Class Sage300BatchImporter

        Private Const ViewJournalEntry As String = "GL0006"
        Private Const ViewBatch As String = "GL0008"
        Private Const ViewDetail As String = "GL0010"
        Private Const ApplicationId As String = "GL"

        ' Codigo de fuente (Source Code) para el ledger GL, registrado en Common Services > Source Codes.
        Private Const SourceLedger As String = "GL"
        Private Const SourceType As String = "AR"

        Public Function ImportBatch(companyId As String, userId As String, password As String, excelPath As String) As Sage300ImportResult
            Dim session As Object = Nothing

            Try
                Dim entriesByNumber = ExcelTransactionReader.ReadEntries(excelPath).
                    GroupBy(Function(l) l.EntryNumber).
                    OrderBy(Function(g) g.Key).
                    ToList()

                If entriesByNumber.Count = 0 Then
                    Return New Sage300ImportResult With {.Success = False, .Message = "El archivo Excel no contiene transacciones válidas."}
                End If

                Dim sessionType = Type.GetTypeFromProgID("ACCPAC.xapiSession")
                If sessionType Is Nothing Then
                    Return New Sage300ImportResult With {.Success = False, .Message = "No se encontró el componente 'ACCPAC.xapiSession' en esta máquina."}
                End If

                session = Activator.CreateInstance(sessionType)
                session.Open(userId, password, companyId, DateTime.Today, 0)

                Dim header As Object = session.OpenView(ViewJournalEntry, ApplicationId)
                Dim headerFields As Object = header.Fields

                Dim batchView As Object = session.OpenView(ViewBatch, ApplicationId)
                Dim batchFields As Object = batchView.Fields

                Dim detailView As Object = session.OpenView(ViewDetail, ApplicationId)
                Dim detailFields As Object = detailView.Fields

                batchView.Compose(New Object() {header})
                header.Compose(New Object() {batchView, detailView})
                detailView.Compose(New Object() {header})

                ' Crear el batch (secuencia tal cual la usa la macro de Access de referencia).
                batchView.Init()
                header.Fetch()
                headerFields("BTCHENTRY").PutWithoutVerification("00000")

                header.Init()
                batchFields("BTCHDESC").Value = $"Importado desde Excel {Path.GetFileName(excelPath)}"
                batchView.Update()

                Dim isFirstEntry = True

                For Each entryGroup In entriesByNumber
                    Dim firstLine = entryGroup.First()

                    If Not isFirstEntry Then
                        header.Init()
                    End If
                    isFirstEntry = False

                    headerFields("SRCELEDGER").Value = SourceLedger
                    headerFields("SRCETYPE").Value = SourceType
                    headerFields("FSCSPERD").Value = firstLine.EntryDate.Month
                    headerFields("DATEENTRY").Value = DateTime.Today
                    headerFields("JRNLDESC").Value = firstLine.Description

                    For Each line In entryGroup
                        detailView.Init()
                        detailFields("ACCTID").Value = line.Account
                        detailFields("TRANSDESC").Value = line.Description
                        detailFields("TRANSREF").Value = line.Reference
                        detailFields("SCURNAMT").Value = line.Amount
                        detailFields("TRANSDATE").Value = line.EntryDate
                        detailView.Insert()
                    Next

                    header.Insert()
                Next

                Dim batchNumber = Convert.ToString(batchFields("BATCHID").Value)

                Return New Sage300ImportResult With {
                    .Success = True,
                    .Message = $"Importación exitosa. Se creó el lote (batch) número {batchNumber} con {entriesByNumber.Count} asiento(s), sin contabilizar. Revísalo en Sage 300 antes de postearlo.",
                    .BatchNumber = batchNumber
                }

            Catch ex As Exception
                Dim detail = TryGetAccpacErrors(session)
                Return New Sage300ImportResult With {.Success = False, .Message = If(detail, ex.Message)}
            Finally
                Try
                    If session IsNot Nothing Then session.Close()
                Catch
                End Try
            End Try
        End Function

        Private Function TryGetAccpacErrors(session As Object) As String
            Try
                If session Is Nothing Then Return Nothing

                Dim errors As Object = session.errors
                If errors Is Nothing Then Return Nothing

                Dim count As Integer = CInt(errors.Count)
                If count = 0 Then Return Nothing

                Dim messages As New List(Of String)
                For Each errorItem In DirectCast(errors, IEnumerable)
                    messages.Add(Convert.ToString(errorItem.Description))
                Next

                errors.Clear()
                Return String.Join(" | ", messages)
            Catch
                Return Nothing
            End Try
        End Function

    End Class

End Namespace
