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

    ''' Creates a new General Ledger batch in Sage 300 from an Excel file, using the
    ''' "ACCPAC.xapiSession" COM object (late binding, no interop library needed) -- the same
    ''' object used by the user's Access applications, which does work on this installation
    ''' (unlike AccpacCOMAPI.AccpacSessionClass, which failed with a licensing error).
    '''
    ''' The view IDs (GL0006 = journal entry/"header", GL0008 = batch, GL0010 = detail) and
    ''' field names, as well as the Compose/Init/Update/Insert sequence, were taken directly
    ''' from an Access macro of the user's that already works against this same database.
    ''' The batch is left unposted (Post is never called) so it can be reviewed in Sage 300
    ''' before posting.
    Public Class Sage300BatchImporter

        Private Const ViewJournalEntry As String = "GL0006"
        Private Const ViewBatch As String = "GL0008"
        Private Const ViewDetail As String = "GL0010"
        Private Const ApplicationId As String = "GL"

        ' Source Code for the GL ledger, registered in Common Services > Source Codes.
        Private Const SourceLedger As String = "GL"
        Private Const SourceType As String = "AR"

        Public Function ImportBatch(companyId As String, userId As String, password As String, excelPath As String, Optional columnMapping As Dictionary(Of String, String) = Nothing) As Sage300ImportResult
            Dim session As Object = Nothing

            Try
                Dim entriesByNumber = ExcelTransactionReader.ReadEntries(excelPath, columnMapping).
                    GroupBy(Function(l) l.EntryNumber).
                    OrderBy(Function(g) g.Key).
                    ToList()

                If entriesByNumber.Count = 0 Then
                    Return New Sage300ImportResult With {.Success = False, .Message = "The Excel file does not contain valid transactions."}
                End If

                Dim sessionType = Type.GetTypeFromProgID("ACCPAC.xapiSession")
                If sessionType Is Nothing Then
                    Return New Sage300ImportResult With {.Success = False, .Message = "Could not find the 'ACCPAC.xapiSession' component on this machine."}
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

                ' Create the batch (sequence exactly as used by the reference Access macro).
                batchView.Init()
                header.Fetch()
                headerFields("BTCHENTRY").PutWithoutVerification("00000")

                header.Init()
                batchFields("BTCHDESC").Value = $"Imported from Excel {Path.GetFileName(excelPath)}"
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
                    .Message = $"Import successful. Created batch number {batchNumber} with {entriesByNumber.Count} entry/entries, not yet posted. Review it in Sage 300 before posting.",
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
