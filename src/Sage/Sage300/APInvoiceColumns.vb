Namespace Sage300

    ''' Minimum columns needed to import an A/P Invoice, taken directly from Sage 300's own
    ''' official import template (ImportTemplates\APInvoice1.xlsx, sheets "Invoices",
    ''' "Invoice_Details" and "Invoice_Payment_Schedules" -- the Optional Fields sheets are
    ''' excluded since Sage's own template says they can be left blank).
    Public Class APInvoiceColumns

        Public Shared ReadOnly RequiredColumnKeys As String() = {
            "CNTBTCH", "CNTITEM", "IDVEND", "IDINVC", "TEXTTRX", "IDACCTSET", "DATEDUE", "AMTGROSTOT",
            "CNTLINE", "IDGLACCT", "AMTTOTTAX", "AMTDIST", "RATETAX1", "RATETAX2",
            "CNTPAYM", "AMTDUE"
        }

    End Class

End Namespace
