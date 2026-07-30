Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Windows.Forms

Namespace Sage300

    ''' Self-contained "link required columns to a sample Excel's columns" screen, used by GL
    ''' Booking > Relate Columns and AP Booking > Configure Columns (and any future module that
    ''' needs the same kind of column mapping). Each instance keeps its own templates under the
    ''' given "category" (Nothing for GL, to preserve the original storage location; e.g. "AP"
    ''' for Accounts Payable) so the two modules' saved templates never collide.
    Public Class ColumnMappingPanel
        Inherits Panel

        Private ReadOnly colorBackground As Color = Color.FromArgb(32, 32, 36)
        Private ReadOnly colorPanel As Color = Color.FromArgb(45, 45, 48)
        Private ReadOnly colorText As Color = Color.FromArgb(230, 230, 230)
        Private ReadOnly colorAccent As Color = Color.FromArgb(0, 122, 204)
        Private ReadOnly colorButton As Color = Color.FromArgb(60, 60, 65)
        Private ReadOnly colorBorder As Color = Color.FromArgb(70, 70, 75)
        Private ReadOnly colorSuccess As Color = Color.FromArgb(120, 220, 120)
        Private ReadOnly colorError As Color = Color.FromArgb(240, 120, 120)
        Private ReadOnly colorInfo As Color = Color.FromArgb(120, 180, 240)

        Private ReadOnly category As String
        Private ReadOnly requiredColumnKeys As String()

        Private cboProfile As ComboBox
        Private btnSaveProfile As Button
        Private btnDeleteProfile As Button
        Private btnChooseSample As Button
        Private lblSampleFile As Label
        Private lstRequiredColumns As ListBox
        Private lstActualColumns As ListBox
        Private pnlLinks As Panel
        Private lblStatus As Label

        Private currentColumnMapping As New Dictionary(Of String, String)
        Private armedRequiredColumn As String = Nothing
        Private currentSamplePath As String = Nothing
        Private hasLoadedInitialTemplate As Boolean = False

        Public Sub New(category As String, requiredColumnKeys As String())
            Me.category = category
            Me.requiredColumnKeys = requiredColumnKeys
            Me.Dock = DockStyle.Fill
            Me.BackColor = colorBackground
            Me.Visible = False
            BuildUI()
        End Sub

        ''' Call when the menu item that shows this panel is clicked; refreshes the template list
        ''' and, the first time only, restores the last template used for this category.
        Public Sub ActivateFirstLoad()
            RefreshTemplateList()

            If Not hasLoadedInitialTemplate Then
                hasLoadedInitialTemplate = True
                Dim lastTemplate = AppSettingsStore.GetLastTemplate(category)
                If Not String.IsNullOrWhiteSpace(lastTemplate) AndAlso cboProfile.Items.Contains(lastTemplate) Then
                    cboProfile.Text = lastTemplate
                    LoadTemplate(lastTemplate)
                End If
            End If
        End Sub

        Private Sub BuildUI()
            Dim lblProfile As New Label With {.Text = "Template:", .Location = New Point(20, 18), .AutoSize = True, .ForeColor = colorText}
            cboProfile = New ComboBox With {
                .Location = New Point(100, 15),
                .Width = 220,
                .DropDownStyle = ComboBoxStyle.DropDown,
                .BackColor = colorPanel,
                .ForeColor = colorText
            }
            AddHandler cboProfile.SelectedIndexChanged, Sub() LoadTemplate(cboProfile.Text)

            btnSaveProfile = New Button With {.Text = "Save", .Location = New Point(330, 14), .Width = 100, .Height = 26}
            StylePrimaryButton(btnSaveProfile)
            AddHandler btnSaveProfile.Click, AddressOf BtnSaveProfile_Click

            btnDeleteProfile = New Button With {.Text = "Delete", .Location = New Point(440, 14), .Width = 100, .Height = 26}
            StyleSecondaryButton(btnDeleteProfile)
            AddHandler btnDeleteProfile.Click, AddressOf BtnDeleteProfile_Click

            btnChooseSample = New Button With {.Text = "Choose sample Excel...", .Location = New Point(20, 55), .Width = 220, .Height = 26}
            StyleSecondaryButton(btnChooseSample)
            AddHandler btnChooseSample.Click, AddressOf BtnChooseSample_Click

            lblSampleFile = New Label With {
                .Text = "(no file selected)",
                .Location = New Point(250, 59),
                .Size = New Size(500, 20),
                .ForeColor = colorText,
                .AutoEllipsis = True
            }

            Dim lblRequiredHeader As New Label With {.Text = "Required columns", .Location = New Point(20, 95), .AutoSize = True, .ForeColor = colorText}
            Dim lblActualHeader As New Label With {.Text = "Excel columns", .Location = New Point(410, 95), .AutoSize = True, .ForeColor = colorText}
            Dim lblHint As New Label With {
                .Text = "Click a column on the left, then the matching one on the right to link them. Double-click a column on the left to remove its link.",
                .Location = New Point(20, 380),
                .Size = New Size(650, 40),
                .ForeColor = colorText
            }

            lstRequiredColumns = New ListBox With {
                .Location = New Point(20, 120),
                .Size = New Size(220, 250),
                .BackColor = colorPanel,
                .ForeColor = colorText,
                .BorderStyle = BorderStyle.FixedSingle,
                .ItemHeight = 22,
                .IntegralHeight = False
            }
            lstRequiredColumns.Items.AddRange(requiredColumnKeys)
            AddHandler lstRequiredColumns.SelectedIndexChanged, AddressOf LstRequiredColumns_SelectedIndexChanged
            AddHandler lstRequiredColumns.DoubleClick, AddressOf LstRequiredColumns_DoubleClick

            pnlLinks = New Panel With {
                .Location = New Point(250, 120),
                .Size = New Size(150, 250),
                .BackColor = colorBackground
            }
            AddHandler pnlLinks.Paint, AddressOf PnlLinks_Paint

            lstActualColumns = New ListBox With {
                .Location = New Point(410, 120),
                .Size = New Size(260, 250),
                .BackColor = colorPanel,
                .ForeColor = colorText,
                .BorderStyle = BorderStyle.FixedSingle,
                .ItemHeight = 22,
                .IntegralHeight = False
            }
            AddHandler lstActualColumns.SelectedIndexChanged, AddressOf LstActualColumns_SelectedIndexChanged

            lblStatus = New Label With {
                .Location = New Point(20, 425),
                .Size = New Size(650, 40),
                .ForeColor = colorInfo,
                .Text = ""
            }

            Me.Controls.AddRange(New Control() {
                lblProfile, cboProfile, btnSaveProfile, btnDeleteProfile,
                btnChooseSample, lblSampleFile,
                lblRequiredHeader, lblActualHeader,
                lstRequiredColumns, pnlLinks, lstActualColumns,
                lblHint,
                lblStatus
            })
        End Sub

        Private Sub StylePrimaryButton(button As Button)
            button.FlatStyle = FlatStyle.Flat
            button.BackColor = colorAccent
            button.ForeColor = Color.White
            button.FlatAppearance.BorderSize = 0
        End Sub

        Private Sub StyleSecondaryButton(button As Button)
            button.FlatStyle = FlatStyle.Flat
            button.BackColor = colorButton
            button.ForeColor = colorText
            button.FlatAppearance.BorderSize = 1
            button.FlatAppearance.BorderColor = colorBorder
        End Sub

        Private Sub RefreshTemplateList()
            Dim currentText = cboProfile.Text
            cboProfile.Items.Clear()
            For Each profileName In ColumnMappingStore.ListProfiles(category)
                cboProfile.Items.Add(profileName)
            Next
            If cboProfile.Items.Contains(currentText) Then cboProfile.Text = currentText
        End Sub

        Private Sub LstRequiredColumns_SelectedIndexChanged(sender As Object, e As EventArgs)
            If lstRequiredColumns.SelectedItem Is Nothing Then Return
            armedRequiredColumn = lstRequiredColumns.SelectedItem.ToString()
            lblStatus.ForeColor = colorInfo
            lblStatus.Text = $"Now select the Excel column that corresponds to '{armedRequiredColumn}'."
        End Sub

        Private Sub LstRequiredColumns_DoubleClick(sender As Object, e As EventArgs)
            If lstRequiredColumns.SelectedItem Is Nothing Then Return
            Dim key = lstRequiredColumns.SelectedItem.ToString()
            If currentColumnMapping.ContainsKey(key) Then
                currentColumnMapping.Remove(key)
                lblStatus.ForeColor = colorInfo
                lblStatus.Text = $"Removed the link for '{key}'."
                pnlLinks.Invalidate()
            End If
            armedRequiredColumn = Nothing
        End Sub

        Private Sub LstActualColumns_SelectedIndexChanged(sender As Object, e As EventArgs)
            If lstActualColumns.SelectedItem Is Nothing OrElse armedRequiredColumn Is Nothing Then Return

            Dim actualColumn = lstActualColumns.SelectedItem.ToString()

            Dim existingKey = currentColumnMapping.
                Where(Function(kv) kv.Value = actualColumn).
                Select(Function(kv) kv.Key).
                FirstOrDefault()
            If Not String.IsNullOrEmpty(existingKey) Then currentColumnMapping.Remove(existingKey)

            currentColumnMapping(armedRequiredColumn) = actualColumn

            lblStatus.ForeColor = colorSuccess
            lblStatus.Text = $"'{armedRequiredColumn}' → '{actualColumn}'"

            armedRequiredColumn = Nothing
            lstRequiredColumns.ClearSelected()
            lstActualColumns.ClearSelected()
            pnlLinks.Invalidate()
        End Sub

        Private Sub PnlLinks_Paint(sender As Object, e As PaintEventArgs)
            e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias

            For Each mapping In currentColumnMapping
                Dim leftIndex = lstRequiredColumns.Items.IndexOf(mapping.Key)
                Dim rightIndex = lstActualColumns.Items.IndexOf(mapping.Value)
                If leftIndex < 0 OrElse rightIndex < 0 Then Continue For

                Dim yLeft = leftIndex * lstRequiredColumns.ItemHeight + lstRequiredColumns.ItemHeight \ 2
                Dim yRight = rightIndex * lstActualColumns.ItemHeight + lstActualColumns.ItemHeight \ 2

                Using pen As New Pen(colorAccent, 2)
                    e.Graphics.DrawLine(pen, 0, yLeft, pnlLinks.Width, yRight)
                End Using
                Using dotBrush As New SolidBrush(colorAccent)
                    e.Graphics.FillEllipse(dotBrush, -3, yLeft - 3, 6, 6)
                    e.Graphics.FillEllipse(dotBrush, pnlLinks.Width - 3, yRight - 3, 6, 6)
                End Using
            Next
        End Sub

        Private Sub BtnChooseSample_Click(sender As Object, e As EventArgs)
            Using dialog As New OpenFileDialog With {
                .Filter = "Excel files (*.xlsx)|*.xlsx",
                .Title = "Select a sample Excel from the third-party application"
            }
                If dialog.ShowDialog() = DialogResult.OK Then
                    Try
                        Dim headers = ExcelTransactionReader.ReadHeaders(dialog.FileName)
                        lstActualColumns.Items.Clear()
                        lstActualColumns.Items.AddRange(headers.ToArray())
                        lblSampleFile.Text = Path.GetFileName(dialog.FileName)
                        currentSamplePath = dialog.FileName
                        lblStatus.ForeColor = colorSuccess
                        lblStatus.Text = $"Found {headers.Count} column(s) in the file."
                        pnlLinks.Invalidate()
                    Catch ex As Exception
                        lblStatus.ForeColor = colorError
                        lblStatus.Text = "Error reading the Excel file: " & ex.Message
                    End Try
                End If
            End Using
        End Sub

        Private Sub LoadTemplate(templateName As String)
            If String.IsNullOrWhiteSpace(templateName) Then Return

            currentColumnMapping = ColumnMappingStore.Load(templateName, category)

            Dim samplePath = TemplateSampleStore.TryGetExistingSamplePath(templateName, category)
            If samplePath IsNot Nothing Then
                Try
                    Dim headers = ExcelTransactionReader.ReadHeaders(samplePath)
                    lstActualColumns.Items.Clear()
                    lstActualColumns.Items.AddRange(headers.ToArray())
                    lblSampleFile.Text = Path.GetFileName(samplePath)
                    currentSamplePath = samplePath
                Catch
                    ' Ignore a stale/corrupted sample copy; the mapping itself still loads fine.
                End Try
            End If

            AppSettingsStore.SetLastTemplate(templateName, category)

            lblStatus.ForeColor = colorInfo
            lblStatus.Text = $"Template '{templateName}' loaded ({currentColumnMapping.Count} link(s))."
            pnlLinks.Invalidate()
        End Sub

        Private Sub BtnSaveProfile_Click(sender As Object, e As EventArgs)
            Dim name = cboProfile.Text.Trim()
            If String.IsNullOrWhiteSpace(name) Then
                lblStatus.ForeColor = colorError
                lblStatus.Text = "Type a name for the template before saving."
                Return
            End If

            ColumnMappingStore.Save(name, currentColumnMapping, category)

            If Not String.IsNullOrWhiteSpace(currentSamplePath) AndAlso File.Exists(currentSamplePath) Then
                Try
                    TemplateSampleStore.SaveSample(name, currentSamplePath, category)
                Catch ex As Exception
                    RefreshTemplateList()
                    cboProfile.Text = name
                    lblStatus.ForeColor = colorError
                    lblStatus.Text = $"Template saved, but the sample file could not be copied: {ex.Message}"
                    Return
                End Try
            End If

            AppSettingsStore.SetLastTemplate(name, category)

            RefreshTemplateList()
            cboProfile.Text = name
            lblStatus.ForeColor = colorSuccess
            lblStatus.Text = $"Template '{name}' saved with {currentColumnMapping.Count} link(s)."
        End Sub

        Private Sub BtnDeleteProfile_Click(sender As Object, e As EventArgs)
            Dim name = cboProfile.Text.Trim()
            If String.IsNullOrWhiteSpace(name) Then Return

            ColumnMappingStore.Delete(name, category)
            TemplateSampleStore.DeleteSample(name, category)

            If AppSettingsStore.GetLastTemplate(category) = name Then
                AppSettingsStore.SetLastTemplate("", category)
            End If

            RefreshTemplateList()
            cboProfile.Text = ""
            currentColumnMapping = New Dictionary(Of String, String)
            currentSamplePath = Nothing
            lstActualColumns.Items.Clear()
            lblSampleFile.Text = "(no file selected)"
            pnlLinks.Invalidate()
            lblStatus.ForeColor = colorInfo
            lblStatus.Text = $"Template '{name}' deleted."
        End Sub

    End Class

End Namespace
