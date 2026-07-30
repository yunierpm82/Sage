Imports System
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms

Namespace Sage300

    ''' One full AP Booking screen (Invoice Entry, Invoice Batch List, Payment Entry or Payment
    ''' Batch List): the column-mapping editor (ColumnMappingPanel) on top, and the "connect to
    ''' Sage" fields (Company, Sage user, Password, Excel file, Template) below it. Each screen's
    ''' templates are stored under ColumnMappings\AP\<moduleName> / Plantillas\AP\<moduleName>, so
    ''' the four screens never share templates with each other.
    Public Class APModulePanel
        Inherits Panel

        Private ReadOnly colorBackground As Color = Color.FromArgb(32, 32, 36)
        Private ReadOnly colorPanel As Color = Color.FromArgb(45, 45, 48)
        Private ReadOnly colorText As Color = Color.FromArgb(230, 230, 230)
        Private ReadOnly colorButton As Color = Color.FromArgb(60, 60, 65)
        Private ReadOnly colorBorder As Color = Color.FromArgb(70, 70, 75)
        Private ReadOnly NoTemplateOption As String = "(No template — exact names)"

        Private ReadOnly category As String
        Private ReadOnly columnMapping As ColumnMappingPanel

        Private txtCompany As TextBox
        Private txtUser As TextBox
        Private txtPassword As TextBox
        Private txtExcelPath As TextBox
        Private btnBrowseExcel As Button
        Private cboTemplate As ComboBox

        Private hasLoadedInitialTemplate As Boolean = False

        Public Sub New(moduleName As String, requiredColumnKeys As String())
            category = Path.Combine("AP", moduleName)

            Me.Dock = DockStyle.Fill
            Me.BackColor = colorBackground
            Me.AutoScroll = True
            Me.Visible = False

            columnMapping = New ColumnMappingPanel(category, requiredColumnKeys) With {
                .Location = New Point(0, 0)
            }

            BuildConnectSection()

            Me.Controls.Add(columnMapping)
        End Sub

        Private Sub BuildConnectSection()
            Const top As Integer = 490

            Dim lblSectionTitle As New Label With {
                .Text = "Connect to Sage",
                .Location = New Point(20, top),
                .AutoSize = True,
                .Font = New Font(Me.Font, FontStyle.Bold),
                .ForeColor = colorText
            }

            Dim lblCompany As New Label With {.Text = "Company:", .Location = New Point(20, top + 32), .AutoSize = True, .ForeColor = colorText}
            txtCompany = New TextBox With {.Location = New Point(160, top + 29), .Width = 300, .BackColor = colorPanel, .ForeColor = colorText, .BorderStyle = BorderStyle.FixedSingle}

            Dim lblUser As New Label With {.Text = "Sage user:", .Location = New Point(20, top + 67), .AutoSize = True, .ForeColor = colorText}
            txtUser = New TextBox With {.Location = New Point(160, top + 64), .Width = 300, .BackColor = colorPanel, .ForeColor = colorText, .BorderStyle = BorderStyle.FixedSingle}

            Dim lblPassword As New Label With {.Text = "Password:", .Location = New Point(20, top + 102), .AutoSize = True, .ForeColor = colorText}
            txtPassword = New TextBox With {.Location = New Point(160, top + 99), .Width = 300, .PasswordChar = "*"c, .BackColor = colorPanel, .ForeColor = colorText, .BorderStyle = BorderStyle.FixedSingle}

            Dim lblExcelPath As New Label With {.Text = "Excel file (.xlsx):", .Location = New Point(20, top + 137), .AutoSize = True, .ForeColor = colorText}
            txtExcelPath = New TextBox With {
                .Location = New Point(160, top + 134),
                .Width = 380,
                .ReadOnly = True,
                .BackColor = colorPanel,
                .ForeColor = colorText,
                .BorderStyle = BorderStyle.FixedSingle
            }
            btnBrowseExcel = New Button With {.Text = "Choose...", .Location = New Point(548, top + 131), .Width = 90, .Height = 30}
            StyleSecondaryButton(btnBrowseExcel)
            AddHandler btnBrowseExcel.Click, AddressOf BtnBrowseExcel_Click

            Dim lblTemplate As New Label With {.Text = "Template:", .Location = New Point(20, top + 172), .AutoSize = True, .ForeColor = colorText}
            cboTemplate = New ComboBox With {
                .Location = New Point(160, top + 169),
                .Width = 300,
                .DropDownStyle = ComboBoxStyle.DropDownList,
                .BackColor = colorPanel,
                .ForeColor = colorText
            }
            AddHandler cboTemplate.SelectedIndexChanged, AddressOf CboTemplate_SelectedIndexChanged

            Me.Controls.AddRange(New Control() {
                lblSectionTitle,
                lblCompany, txtCompany,
                lblUser, txtUser,
                lblPassword, txtPassword,
                lblExcelPath, txtExcelPath, btnBrowseExcel,
                lblTemplate, cboTemplate
            })
        End Sub

        Private Sub StyleSecondaryButton(button As Button)
            button.FlatStyle = FlatStyle.Flat
            button.BackColor = colorButton
            button.ForeColor = colorText
            button.FlatAppearance.BorderSize = 1
            button.FlatAppearance.BorderColor = colorBorder
        End Sub

        ''' Call when the menu item that shows this panel is clicked.
        Public Sub ActivateFirstLoad()
            columnMapping.ActivateFirstLoad()
            RefreshTemplateCombo()

            If Not hasLoadedInitialTemplate Then
                hasLoadedInitialTemplate = True
                Dim lastTemplate = AppSettingsStore.GetLastTemplate(category)
                If Not String.IsNullOrWhiteSpace(lastTemplate) AndAlso cboTemplate.Items.Contains(lastTemplate) Then
                    cboTemplate.Text = lastTemplate
                End If
            End If
        End Sub

        Private Sub RefreshTemplateCombo()
            Dim currentText = cboTemplate.Text
            cboTemplate.Items.Clear()
            cboTemplate.Items.Add(NoTemplateOption)
            For Each profileName In ColumnMappingStore.ListProfiles(category)
                cboTemplate.Items.Add(profileName)
            Next
            If cboTemplate.Items.Contains(currentText) Then
                cboTemplate.Text = currentText
            Else
                cboTemplate.SelectedIndex = 0
            End If
        End Sub

        Private Sub CboTemplate_SelectedIndexChanged(sender As Object, e As EventArgs)
            If cboTemplate.Text = "" OrElse cboTemplate.Text = NoTemplateOption Then Return
            AppSettingsStore.SetLastTemplate(cboTemplate.Text, category)
        End Sub

        Private Sub BtnBrowseExcel_Click(sender As Object, e As EventArgs)
            Using dialog As New OpenFileDialog With {
                .Filter = "Excel files (*.xlsx)|*.xlsx",
                .Title = "Select the file to import"
            }
                If dialog.ShowDialog() = DialogResult.OK Then
                    txtExcelPath.Text = dialog.FileName
                End If
            End Using
        End Sub

    End Class

End Namespace
