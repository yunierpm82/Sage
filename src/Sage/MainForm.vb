Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Data.Sql
Imports Microsoft.Data.SqlClient

Public Class MainForm
    Inherits Form

    Private ReadOnly colorBackground As Color = Color.FromArgb(32, 32, 36)
    Private ReadOnly colorPanel As Color = Color.FromArgb(45, 45, 48)
    Private ReadOnly colorAltRow As Color = Color.FromArgb(40, 40, 44)
    Private ReadOnly colorHeader As Color = Color.FromArgb(55, 55, 60)
    Private ReadOnly colorText As Color = Color.FromArgb(230, 230, 230)
    Private ReadOnly colorAccent As Color = Color.FromArgb(0, 122, 204)
    Private ReadOnly colorButton As Color = Color.FromArgb(60, 60, 65)
    Private ReadOnly colorBorder As Color = Color.FromArgb(70, 70, 75)
    Private ReadOnly colorSuccess As Color = Color.FromArgb(120, 220, 120)
    Private ReadOnly colorError As Color = Color.FromArgb(240, 120, 120)
    Private ReadOnly colorInfo As Color = Color.FromArgb(120, 180, 240)

    ' --- Panel: conexion a base de datos SQL ---
    Private pnlDatabase As Panel
    Private cboServer As ComboBox
    Private cboDatabase As ComboBox
    Private txtUser As TextBox
    Private txtPassword As TextBox
    Private chkIntegratedSecurity As CheckBox
    Private btnDiscoverServers As Button
    Private btnDiscoverDatabases As Button
    Private btnConnect As Button
    Private lstTables As ListView
    Private dgvData As DataGridView
    Private lstRelations As ListView
    Private lblStatus As Label

    Private currentConnectionStringBuilder As SqlConnectionStringBuilder

    ' --- Panel: GL Booking > Batch List (importacion a Sage 300) ---
    Private pnlBatchList As Panel
    Private txtSageCompany As TextBox
    Private txtSageUser As TextBox
    Private txtSagePassword As TextBox
    Private txtExcelPath As TextBox
    Private btnBrowseExcel As Button
    Private cboBatchTemplate As ComboBox
    Private btnImportBatch As Button
    Private lblBatchStatus As Label

    ' --- Panel: GL Booking > Journal Entry (pendiente) ---
    Private pnlJournalEntry As Panel

    ' --- Panel: GL Booking > Relate Columns (mapeo de columnas del Excel) ---
    Private pnlRelateColumns As Panel
    Private cboRelateProfile As ComboBox
    Private btnSaveProfile As Button
    Private btnDeleteProfile As Button
    Private btnChooseSample As Button
    Private lblSampleFile As Label
    Private lstRequiredColumns As ListBox
    Private lstActualColumns As ListBox
    Private pnlLinks As Panel
    Private lblRelateStatus As Label
    Private currentColumnMapping As New Dictionary(Of String, String)
    Private armedRequiredColumn As String = Nothing
    Private ReadOnly NoTemplateOption As String = "(Sin plantilla — nombres exactos)"

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "Sage - Visor de Tablas SQL"
        Me.ClientSize = New Size(900, 650)
        Me.MinimumSize = New Size(760, 500)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = colorBackground
        Me.ForeColor = colorText

        Dim menuStrip = BuildMenuStrip()
        pnlDatabase = BuildDatabasePanel()
        pnlBatchList = BuildBatchListPanel()
        pnlJournalEntry = BuildJournalEntryPanel()
        pnlRelateColumns = BuildRelateColumnsPanel()

        Me.MainMenuStrip = menuStrip
        Me.Controls.AddRange(New Control() {pnlJournalEntry, pnlRelateColumns, pnlBatchList, pnlDatabase, menuStrip})

        ShowPanel(pnlDatabase)
    End Sub

    ' ==================== MENU ====================

    Private Function BuildMenuStrip() As MenuStrip
        Dim menuStrip As New MenuStrip With {
            .Dock = DockStyle.Top,
            .BackColor = colorPanel,
            .ForeColor = colorText,
            .Renderer = New ToolStripProfessionalRenderer(New DarkMenuColorTable(colorPanel, colorAccent))
        }

        Dim mnuDataBase As New ToolStripMenuItem("Data Base") With {.ForeColor = colorText}
        Dim mnuConnect As New ToolStripMenuItem("Connect") With {.ForeColor = colorText}
        Dim mnuDisconnect As New ToolStripMenuItem("Disconnect") With {.ForeColor = colorText}
        AddHandler mnuConnect.Click, Sub() ShowPanel(pnlDatabase)
        AddHandler mnuDisconnect.Click, AddressOf MnuDisconnect_Click
        mnuDataBase.DropDownItems.Add(mnuConnect)
        mnuDataBase.DropDownItems.Add(mnuDisconnect)

        Dim mnuGLBooking As New ToolStripMenuItem("GL Booking") With {.ForeColor = colorText}
        Dim mnuJournalEntry As New ToolStripMenuItem("Journal Entry") With {.ForeColor = colorText}
        Dim mnuBatchList As New ToolStripMenuItem("Batch List") With {.ForeColor = colorText}
        Dim mnuRelateColumns As New ToolStripMenuItem("Relate Columns") With {.ForeColor = colorText}
        AddHandler mnuJournalEntry.Click, Sub() ShowPanel(pnlJournalEntry)
        AddHandler mnuBatchList.Click, AddressOf MnuBatchList_Click
        AddHandler mnuRelateColumns.Click, AddressOf MnuRelateColumns_Click
        mnuGLBooking.DropDownItems.Add(mnuJournalEntry)
        mnuGLBooking.DropDownItems.Add(mnuBatchList)
        mnuGLBooking.DropDownItems.Add(mnuRelateColumns)

        menuStrip.Items.Add(mnuDataBase)
        menuStrip.Items.Add(mnuGLBooking)

        Return menuStrip
    End Function

    Private Sub ShowPanel(panel As Panel)
        pnlDatabase.Visible = (panel Is pnlDatabase)
        pnlBatchList.Visible = (panel Is pnlBatchList)
        pnlJournalEntry.Visible = (panel Is pnlJournalEntry)
        pnlRelateColumns.Visible = (panel Is pnlRelateColumns)
    End Sub

    Private Sub MnuBatchList_Click(sender As Object, e As EventArgs)
        RefreshTemplateList(cboBatchTemplate, includeNoneOption:=True)
        ShowPanel(pnlBatchList)
    End Sub

    Private Sub MnuRelateColumns_Click(sender As Object, e As EventArgs)
        RefreshTemplateList(cboRelateProfile, includeNoneOption:=False)
        ShowPanel(pnlRelateColumns)
    End Sub

    Private Sub RefreshTemplateList(comboBox As ComboBox, includeNoneOption As Boolean)
        Dim currentText = comboBox.Text
        comboBox.Items.Clear()
        If includeNoneOption Then comboBox.Items.Add(NoTemplateOption)
        For Each profileName In Sage300.ColumnMappingStore.ListProfiles()
            comboBox.Items.Add(profileName)
        Next
        If comboBox.Items.Contains(currentText) Then
            comboBox.Text = currentText
        ElseIf includeNoneOption Then
            comboBox.SelectedIndex = 0
        End If
    End Sub

    Private Sub MnuDisconnect_Click(sender As Object, e As EventArgs)
        currentConnectionStringBuilder = Nothing
        lstTables.Items.Clear()
        dgvData.DataSource = Nothing
        lstRelations.Items.Clear()
        lblStatus.ForeColor = colorInfo
        lblStatus.Text = "Desconectado."
        ShowPanel(pnlDatabase)
    End Sub

    ' ==================== PANEL: DATA BASE ====================

    Private Function BuildDatabasePanel() As Panel
        Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = colorBackground}

        Dim lblServer As New Label With {.Text = "Servidor:", .Location = New Point(20, 22), .AutoSize = True, .ForeColor = colorText}
        cboServer = New ComboBox With {
            .Location = New Point(150, 19),
            .Width = 250,
            .DropDownStyle = ComboBoxStyle.DropDown,
            .Text = "localhost",
            .BackColor = colorPanel,
            .ForeColor = colorText
        }
        btnDiscoverServers = New Button With {.Text = "Buscar servidores", .Location = New Point(410, 18), .Width = 160, .Height = 26}
        StyleSecondaryButton(btnDiscoverServers)
        AddHandler btnDiscoverServers.Click, AddressOf BtnDiscoverServers_Click

        Dim lblDatabase As New Label With {.Text = "Base de datos:", .Location = New Point(20, 57), .AutoSize = True, .ForeColor = colorText}
        cboDatabase = New ComboBox With {
            .Location = New Point(150, 54),
            .Width = 250,
            .DropDownStyle = ComboBoxStyle.DropDown,
            .BackColor = colorPanel,
            .ForeColor = colorText
        }
        btnDiscoverDatabases = New Button With {.Text = "Buscar bases de datos", .Location = New Point(410, 53), .Width = 160, .Height = 26}
        StyleSecondaryButton(btnDiscoverDatabases)
        AddHandler btnDiscoverDatabases.Click, AddressOf BtnDiscoverDatabases_Click

        Dim lblUser As New Label With {.Text = "Usuario:", .Location = New Point(20, 92), .AutoSize = True, .ForeColor = colorText}
        txtUser = New TextBox With {.Location = New Point(150, 89), .Width = 250, .BackColor = colorPanel, .ForeColor = colorText, .BorderStyle = BorderStyle.FixedSingle}

        Dim lblPassword As New Label With {.Text = "Contraseña:", .Location = New Point(20, 127), .AutoSize = True, .ForeColor = colorText}
        txtPassword = New TextBox With {.Location = New Point(150, 124), .Width = 250, .PasswordChar = "*"c, .BackColor = colorPanel, .ForeColor = colorText, .BorderStyle = BorderStyle.FixedSingle}

        chkIntegratedSecurity = New CheckBox With {
            .Text = "Usar autenticación de Windows (ignora usuario/contraseña)",
            .Location = New Point(150, 157),
            .AutoSize = True,
            .ForeColor = colorText
        }
        AddHandler chkIntegratedSecurity.CheckedChanged, AddressOf ChkIntegratedSecurity_CheckedChanged

        btnConnect = New Button With {.Text = "Conectar y listar tablas", .Location = New Point(150, 192), .Width = 220, .Height = 32}
        StylePrimaryButton(btnConnect)
        AddHandler btnConnect.Click, AddressOf BtnConnect_Click

        Dim lblRelationsTitle As New Label With {
            .Text = "Relaciones (claves foráneas):",
            .Location = New Point(600, 15),
            .Size = New Size(280, 20),
            .ForeColor = colorText
        }

        lstRelations = New ListView With {
            .Location = New Point(600, 38),
            .Size = New Size(280, 187),
            .View = View.Details,
            .FullRowSelect = True,
            .GridLines = True,
            .MultiSelect = False,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .BackColor = colorPanel,
            .ForeColor = colorText,
            .BorderStyle = BorderStyle.FixedSingle,
            .OwnerDraw = True
        }
        lstRelations.Columns.Add("Tipo", 100)
        lstRelations.Columns.Add("Tabla relacionada", 172)
        AddHandler lstRelations.DrawColumnHeader, AddressOf ListView_DrawColumnHeader
        AddHandler lstRelations.DrawItem, AddressOf ListView_DrawItem
        AddHandler lstRelations.DrawSubItem, AddressOf ListView_DrawSubItem

        lblStatus = New Label With {
            .Location = New Point(20, 234),
            .Size = New Size(860, 40),
            .ForeColor = colorInfo,
            .Text = ""
        }

        lstTables = New ListView With {
            .Location = New Point(20, 280),
            .Size = New Size(260, 350),
            .View = View.Details,
            .FullRowSelect = True,
            .GridLines = True,
            .MultiSelect = False,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left,
            .BackColor = colorPanel,
            .ForeColor = colorText,
            .BorderStyle = BorderStyle.FixedSingle,
            .OwnerDraw = True
        }
        lstTables.Columns.Add("Tabla", 256)
        AddHandler lstTables.ItemSelectionChanged, AddressOf LstTables_ItemSelectionChanged
        AddHandler lstTables.DrawColumnHeader, AddressOf ListView_DrawColumnHeader
        AddHandler lstTables.DrawItem, AddressOf ListView_DrawItem
        AddHandler lstTables.DrawSubItem, AddressOf ListView_DrawSubItem

        dgvData = New DataGridView With {
            .Location = New Point(300, 280),
            .Size = New Size(580, 350),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right,
            .BackgroundColor = colorPanel,
            .ForeColor = colorText,
            .GridColor = colorBorder,
            .BorderStyle = BorderStyle.FixedSingle,
            .EnableHeadersVisualStyles = False,
            .ReadOnly = True,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .AllowUserToResizeRows = False,
            .RowHeadersVisible = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
        }
        dgvData.DefaultCellStyle.BackColor = colorPanel
        dgvData.DefaultCellStyle.ForeColor = colorText
        dgvData.DefaultCellStyle.SelectionBackColor = colorAccent
        dgvData.DefaultCellStyle.SelectionForeColor = Color.White
        dgvData.ColumnHeadersDefaultCellStyle.BackColor = colorHeader
        dgvData.ColumnHeadersDefaultCellStyle.ForeColor = colorText
        dgvData.AlternatingRowsDefaultCellStyle.BackColor = colorAltRow
        dgvData.AlternatingRowsDefaultCellStyle.ForeColor = colorText

        panel.Controls.AddRange(New Control() {
            lblServer, cboServer, btnDiscoverServers,
            lblDatabase, cboDatabase, btnDiscoverDatabases,
            lblUser, txtUser,
            lblPassword, txtPassword,
            chkIntegratedSecurity,
            btnConnect,
            lblRelationsTitle, lstRelations,
            lblStatus,
            lstTables,
            dgvData
        })

        Return panel
    End Function

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

    Private Sub ListView_DrawColumnHeader(sender As Object, e As DrawListViewColumnHeaderEventArgs)
        Using bgBrush As New SolidBrush(colorHeader)
            e.Graphics.FillRectangle(bgBrush, e.Bounds)
        End Using
        Dim sf As New StringFormat With {.LineAlignment = StringAlignment.Center}
        Using textBrush As New SolidBrush(colorText)
            e.Graphics.DrawString(e.Header.Text, e.Font, textBrush, e.Bounds, sf)
        End Using
    End Sub

    Private Sub ListView_DrawItem(sender As Object, e As DrawListViewItemEventArgs)
        ' El dibujo real ocurre en DrawSubItem (vista Details); no se necesita nada aquí.
    End Sub

    Private Sub ListView_DrawSubItem(sender As Object, e As DrawListViewSubItemEventArgs)
        Dim isSelected = e.Item.Selected
        Dim backColor = If(isSelected, colorAccent, colorPanel)
        Dim foreColor = If(isSelected, Color.White, colorText)

        Using bgBrush As New SolidBrush(backColor)
            e.Graphics.FillRectangle(bgBrush, e.Bounds)
        End Using
        Dim sf As New StringFormat With {.LineAlignment = StringAlignment.Center}
        Using textBrush As New SolidBrush(foreColor)
            e.Graphics.DrawString(e.SubItem.Text, e.Item.ListView.Font, textBrush, e.Bounds, sf)
        End Using
    End Sub

    Private Sub ChkIntegratedSecurity_CheckedChanged(sender As Object, e As EventArgs)
        txtUser.Enabled = Not chkIntegratedSecurity.Checked
        txtPassword.Enabled = Not chkIntegratedSecurity.Checked
    End Sub

    Private Function EscapeIdentifier(name As String) As String
        Return "[" & name.Replace("]", "]]") & "]"
    End Function

    Private Function BuildConnectionStringBuilder(targetDatabase As String) As SqlConnectionStringBuilder
        Dim builder As New SqlConnectionStringBuilder With {
            .DataSource = cboServer.Text.Trim(),
            .InitialCatalog = targetDatabase,
            .TrustServerCertificate = True,
            .ConnectTimeout = 10
        }

        If chkIntegratedSecurity.Checked Then
            builder.IntegratedSecurity = True
        Else
            builder.UserID = txtUser.Text.Trim()
            builder.Password = txtPassword.Text
        End If

        Return builder
    End Function

    Private Async Sub BtnConnect_Click(sender As Object, e As EventArgs)
        lstTables.Items.Clear()
        dgvData.DataSource = Nothing
        lstRelations.Items.Clear()
        lblStatus.ForeColor = colorInfo
        lblStatus.Text = "Conectando..."
        btnConnect.Enabled = False

        Try
            Dim builder = BuildConnectionStringBuilder(cboDatabase.Text.Trim())

            Using connection As New SqlConnection(builder.ConnectionString)
                Await connection.OpenAsync()

                Const query As String =
                    "SELECT TABLE_SCHEMA, TABLE_NAME FROM INFORMATION_SCHEMA.TABLES " &
                    "WHERE TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_SCHEMA, TABLE_NAME"

                Using command As New SqlCommand(query, connection)
                    Using reader = Await command.ExecuteReaderAsync()
                        While Await reader.ReadAsync()
                            Dim schemaValue = reader.GetString(0)
                            Dim tableValue = reader.GetString(1)
                            Dim item As New ListViewItem(tableValue) With {.Tag = schemaValue}
                            lstTables.Items.Add(item)
                        End While
                    End Using
                End Using
            End Using

            currentConnectionStringBuilder = builder

            lblStatus.ForeColor = colorSuccess
            lblStatus.Text = $"Conectado correctamente. {lstTables.Items.Count} tabla(s) encontrada(s). Selecciona una tabla para ver sus datos."

        Catch ex As Exception
            currentConnectionStringBuilder = Nothing
            lblStatus.ForeColor = colorError
            lblStatus.Text = "Error: " & ex.Message
        Finally
            btnConnect.Enabled = True
        End Try
    End Sub

    Private Async Sub LstTables_ItemSelectionChanged(sender As Object, e As ListViewItemSelectionChangedEventArgs)
        If Not e.IsSelected Then Return

        Dim schemaName = CStr(e.Item.Tag)
        Dim tableName = e.Item.Text
        Await LoadTableDataAsync(schemaName, tableName)
        Await LoadRelationshipsAsync(schemaName, tableName)
    End Sub

    Private Async Function LoadTableDataAsync(schemaName As String, tableName As String) As Task
        If currentConnectionStringBuilder Is Nothing Then Return

        dgvData.DataSource = Nothing
        lblStatus.ForeColor = colorInfo
        lblStatus.Text = $"Cargando datos de {schemaName}.{tableName}..."

        Try
            Using connection As New SqlConnection(currentConnectionStringBuilder.ConnectionString)
                Await connection.OpenAsync()

                Dim query = $"SELECT TOP 200 * FROM {EscapeIdentifier(schemaName)}.{EscapeIdentifier(tableName)}"

                Using command As New SqlCommand(query, connection)
                    Using reader = Await command.ExecuteReaderAsync()
                        Dim table As New DataTable()
                        table.Load(reader)
                        dgvData.DataSource = table
                    End Using
                End Using
            End Using

            lblStatus.ForeColor = colorSuccess
            lblStatus.Text = $"{dgvData.Rows.Count} fila(s) cargada(s) de {schemaName}.{tableName} (máximo 200)."

        Catch ex As Exception
            lblStatus.ForeColor = colorError
            lblStatus.Text = $"Error cargando datos de {schemaName}.{tableName}: " & ex.Message
        End Try
    End Function

    Private Async Function LoadRelationshipsAsync(schemaName As String, tableName As String) As Task
        lstRelations.Items.Clear()
        If currentConnectionStringBuilder Is Nothing Then Return

        Try
            Using connection As New SqlConnection(currentConnectionStringBuilder.ConnectionString)
                Await connection.OpenAsync()

                Const query As String =
                    "SELECT " &
                    "  CASE WHEN fk.parent_object_id = OBJECT_ID(@t) THEN 'Referencia a' ELSE 'Referenciada por' END AS Tipo, " &
                    "  CASE WHEN fk.parent_object_id = OBJECT_ID(@t) " &
                    "       THEN OBJECT_SCHEMA_NAME(fk.referenced_object_id) + '.' + OBJECT_NAME(fk.referenced_object_id) " &
                    "       ELSE OBJECT_SCHEMA_NAME(fk.parent_object_id) + '.' + OBJECT_NAME(fk.parent_object_id) " &
                    "  END AS TablaRelacionada " &
                    "FROM sys.foreign_keys fk " &
                    "WHERE fk.parent_object_id = OBJECT_ID(@t) OR fk.referenced_object_id = OBJECT_ID(@t) " &
                    "ORDER BY Tipo, TablaRelacionada"

                Using command As New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@t", $"{schemaName}.{tableName}")
                    Using reader = Await command.ExecuteReaderAsync()
                        While Await reader.ReadAsync()
                            Dim item As New ListViewItem(reader.GetString(0))
                            item.SubItems.Add(reader.GetString(1))
                            lstRelations.Items.Add(item)
                        End While
                    End Using
                End Using
            End Using

            If lstRelations.Items.Count = 0 Then
                Dim noneItem As New ListViewItem("—")
                noneItem.SubItems.Add("Sin relaciones (claves foráneas).")
                lstRelations.Items.Add(noneItem)
            End If

        Catch ex As Exception
            lstRelations.Items.Clear()
            Dim errorItem As New ListViewItem("Error")
            errorItem.SubItems.Add(ex.Message)
            lstRelations.Items.Add(errorItem)
        End Try
    End Function

    Private Async Sub BtnDiscoverServers_Click(sender As Object, e As EventArgs)
        btnDiscoverServers.Enabled = False
        lblStatus.ForeColor = colorInfo
        lblStatus.Text = "Buscando servidores SQL en la red... (puede tardar unos segundos)"

        Try
            Dim dataSources = Await Task.Run(Function() SqlDataSourceEnumerator.Instance.GetDataSources())

            cboServer.Items.Clear()
            For Each row As DataRow In dataSources.Rows
                Dim serverName = row("ServerName").ToString()
                Dim instanceName = row("InstanceName").ToString()
                Dim fullName = If(String.IsNullOrEmpty(instanceName), serverName, $"{serverName}\{instanceName}")
                If Not cboServer.Items.Contains(fullName) Then
                    cboServer.Items.Add(fullName)
                End If
            Next

            If cboServer.Items.Count > 0 Then
                cboServer.DroppedDown = True
                lblStatus.ForeColor = colorSuccess
                lblStatus.Text = $"Se encontraron {cboServer.Items.Count} servidor(es) en la red local."
            Else
                lblStatus.ForeColor = colorError
                lblStatus.Text = "No se encontraron servidores SQL en la red local."
            End If

        Catch ex As Exception
            lblStatus.ForeColor = colorError
            lblStatus.Text = "Error buscando servidores: " & ex.Message
        Finally
            btnDiscoverServers.Enabled = True
        End Try
    End Sub

    Private Async Sub BtnDiscoverDatabases_Click(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(cboServer.Text) Then
            lblStatus.ForeColor = colorError
            lblStatus.Text = "Especifica primero un servidor."
            Return
        End If

        btnDiscoverDatabases.Enabled = False
        lblStatus.ForeColor = colorInfo
        lblStatus.Text = "Buscando bases de datos..."

        Try
            Dim builder = BuildConnectionStringBuilder("master")

            Using connection As New SqlConnection(builder.ConnectionString)
                Await connection.OpenAsync()

                Using command As New SqlCommand("SELECT name FROM sys.databases ORDER BY name", connection)
                    Using reader = Await command.ExecuteReaderAsync()
                        cboDatabase.Items.Clear()
                        While Await reader.ReadAsync()
                            cboDatabase.Items.Add(reader.GetString(0))
                        End While
                    End Using
                End Using
            End Using

            If cboDatabase.Items.Count > 0 Then
                cboDatabase.DroppedDown = True
                lblStatus.ForeColor = colorSuccess
                lblStatus.Text = $"Se encontraron {cboDatabase.Items.Count} base(s) de datos."
            Else
                lblStatus.ForeColor = colorError
                lblStatus.Text = "No se encontraron bases de datos."
            End If

        Catch ex As Exception
            lblStatus.ForeColor = colorError
            lblStatus.Text = "Error buscando bases de datos: " & ex.Message
        Finally
            btnDiscoverDatabases.Enabled = True
        End Try
    End Sub

    ' ==================== PANEL: GL BOOKING > BATCH LIST ====================

    Private Function BuildBatchListPanel() As Panel
        Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = colorBackground, .Visible = False}

        Dim lblCompany As New Label With {.Text = "Compañía:", .Location = New Point(20, 22), .AutoSize = True, .ForeColor = colorText}
        txtSageCompany = New TextBox With {.Location = New Point(160, 19), .Width = 300, .BackColor = colorPanel, .ForeColor = colorText, .BorderStyle = BorderStyle.FixedSingle}

        Dim lblSageUser As New Label With {.Text = "Usuario Sage:", .Location = New Point(20, 57), .AutoSize = True, .ForeColor = colorText}
        txtSageUser = New TextBox With {.Location = New Point(160, 54), .Width = 300, .BackColor = colorPanel, .ForeColor = colorText, .BorderStyle = BorderStyle.FixedSingle}

        Dim lblSagePassword As New Label With {.Text = "Contraseña:", .Location = New Point(20, 92), .AutoSize = True, .ForeColor = colorText}
        txtSagePassword = New TextBox With {.Location = New Point(160, 89), .Width = 300, .PasswordChar = "*"c, .BackColor = colorPanel, .ForeColor = colorText, .BorderStyle = BorderStyle.FixedSingle}

        Dim lblExcelPath As New Label With {.Text = "Archivo Excel (.xlsx):", .Location = New Point(20, 127), .AutoSize = True, .ForeColor = colorText}
        txtExcelPath = New TextBox With {
            .Location = New Point(160, 124),
            .Width = 380,
            .ReadOnly = True,
            .BackColor = colorPanel,
            .ForeColor = colorText,
            .BorderStyle = BorderStyle.FixedSingle
        }
        btnBrowseExcel = New Button With {.Text = "Elegir...", .Location = New Point(548, 123), .Width = 90, .Height = 26}
        StyleSecondaryButton(btnBrowseExcel)
        AddHandler btnBrowseExcel.Click, AddressOf BtnBrowseExcel_Click

        Dim lblTemplate As New Label With {.Text = "Plantilla de columnas:", .Location = New Point(20, 162), .AutoSize = True, .ForeColor = colorText}
        cboBatchTemplate = New ComboBox With {
            .Location = New Point(160, 159),
            .Width = 300,
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .BackColor = colorPanel,
            .ForeColor = colorText
        }

        btnImportBatch = New Button With {.Text = "Importar", .Location = New Point(160, 200), .Width = 220, .Height = 32}
        StylePrimaryButton(btnImportBatch)
        AddHandler btnImportBatch.Click, AddressOf BtnImportBatch_Click

        lblBatchStatus = New Label With {
            .Location = New Point(20, 248),
            .Size = New Size(820, 120),
            .ForeColor = colorInfo,
            .Text = ""
        }

        panel.Controls.AddRange(New Control() {
            lblCompany, txtSageCompany,
            lblSageUser, txtSageUser,
            lblSagePassword, txtSagePassword,
            lblExcelPath, txtExcelPath, btnBrowseExcel,
            lblTemplate, cboBatchTemplate,
            btnImportBatch,
            lblBatchStatus
        })

        Return panel
    End Function

    Private Sub BtnBrowseExcel_Click(sender As Object, e As EventArgs)
        Using dialog As New OpenFileDialog With {
            .Filter = "Archivos Excel (*.xlsx)|*.xlsx",
            .Title = "Selecciona el archivo de transacciones"
        }
            If dialog.ShowDialog() = DialogResult.OK Then
                txtExcelPath.Text = dialog.FileName
            End If
        End Using
    End Sub

    Private Async Sub BtnImportBatch_Click(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(txtSageCompany.Text) OrElse String.IsNullOrWhiteSpace(txtSageUser.Text) Then
            lblBatchStatus.ForeColor = colorError
            lblBatchStatus.Text = "Especifica la compañía y el usuario de Sage."
            Return
        End If

        If String.IsNullOrWhiteSpace(txtExcelPath.Text) OrElse Not File.Exists(txtExcelPath.Text) Then
            lblBatchStatus.ForeColor = colorError
            lblBatchStatus.Text = "Selecciona un archivo Excel válido."
            Return
        End If

        btnImportBatch.Enabled = False
        lblBatchStatus.ForeColor = colorInfo
        lblBatchStatus.Text = "Importando..."

        Dim company = txtSageCompany.Text.Trim()
        Dim user = txtSageUser.Text.Trim()
        Dim password = txtSagePassword.Text
        Dim path = txtExcelPath.Text

        Dim mapping As Dictionary(Of String, String) = Nothing
        If cboBatchTemplate.Text <> "" AndAlso cboBatchTemplate.Text <> NoTemplateOption Then
            mapping = Sage300.ColumnMappingStore.Load(cboBatchTemplate.Text)
        End If

        Try
            Dim result = Await Task.Run(Function() New Sage300.Sage300BatchImporter().ImportBatch(company, user, password, path, mapping))
            lblBatchStatus.ForeColor = If(result.Success, colorSuccess, colorError)
            lblBatchStatus.Text = result.Message
        Catch ex As Exception
            lblBatchStatus.ForeColor = colorError
            lblBatchStatus.Text = "Error inesperado: " & ex.Message
        Finally
            btnImportBatch.Enabled = True
        End Try
    End Sub

    ' ==================== PANEL: GL BOOKING > JOURNAL ENTRY (pendiente) ====================

    Private Function BuildJournalEntryPanel() As Panel
        Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = colorBackground, .Visible = False}

        Dim lblComingSoon As New Label With {
            .Text = "Journal Entry — próximamente.",
            .Location = New Point(20, 20),
            .AutoSize = True,
            .ForeColor = colorText
        }
        panel.Controls.Add(lblComingSoon)

        Return panel
    End Function

    ' ==================== PANEL: GL BOOKING > RELATE COLUMNS ====================

    Private Function BuildRelateColumnsPanel() As Panel
        Dim panel As New Panel With {.Dock = DockStyle.Fill, .BackColor = colorBackground, .Visible = False}

        Dim lblProfile As New Label With {.Text = "Plantilla:", .Location = New Point(20, 18), .AutoSize = True, .ForeColor = colorText}
        cboRelateProfile = New ComboBox With {
            .Location = New Point(100, 15),
            .Width = 220,
            .DropDownStyle = ComboBoxStyle.DropDown,
            .BackColor = colorPanel,
            .ForeColor = colorText
        }
        AddHandler cboRelateProfile.SelectedIndexChanged, AddressOf CboRelateProfile_SelectedIndexChanged

        btnSaveProfile = New Button With {.Text = "Guardar", .Location = New Point(330, 14), .Width = 100, .Height = 26}
        StylePrimaryButton(btnSaveProfile)
        AddHandler btnSaveProfile.Click, AddressOf BtnSaveProfile_Click

        btnDeleteProfile = New Button With {.Text = "Eliminar", .Location = New Point(440, 14), .Width = 100, .Height = 26}
        StyleSecondaryButton(btnDeleteProfile)
        AddHandler btnDeleteProfile.Click, AddressOf BtnDeleteProfile_Click

        btnChooseSample = New Button With {.Text = "Elegir Excel de muestra...", .Location = New Point(20, 55), .Width = 220, .Height = 26}
        StyleSecondaryButton(btnChooseSample)
        AddHandler btnChooseSample.Click, AddressOf BtnChooseSample_Click

        lblSampleFile = New Label With {
            .Text = "(ningún archivo elegido)",
            .Location = New Point(250, 59),
            .Size = New Size(500, 20),
            .ForeColor = colorText,
            .AutoEllipsis = True
        }

        Dim lblRequiredHeader As New Label With {.Text = "Columnas necesarias", .Location = New Point(20, 95), .AutoSize = True, .ForeColor = colorText}
        Dim lblActualHeader As New Label With {.Text = "Columnas del Excel", .Location = New Point(410, 95), .AutoSize = True, .ForeColor = colorText}
        Dim lblHint As New Label With {
            .Text = "Haz clic en una columna de la izquierda y luego en la de la derecha que le corresponde para enlazarlas. Doble clic en la izquierda para quitar un enlace.",
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
        lstRequiredColumns.Items.AddRange(Sage300.ExcelTransactionReader.RequiredColumnKeys)
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

        lblRelateStatus = New Label With {
            .Location = New Point(20, 425),
            .Size = New Size(650, 40),
            .ForeColor = colorInfo,
            .Text = ""
        }

        panel.Controls.AddRange(New Control() {
            lblProfile, cboRelateProfile, btnSaveProfile, btnDeleteProfile,
            btnChooseSample, lblSampleFile,
            lblRequiredHeader, lblActualHeader,
            lstRequiredColumns, pnlLinks, lstActualColumns,
            lblHint,
            lblRelateStatus
        })

        Return panel
    End Function

    Private Sub LstRequiredColumns_SelectedIndexChanged(sender As Object, e As EventArgs)
        If lstRequiredColumns.SelectedItem Is Nothing Then Return
        armedRequiredColumn = lstRequiredColumns.SelectedItem.ToString()
        lblRelateStatus.ForeColor = colorInfo
        lblRelateStatus.Text = $"Selecciona ahora la columna del Excel que corresponde a '{armedRequiredColumn}'."
    End Sub

    Private Sub LstRequiredColumns_DoubleClick(sender As Object, e As EventArgs)
        If lstRequiredColumns.SelectedItem Is Nothing Then Return
        Dim key = lstRequiredColumns.SelectedItem.ToString()
        If currentColumnMapping.ContainsKey(key) Then
            currentColumnMapping.Remove(key)
            lblRelateStatus.ForeColor = colorInfo
            lblRelateStatus.Text = $"Se quitó el enlace de '{key}'."
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

        lblRelateStatus.ForeColor = colorSuccess
        lblRelateStatus.Text = $"'{armedRequiredColumn}' → '{actualColumn}'"

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
            .Filter = "Archivos Excel (*.xlsx)|*.xlsx",
            .Title = "Selecciona un Excel de muestra de la aplicación de terceros"
        }
            If dialog.ShowDialog() = DialogResult.OK Then
                Try
                    Dim headers = Sage300.ExcelTransactionReader.ReadHeaders(dialog.FileName)
                    lstActualColumns.Items.Clear()
                    lstActualColumns.Items.AddRange(headers.ToArray())
                    lblSampleFile.Text = Path.GetFileName(dialog.FileName)
                    lblRelateStatus.ForeColor = colorSuccess
                    lblRelateStatus.Text = $"Se encontraron {headers.Count} columna(s) en el archivo."
                    pnlLinks.Invalidate()
                Catch ex As Exception
                    lblRelateStatus.ForeColor = colorError
                    lblRelateStatus.Text = "Error leyendo el Excel: " & ex.Message
                End Try
            End If
        End Using
    End Sub

    Private Sub CboRelateProfile_SelectedIndexChanged(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(cboRelateProfile.Text) Then Return
        currentColumnMapping = Sage300.ColumnMappingStore.Load(cboRelateProfile.Text)
        lblRelateStatus.ForeColor = colorInfo
        lblRelateStatus.Text = $"Plantilla '{cboRelateProfile.Text}' cargada ({currentColumnMapping.Count} enlace(s))."
        pnlLinks.Invalidate()
    End Sub

    Private Sub BtnSaveProfile_Click(sender As Object, e As EventArgs)
        Dim name = cboRelateProfile.Text.Trim()
        If String.IsNullOrWhiteSpace(name) Then
            lblRelateStatus.ForeColor = colorError
            lblRelateStatus.Text = "Escribe un nombre para la plantilla antes de guardar."
            Return
        End If

        Sage300.ColumnMappingStore.Save(name, currentColumnMapping)
        RefreshTemplateList(cboRelateProfile, includeNoneOption:=False)
        cboRelateProfile.Text = name
        lblRelateStatus.ForeColor = colorSuccess
        lblRelateStatus.Text = $"Plantilla '{name}' guardada con {currentColumnMapping.Count} enlace(s)."
    End Sub

    Private Sub BtnDeleteProfile_Click(sender As Object, e As EventArgs)
        Dim name = cboRelateProfile.Text.Trim()
        If String.IsNullOrWhiteSpace(name) Then Return

        Sage300.ColumnMappingStore.Delete(name)
        RefreshTemplateList(cboRelateProfile, includeNoneOption:=False)
        cboRelateProfile.Text = ""
        currentColumnMapping = New Dictionary(Of String, String)
        pnlLinks.Invalidate()
        lblRelateStatus.ForeColor = colorInfo
        lblRelateStatus.Text = $"Plantilla '{name}' eliminada."
    End Sub

End Class

''' Tabla de colores oscura para el MenuStrip (ToolStripProfessionalRenderer).
Friend Class DarkMenuColorTable
    Inherits ProfessionalColorTable

    Private ReadOnly _background As Color
    Private ReadOnly _highlight As Color

    Public Sub New(background As Color, highlight As Color)
        _background = background
        _highlight = highlight
    End Sub

    Public Overrides ReadOnly Property ToolStripDropDownBackground As Color
        Get
            Return _background
        End Get
    End Property

    Public Overrides ReadOnly Property ImageMarginGradientBegin As Color
        Get
            Return _background
        End Get
    End Property

    Public Overrides ReadOnly Property ImageMarginGradientMiddle As Color
        Get
            Return _background
        End Get
    End Property

    Public Overrides ReadOnly Property ImageMarginGradientEnd As Color
        Get
            Return _background
        End Get
    End Property

    Public Overrides ReadOnly Property MenuItemSelected As Color
        Get
            Return _highlight
        End Get
    End Property

    Public Overrides ReadOnly Property MenuItemSelectedGradientBegin As Color
        Get
            Return _highlight
        End Get
    End Property

    Public Overrides ReadOnly Property MenuItemSelectedGradientEnd As Color
        Get
            Return _highlight
        End Get
    End Property

    Public Overrides ReadOnly Property MenuItemBorder As Color
        Get
            Return _highlight
        End Get
    End Property

    Public Overrides ReadOnly Property MenuStripGradientBegin As Color
        Get
            Return _background
        End Get
    End Property

    Public Overrides ReadOnly Property MenuStripGradientEnd As Color
        Get
            Return _background
        End Get
    End Property

    Public Overrides ReadOnly Property MenuBorder As Color
        Get
            Return _highlight
        End Get
    End Property
End Class
