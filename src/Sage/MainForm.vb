Imports System
Imports System.Data
Imports System.Drawing
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
    Private lblStatus As Label

    Private currentConnectionStringBuilder As SqlConnectionStringBuilder

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
        lstTables.Columns.Add("Esquema", 95)
        lstTables.Columns.Add("Tabla", 161)
        AddHandler lstTables.ItemSelectionChanged, AddressOf LstTables_ItemSelectionChanged
        AddHandler lstTables.DrawColumnHeader, AddressOf LstTables_DrawColumnHeader
        AddHandler lstTables.DrawItem, AddressOf LstTables_DrawItem
        AddHandler lstTables.DrawSubItem, AddressOf LstTables_DrawSubItem

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

        Me.Controls.AddRange(New Control() {
            lblServer, cboServer, btnDiscoverServers,
            lblDatabase, cboDatabase, btnDiscoverDatabases,
            lblUser, txtUser,
            lblPassword, txtPassword,
            chkIntegratedSecurity,
            btnConnect,
            lblStatus,
            lstTables,
            dgvData
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

    Private Sub LstTables_DrawColumnHeader(sender As Object, e As DrawListViewColumnHeaderEventArgs)
        Using bgBrush As New SolidBrush(colorHeader)
            e.Graphics.FillRectangle(bgBrush, e.Bounds)
        End Using
        Dim sf As New StringFormat With {.LineAlignment = StringAlignment.Center}
        Using textBrush As New SolidBrush(colorText)
            e.Graphics.DrawString(e.Header.Text, e.Font, textBrush, e.Bounds, sf)
        End Using
    End Sub

    Private Sub LstTables_DrawItem(sender As Object, e As DrawListViewItemEventArgs)
        ' El dibujo real ocurre en DrawSubItem (vista Details); no se necesita nada aquí.
    End Sub

    Private Sub LstTables_DrawSubItem(sender As Object, e As DrawListViewSubItemEventArgs)
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
                            Dim item As New ListViewItem(reader.GetString(0))
                            item.SubItems.Add(reader.GetString(1))
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

        Dim schemaName = e.Item.Text
        Dim tableName = e.Item.SubItems(1).Text
        Await LoadTableDataAsync(schemaName, tableName)
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

End Class
