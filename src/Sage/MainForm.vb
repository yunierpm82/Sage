Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports Microsoft.Data.SqlClient

Public Class MainForm
        Inherits Form

        Private txtServer As TextBox
        Private txtDatabase As TextBox
        Private txtUser As TextBox
        Private txtPassword As TextBox
        Private chkIntegratedSecurity As CheckBox
        Private btnConnect As Button
        Private lstTables As ListView
        Private lblStatus As Label

        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub InitializeComponent()
            Me.Text = "Sage - Visor de Tablas SQL"
            Me.ClientSize = New Size(520, 500)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.MinimumSize = New Size(540, 540)

            Dim lblServer As New Label With {.Text = "Servidor:", .Location = New Point(20, 20), .AutoSize = True}
            txtServer = New TextBox With {.Location = New Point(150, 17), .Width = 330, .Text = "localhost"}

            Dim lblDatabase As New Label With {.Text = "Base de datos:", .Location = New Point(20, 55), .AutoSize = True}
            txtDatabase = New TextBox With {.Location = New Point(150, 52), .Width = 330}

            Dim lblUser As New Label With {.Text = "Usuario:", .Location = New Point(20, 90), .AutoSize = True}
            txtUser = New TextBox With {.Location = New Point(150, 87), .Width = 330}

            Dim lblPassword As New Label With {.Text = "Contraseña:", .Location = New Point(20, 125), .AutoSize = True}
            txtPassword = New TextBox With {.Location = New Point(150, 122), .Width = 330, .PasswordChar = "*"c}

            chkIntegratedSecurity = New CheckBox With {
                .Text = "Usar autenticación de Windows (ignora usuario/contraseña)",
                .Location = New Point(150, 155),
                .AutoSize = True
            }
            AddHandler chkIntegratedSecurity.CheckedChanged, AddressOf ChkIntegratedSecurity_CheckedChanged

            btnConnect = New Button With {.Text = "Conectar y listar tablas", .Location = New Point(150, 190), .Width = 220, .Height = 30}
            AddHandler btnConnect.Click, AddressOf BtnConnect_Click

            lblStatus = New Label With {
                .Location = New Point(20, 230),
                .Size = New Size(480, 40),
                .ForeColor = Color.DarkRed,
                .Text = ""
            }

            lstTables = New ListView With {
                .Location = New Point(20, 275),
                .Size = New Size(480, 190),
                .View = View.Details,
                .FullRowSelect = True,
                .GridLines = True,
                .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
            }
            lstTables.Columns.Add("Esquema", 150)
            lstTables.Columns.Add("Tabla", 300)

            Me.Controls.AddRange(New Control() {
                lblServer, txtServer,
                lblDatabase, txtDatabase,
                lblUser, txtUser,
                lblPassword, txtPassword,
                chkIntegratedSecurity,
                btnConnect,
                lblStatus,
                lstTables
            })
        End Sub

        Private Sub ChkIntegratedSecurity_CheckedChanged(sender As Object, e As EventArgs)
            txtUser.Enabled = Not chkIntegratedSecurity.Checked
            txtPassword.Enabled = Not chkIntegratedSecurity.Checked
        End Sub

        Private Async Sub BtnConnect_Click(sender As Object, e As EventArgs)
            lstTables.Items.Clear()
            lblStatus.ForeColor = Color.DarkBlue
            lblStatus.Text = "Conectando..."
            btnConnect.Enabled = False

            Try
                Dim builder As New SqlConnectionStringBuilder With {
                    .DataSource = txtServer.Text.Trim(),
                    .InitialCatalog = txtDatabase.Text.Trim(),
                    .TrustServerCertificate = True,
                    .ConnectTimeout = 10
                }

                If chkIntegratedSecurity.Checked Then
                    builder.IntegratedSecurity = True
                Else
                    builder.UserID = txtUser.Text.Trim()
                    builder.Password = txtPassword.Text
                End If

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

                lblStatus.ForeColor = Color.DarkGreen
                lblStatus.Text = $"Conectado correctamente. {lstTables.Items.Count} tabla(s) encontrada(s)."

            Catch ex As Exception
                lblStatus.ForeColor = Color.DarkRed
                lblStatus.Text = "Error: " & ex.Message
            Finally
                btnConnect.Enabled = True
            End Try
        End Sub

End Class
