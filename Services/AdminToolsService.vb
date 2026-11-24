Imports System.IO
Imports Microsoft.Data.Sqlite

''' <summary>Business settings, backups and the shared administrator audit writer.</summary>
Public Module AdminToolsService

    ''' <summary>Writes an administrator audit row inside the caller's transaction.</summary>
    Public Sub WriteAudit(conn As SqliteConnection, tx As SqliteTransaction,
                          eventType As String, details As String)
        Using cmd As SqliteCommand = conn.CreateCommand()
            cmd.Transaction = tx
            cmd.CommandText =
                "INSERT INTO AdministratorHistory (EventType, PerformedById, PerformedByName, " &
                "PerformedByRole, Details, EventAt) " &
                "VALUES (@e, @id, @n, 'Administrator', @d, @t)"
            cmd.Parameters.Add(DatabaseManager.P("@e", eventType))
            cmd.Parameters.Add(DatabaseManager.P("@id", SessionManager.CurrentUserId))
            cmd.Parameters.Add(DatabaseManager.P("@n", SessionManager.CurrentUserName))
            cmd.Parameters.Add(DatabaseManager.P("@d", details))
            cmd.Parameters.Add(DatabaseManager.P("@t", DatabaseManager.NowText()))
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Public Function GetSetting(key As String) As String
        Return Convert.ToString(DatabaseManager.ExecuteScalar(
            "SELECT Value FROM Settings WHERE Key = @k", DatabaseManager.P("@k", key)))
    End Function

    Private Sub SaveSetting(conn As SqliteConnection, tx As SqliteTransaction,
                            key As String, value As String)
        Using cmd As SqliteCommand = conn.CreateCommand()
            cmd.Transaction = tx
            cmd.CommandText = "INSERT OR REPLACE INTO Settings (Key, Value) VALUES (@k, @v)"
            cmd.Parameters.Add(DatabaseManager.P("@k", key))
            cmd.Parameters.Add(DatabaseManager.P("@v", value))
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    ''' <summary>Returns "" on success, otherwise an error message.</summary>
    Public Function SaveBusinessSettings(bizName As String, phone As String,
                                         address As String) As String
        SessionManager.RequireAdmin()
        Dim n As String = If(bizName, "").Trim()
        Dim p As String = If(phone, "").Trim()
        Dim a As String = If(address, "").Trim()
        If n = "" Then Return "The business name is required."
        If n.Length > 60 Then Return "The business name is too long (60 characters maximum)."
        If p.Length > 30 Then Return "The phone number is too long (30 characters maximum)."
        If a.Length > 120 Then Return "The address is too long (120 characters maximum)."

        Try
            Using conn As SqliteConnection = DatabaseManager.GetConnection()
                Using tx As SqliteTransaction = conn.BeginTransaction()
                    Try
                        SaveSetting(conn, tx, "BusinessName", n)
                        SaveSetting(conn, tx, "BusinessPhone", p)
                        SaveSetting(conn, tx, "BusinessAddress", a)
                        WriteAudit(conn, tx, "SettingsChanged",
                                   "Updated the business details (name: '" & n & "').")
                        tx.Commit()
                    Catch
                        tx.Rollback()
                        Throw
                    End Try
                End Using
            End Using
            Return ""
        Catch ex As Exception
            Return "The settings could not be saved. (" & ex.Message & ")"
        End Try
    End Function

    ''' <summary>Consistent copy of the live database using SQLite's backup API.</summary>
    Public Function CreateBackup(destinationPath As String) As String
        SessionManager.RequireAdmin()
        Try
            If String.IsNullOrWhiteSpace(destinationPath) Then Return "Choose where to save the backup."
            Dim full As String = Path.GetFullPath(destinationPath)
            If String.Equals(full, Path.GetFullPath(DatabaseManager.DbPath),
                             StringComparison.OrdinalIgnoreCase) Then
                Return "A backup cannot overwrite the live database."
            End If

            Directory.CreateDirectory(Path.GetDirectoryName(full))
            If File.Exists(full) Then File.Delete(full)

            Using src As SqliteConnection = DatabaseManager.GetConnection()
                Using dest As New SqliteConnection("Data Source=" & full & ";Pooling=False")
                    dest.Open()
                    src.BackupDatabase(dest)
                End Using
            End Using

            DatabaseManager.ExecuteNonQuery(
                "INSERT INTO AdministratorHistory (EventType, PerformedById, PerformedByName, " &
                "PerformedByRole, Details, EventAt) VALUES ('DatabaseBackup', @id, @n, " &
                "'Administrator', @d, @t)",
                DatabaseManager.P("@id", SessionManager.CurrentUserId),
                DatabaseManager.P("@n", SessionManager.CurrentUserName),
                DatabaseManager.P("@d", "Created a database backup: " & full),
                DatabaseManager.P("@t", DatabaseManager.NowText()))
            Return ""
        Catch ex As Exception
            Return "The backup failed: " & ex.Message
        End Try
    End Function

End Module