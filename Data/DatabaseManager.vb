Imports System.IO
Imports Microsoft.Data.Sqlite

''' <summary>
''' Central place for all database access: file location,
''' schema creation and reusable query helpers.
''' </summary>
Public Module DatabaseManager

    Private ReadOnly DbFolder As String =
        Path.Combine(Environment.GetFolderPath(
            Environment.SpecialFolder.ApplicationData), "PhoneRepairSystem")

    Public ReadOnly DbPath As String = Path.Combine(DbFolder, "repair.db")

    Private ReadOnly ConnectionString As String =
        "Data Source=" & DbPath

    ' ---------- Connection ----------

    ''' <summary>Opens a connection with foreign keys enforced.</summary>
    Public Function GetConnection() As SqliteConnection
        Dim conn As New SqliteConnection(ConnectionString)
        conn.Open()
        Using cmd As SqliteCommand = conn.CreateCommand()
            cmd.CommandText = "PRAGMA foreign_keys = ON;"
            cmd.ExecuteNonQuery()
        End Using
        Return conn
    End Function

    ' ---------- Helpers ----------

    ''' <summary>Adds parameters to a command one by one.</summary>
    Private Sub AddParams(cmd As SqliteCommand, params() As SqliteParameter)
        If params Is Nothing Then Return
        For Each prm As SqliteParameter In params
            cmd.Parameters.Add(prm)
        Next
    End Sub

    ''' <summary>Creates a SQL parameter. Nothing becomes NULL.</summary>
    Public Function P(name As String, value As Object) As SqliteParameter
        If value Is Nothing Then value = DBNull.Value
        Return New SqliteParameter(name, value)
    End Function

    ''' <summary>Current local time as stored text.</summary>
    Public Function NowText() As String
        Return DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
    End Function

    Public Function TodayText() As String
        Return DateTime.Now.ToString("yyyy-MM-dd")
    End Function

    ''' <summary>Runs INSERT/UPDATE/DELETE. Returns rows affected.</summary>
    Public Function ExecuteNonQuery(sql As String,
                                    ParamArray params() As SqliteParameter) As Integer
        Using conn As SqliteConnection = GetConnection()
            Using cmd As SqliteCommand = conn.CreateCommand()
                cmd.CommandText = sql
                AddParams(cmd, params)
                Return cmd.ExecuteNonQuery()
            End Using
        End Using
    End Function

    ''' <summary>Runs a query and returns the first value.</summary>
    Public Function ExecuteScalar(sql As String,
                                  ParamArray params() As SqliteParameter) As Object
        Using conn As SqliteConnection = GetConnection()
            Using cmd As SqliteCommand = conn.CreateCommand()
                cmd.CommandText = sql
                AddParams(cmd, params)
                Return cmd.ExecuteScalar()
            End Using
        End Using
    End Function

    ''' <summary>Runs a SELECT and returns the rows as a DataTable.</summary>
    Public Function GetTable(sql As String,
                             ParamArray params() As SqliteParameter) As DataTable
        Dim dt As New DataTable()
        Using conn As SqliteConnection = GetConnection()
            Using cmd As SqliteCommand = conn.CreateCommand()
                cmd.CommandText = sql
                AddParams(cmd, params)
                Using reader As SqliteDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return dt
    End Function

    ' ---------- Setup ----------

    ''' <summary>Call once at startup. Creates folder, file and tables.</summary>
    Public Sub Initialize()
        Directory.CreateDirectory(DbFolder)

        Using conn As SqliteConnection = GetConnection()
            Using tx As SqliteTransaction = conn.BeginTransaction()
                Using cmd As SqliteCommand = conn.CreateCommand()
                    cmd.Transaction = tx
                    cmd.CommandText = SchemaSql
                    cmd.ExecuteNonQuery()
                End Using
                tx.Commit()
            End Using
        End Using
    End Sub

    Private ReadOnly SchemaSql As String = "
CREATE TABLE IF NOT EXISTS Administrators (
    AdminId       INTEGER PRIMARY KEY AUTOINCREMENT,
    Username      TEXT NOT NULL UNIQUE,
    PasswordHash  TEXT NOT NULL,
    PasswordSalt  TEXT NOT NULL,
    IsActive      INTEGER NOT NULL DEFAULT 1,
    CreatedAt     TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS Employees (
    EmployeeId      INTEGER PRIMARY KEY AUTOINCREMENT,
    FullName        TEXT NOT NULL,
    AccessCodeHash  TEXT NOT NULL,
    AccessCodeSalt  TEXT NOT NULL,
    RegisteredAt    TEXT NOT NULL,
    IsActive        INTEGER NOT NULL DEFAULT 1
);

CREATE TABLE IF NOT EXISTS EmployeeSessions (
    SessionId     INTEGER PRIMARY KEY AUTOINCREMENT,
    EmployeeId    INTEGER NOT NULL,
    LoginAt       TEXT NOT NULL,
    LogoutAt      TEXT,
    LogoutReason  TEXT,
    FOREIGN KEY (EmployeeId) REFERENCES Employees(EmployeeId)
);

CREATE TABLE IF NOT EXISTS Repairs (
    RepairId               INTEGER PRIMARY KEY AUTOINCREMENT,
    RepairCode             TEXT NOT NULL UNIQUE,
    CustomerName           TEXT NOT NULL,
    CustomerPhone          TEXT NOT NULL,
    Brand                  TEXT NOT NULL,
    Model                  TEXT NOT NULL,
    ProblemDescription     TEXT NOT NULL,
    EstimatedCost          REAL NOT NULL CHECK (EstimatedCost >= 0),
    Status                 TEXT NOT NULL DEFAULT 'Received'
        CHECK (Status IN ('Received','In Progress','Completed','Collected','Unrepairable')),
    DateReceived           TEXT NOT NULL,
    ExpectedCollectionDate TEXT NOT NULL,
    CollectedAt            TEXT,
    EmployeeId             INTEGER NOT NULL,
    CreatedAt              TEXT NOT NULL,
    UpdatedAt              TEXT,
    FOREIGN KEY (EmployeeId) REFERENCES Employees(EmployeeId)
);

CREATE TABLE IF NOT EXISTS Payments (
    PaymentId      INTEGER PRIMARY KEY AUTOINCREMENT,
    RepairCode     TEXT NOT NULL,
    Amount         REAL NOT NULL CHECK (Amount > 0),
    PaymentMethod  TEXT NOT NULL DEFAULT 'Cash',
    PaidAt         TEXT NOT NULL,
    EmployeeId     INTEGER NOT NULL,
    Note           TEXT,
    FOREIGN KEY (EmployeeId) REFERENCES Employees(EmployeeId)
);

CREATE TABLE IF NOT EXISTS Receipts (
    ReceiptId      INTEGER PRIMARY KEY AUTOINCREMENT,
    ReceiptNumber  TEXT NOT NULL UNIQUE,
    PaymentId      INTEGER NOT NULL UNIQUE,
    IssuedAt       TEXT NOT NULL,
    PrintCount     INTEGER NOT NULL DEFAULT 0,
    LastPrintedAt  TEXT,
    FOREIGN KEY (PaymentId) REFERENCES Payments(PaymentId)
);

CREATE TABLE IF NOT EXISTS ActivityLogs (
    ActivityId          INTEGER PRIMARY KEY AUTOINCREMENT,
    EmployeeId          INTEGER NOT NULL,
    EmployeeName        TEXT NOT NULL,
    ActionType          TEXT NOT NULL,
    RepairCode          TEXT,
    Description         TEXT NOT NULL,
    ActivityDate        TEXT NOT NULL,
    ActivityTime        TEXT NOT NULL,
    Timestamp           TEXT NOT NULL,
    WasSuccessful       INTEGER NOT NULL DEFAULT 1,
    HiddenFromEmployee  INTEGER NOT NULL DEFAULT 0,
    FOREIGN KEY (EmployeeId) REFERENCES Employees(EmployeeId)
);

CREATE TABLE IF NOT EXISTS DeletedRepairs (
    DeletedId               INTEGER PRIMARY KEY AUTOINCREMENT,
    RepairCode              TEXT NOT NULL,
    CustomerName            TEXT NOT NULL,
    CustomerPhone           TEXT NOT NULL,
    Brand                   TEXT NOT NULL,
    Model                   TEXT NOT NULL,
    ProblemDescription      TEXT NOT NULL,
    EstimatedCost           REAL NOT NULL,
    Status                  TEXT NOT NULL,
    DateReceived            TEXT NOT NULL,
    ExpectedCollectionDate  TEXT NOT NULL,
    TotalPaidAtDeletion     REAL NOT NULL DEFAULT 0,
    OriginalEmployeeId      INTEGER NOT NULL,
    DeletedByEmployeeId     INTEGER NOT NULL,
    DeletedByName           TEXT NOT NULL,
    DeletedAt               TEXT NOT NULL,
    DeletionReason          TEXT,
    FOREIGN KEY (DeletedByEmployeeId) REFERENCES Employees(EmployeeId)
);

CREATE TABLE IF NOT EXISTS AdministratorHistory (
    HistoryId        INTEGER PRIMARY KEY AUTOINCREMENT,
    EventType        TEXT NOT NULL,
    PerformedById    INTEGER,
    PerformedByName  TEXT NOT NULL,
    PerformedByRole  TEXT NOT NULL,
    RepairCode       TEXT,
    OriginalData     TEXT,
    Details          TEXT NOT NULL,
    EventAt          TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS DailyClosings (
    ClosingId           INTEGER PRIMARY KEY AUTOINCREMENT,
    BusinessDate        TEXT NOT NULL UNIQUE,
    TotalRevenue        REAL NOT NULL,
    PaymentCount        INTEGER NOT NULL,
    RepairsRegistered   INTEGER NOT NULL,
    RepairsCompleted    INTEGER NOT NULL,
    OutstandingBalance  REAL NOT NULL,
    ClosedByName        TEXT NOT NULL,
    ClosedByRole        TEXT NOT NULL,
    ClosedById          INTEGER,
    ClosedAt            TEXT NOT NULL,
    ReportPath          TEXT
);

CREATE TABLE IF NOT EXISTS EmployeeDailyRevenue (
    Id            INTEGER PRIMARY KEY AUTOINCREMENT,
    ClosingId     INTEGER NOT NULL,
    EmployeeId    INTEGER NOT NULL,
    EmployeeName  TEXT NOT NULL,
    Revenue       REAL NOT NULL,
    PaymentCount  INTEGER NOT NULL,
    UNIQUE (ClosingId, EmployeeId),
    FOREIGN KEY (ClosingId) REFERENCES DailyClosings(ClosingId),
    FOREIGN KEY (EmployeeId) REFERENCES Employees(EmployeeId)
);
CREATE TABLE IF NOT EXISTS Refunds (
    RefundId      INTEGER PRIMARY KEY AUTOINCREMENT,
    RepairCode    TEXT NOT NULL,
    Amount        REAL NOT NULL CHECK (Amount > 0),
    RefundMethod  TEXT NOT NULL DEFAULT 'Cash',
    RefundedAt    TEXT NOT NULL,
    EmployeeId    INTEGER NOT NULL,
    EmployeeName  TEXT NOT NULL,
    Reason        TEXT,
    FOREIGN KEY (EmployeeId) REFERENCES Employees(EmployeeId)
);
CREATE INDEX IF NOT EXISTS idx_refunds_date ON Refunds(RefundedAt);
CREATE TABLE IF NOT EXISTS Counters (
    Name       TEXT PRIMARY KEY,
    LastValue  INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS Settings (
    Key    TEXT PRIMARY KEY,
    Value  TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_repairs_customer  ON Repairs(CustomerName);
CREATE INDEX IF NOT EXISTS idx_repairs_phone     ON Repairs(CustomerPhone);
CREATE INDEX IF NOT EXISTS idx_repairs_status    ON Repairs(Status);
CREATE INDEX IF NOT EXISTS idx_repairs_due       ON Repairs(ExpectedCollectionDate);
CREATE INDEX IF NOT EXISTS idx_payments_repair   ON Payments(RepairCode);
CREATE INDEX IF NOT EXISTS idx_payments_paidat   ON Payments(PaidAt);
CREATE INDEX IF NOT EXISTS idx_payments_employee ON Payments(EmployeeId);
CREATE INDEX IF NOT EXISTS idx_activity_employee ON ActivityLogs(EmployeeId);
CREATE INDEX IF NOT EXISTS idx_activity_date     ON ActivityLogs(ActivityDate);
CREATE INDEX IF NOT EXISTS idx_activity_type     ON ActivityLogs(ActionType);
CREATE INDEX IF NOT EXISTS idx_activity_repair   ON ActivityLogs(RepairCode);
CREATE INDEX IF NOT EXISTS idx_adminhist_event   ON AdministratorHistory(EventType);

INSERT OR IGNORE INTO Counters (Name, LastValue) VALUES ('RepairCode', 0);
INSERT OR IGNORE INTO Counters (Name, LastValue) VALUES ('ReceiptNumber', 0);

INSERT OR IGNORE INTO Settings (Key, Value) VALUES ('BusinessName', 'My Phone Repair Shop');
INSERT OR IGNORE INTO Settings (Key, Value) VALUES ('BusinessPhone', '');
INSERT OR IGNORE INTO Settings (Key, Value) VALUES ('BusinessAddress', '');
"

End Module