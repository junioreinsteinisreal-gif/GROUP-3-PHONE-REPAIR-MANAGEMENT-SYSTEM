''' <summary>Creates the tables and column used by customer messages, reviews and tips.
''' Safe to run on every start.</summary>
Public Module SchemaUpdates

    Public Sub Apply()
        DatabaseManager.ExecuteNonQuery(NewTablesSql)

        Dim emailColumns As Integer = Convert.ToInt32(DatabaseManager.ExecuteScalar(
            "SELECT COUNT(*) FROM pragma_table_info('Repairs') WHERE name = 'CustomerEmail'"))
        If emailColumns = 0 Then
            DatabaseManager.ExecuteNonQuery("ALTER TABLE Repairs ADD COLUMN CustomerEmail TEXT")
        End If
    End Sub

    Private ReadOnly NewTablesSql As String = "
CREATE TABLE IF NOT EXISTS Notifications (
    NotificationId  INTEGER PRIMARY KEY AUTOINCREMENT,
    RepairCode      TEXT NOT NULL,
    Kind            TEXT NOT NULL,
    Channel         TEXT NOT NULL,
    Recipient       TEXT NOT NULL,
    Subject         TEXT,
    Body            TEXT NOT NULL,
    Status          TEXT NOT NULL DEFAULT 'Pending',
    Attempts        INTEGER NOT NULL DEFAULT 0,
    LastError       TEXT,
    CreatedAt       TEXT NOT NULL,
    SentAt          TEXT,
    UNIQUE (RepairCode, Kind, Channel)
);

CREATE TABLE IF NOT EXISTS FeedbackInvites (
    RepairCode  TEXT PRIMARY KEY,
    Pin         TEXT NOT NULL,
    CreatedAt   TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS Feedback (
    FeedbackId    INTEGER PRIMARY KEY AUTOINCREMENT,
    RepairCode    TEXT NOT NULL UNIQUE,
    CustomerName  TEXT NOT NULL,
    EmployeeId    INTEGER NOT NULL,
    Rating        INTEGER NOT NULL CHECK (Rating BETWEEN 1 AND 5),
    Comment       TEXT,
    Verified      INTEGER NOT NULL DEFAULT 0,
    RecordedAt    TEXT NOT NULL,
    FOREIGN KEY (EmployeeId) REFERENCES Employees(EmployeeId)
);

CREATE TABLE IF NOT EXISTS Tips (
    TipId       INTEGER PRIMARY KEY AUTOINCREMENT,
    RepairCode  TEXT NOT NULL,
    FeedbackId  INTEGER,
    EmployeeId  INTEGER NOT NULL,
    Amount      REAL NOT NULL CHECK (Amount > 0),
    Method      TEXT NOT NULL,
    ReceivedAt  TEXT NOT NULL,
    FOREIGN KEY (EmployeeId) REFERENCES Employees(EmployeeId)
);

CREATE INDEX IF NOT EXISTS idx_feedback_emp ON Feedback(EmployeeId);
CREATE INDEX IF NOT EXISTS idx_tips_emp ON Tips(EmployeeId);
CREATE INDEX IF NOT EXISTS idx_notif_status ON Notifications(Status);

INSERT OR IGNORE INTO Settings (Key, Value) VALUES ('SmsEnabled', '0');
INSERT OR IGNORE INTO Settings (Key, Value) VALUES ('SmsUrl', '');
INSERT OR IGNORE INTO Settings (Key, Value) VALUES ('SmsSender', '');
INSERT OR IGNORE INTO Settings (Key, Value) VALUES ('CountryCode', '233');
INSERT OR IGNORE INTO Settings (Key, Value) VALUES ('EmailEnabled', '0');
INSERT OR IGNORE INTO Settings (Key, Value) VALUES ('SmtpHost', '');
INSERT OR IGNORE INTO Settings (Key, Value) VALUES ('SmtpPort', '587');
INSERT OR IGNORE INTO Settings (Key, Value) VALUES ('SmtpSsl', '1');
INSERT OR IGNORE INTO Settings (Key, Value) VALUES ('SmtpUser', '');
INSERT OR IGNORE INTO Settings (Key, Value) VALUES ('SmtpPassword', '');
INSERT OR IGNORE INTO Settings (Key, Value) VALUES ('SmtpFrom', '');
"

End Module