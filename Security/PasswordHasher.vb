Imports System.Security.Cryptography

''' <summary>Salted PBKDF2 hashing for admin passwords and employee codes.</summary>
Public Module PasswordHasher

    Private Const Iterations As Integer = 100000
    Private Const KeySize As Integer = 32

    Public Function GenerateSalt() As String
        Return Convert.ToBase64String(RandomNumberGenerator.GetBytes(16))
    End Function

    Public Function HashSecret(secret As String, saltBase64 As String) As String
        Dim salt As Byte() = Convert.FromBase64String(saltBase64)
        Dim hash As Byte() = Rfc2898DeriveBytes.Pbkdf2(
            secret, salt, Iterations, HashAlgorithmName.SHA256, KeySize)
        Return Convert.ToBase64String(hash)
    End Function

    Public Function Verify(secret As String, saltBase64 As String,
                           expectedHashBase64 As String) As Boolean
        Dim actual As Byte() = Convert.FromBase64String(HashSecret(secret, saltBase64))
        Dim expected As Byte() = Convert.FromBase64String(expectedHashBase64)
        Return CryptographicOperations.FixedTimeEquals(actual, expected)
    End Function

End Module