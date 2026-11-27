Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.Json


Public NotInheritable Class AdminAuthManager

    '====================================================
    ' PASSWORD SECURITY
    '====================================================
    Private Const Iterations As Integer =
        210000

    Private Const SaltSize As Integer =
        32

    Private Const HashSize As Integer =
        32


    '====================================================
    ' ADMIN STORAGE
    '====================================================
    Private Shared ReadOnly AdminFolder As String =
        Path.Combine(
            Environment.GetFolderPath(
                Environment.SpecialFolder.LocalApplicationData
            ),
            "StudentManagement"
        )


    Private Shared ReadOnly AdminFile As String =
        Path.Combine(
            AdminFolder,
            "admin.json"
        )


    '====================================================
    ' SECURITY BACKUP
    '====================================================
    Private Shared ReadOnly SecurityFolder As String =
        Path.Combine(
            AdminFolder,
            "Security"
        )


    Private Shared ReadOnly AdminBackupFile As String =
        Path.Combine(
            SecurityFolder,
            "admin.backup.json"
        )


    '====================================================
    ' CROSS-PLATFORM SECURITY MARKER
    '====================================================
    Private Shared ReadOnly ProvisionedMarkerFile As String =
        Path.Combine(SecurityFolder, "admin.provisioned")


    '====================================================
    ' STORED ADMIN MODEL
    '====================================================
    Private Class AdminAccount

        Public Property Username As String =
            ""

        Public Property Salt As String =
            ""

        Public Property PasswordHash As String =
            ""

        Public Property CreatedAt As DateTime

    End Class


    '====================================================
    ' NO INSTANCE
    '====================================================
    Private Sub New()

    End Sub


    '====================================================
    ' DOES ADMIN EXIST?
    '
    ' IMPORTANT:
    ' If the system was previously provisioned but the
    ' credential files are gone, this still returns TRUE.
    '
    ' That prevents FrmLogin from returning to first-run
    ' CREATE ADMIN ACCOUNT mode.
    '====================================================
    Public Shared Function AdminExists() As Boolean

        Try

            Dim account As AdminAccount =
                LoadUsableAccount()


            If account IsNot Nothing Then

                Return True

            End If


            '============================================
            ' ADMIN WAS PREVIOUSLY CREATED.
            '
            ' Do NOT allow first-run setup again merely
            ' because files were deleted.
            '============================================
            If IsProvisionedMarkerSet() Then

                Return True

            End If


            Return False


        Catch

            '============================================
            ' Even if credential loading fails, preserve
            ' the security state if Windows says an admin
            ' was already provisioned.
            '============================================
            Return IsProvisionedMarkerSet()

        End Try

    End Function


    '====================================================
    ' HAS THIS INSTALLATION EVER BEEN PROVISIONED?
    '====================================================
    Public Shared Function IsAdminProvisioned() As Boolean

        Try

            If IsProvisionedMarkerSet() Then

                Return True

            End If


            If IsValidAccountFile(
                AdminFile
            ) Then

                Return True

            End If


            If IsValidAccountFile(
                AdminBackupFile
            ) Then

                Return True

            End If


            Return False


        Catch

            Return False

        End Try

    End Function


    '====================================================
    ' CREATE FIRST ADMIN
    '====================================================
    Public Shared Sub CreateAdmin(
        username As String,
        password As String
    )

        username =
            If(
                username,
                ""
            ).Trim()


        password =
            If(
                password,
                ""
            )


        If username = "" Then

            Throw New Exception(
                "Administrator username is required."
            )

        End If


        If username.Length < 3 Then

            Throw New Exception(
                "Administrator username must contain at least 3 characters."
            )

        End If


        ValidateNewPassword(
            password
        )


        '============================================
        ' NEVER ALLOW ANOTHER ADMIN TO BE CREATED
        ' IF THIS INSTALLATION WAS ALREADY SET UP.
        '============================================
        If IsAdminProvisioned() Then

            Throw New Exception(
                "An administrator account has already been created for this application."
            )

        End If


        Directory.CreateDirectory(
            AdminFolder
        )


        Directory.CreateDirectory(
            SecurityFolder
        )


        Dim salt As Byte() =
            RandomNumberGenerator.GetBytes(
                SaltSize
            )


        Dim passwordHash As Byte() =
            HashPassword(
                password,
                salt
            )


        Dim account As New AdminAccount With {
            .Username = username,
            .Salt = Convert.ToBase64String(
                salt
            ),
            .PasswordHash = Convert.ToBase64String(
                passwordHash
            ),
            .CreatedAt = DateTime.Now
        }


        SaveAccount(
            account
        )

    End Sub


    '====================================================
    ' VERIFY LOGIN
    '====================================================
    Public Shared Function VerifyCredentials(
        username As String,
        password As String
    ) As Boolean

        Try

            username =
                If(
                    username,
                    ""
                ).Trim()


            password =
                If(
                    password,
                    ""
                )


            If username = "" OrElse
               password = "" Then

                Return False

            End If


            Dim account As AdminAccount =
                LoadUsableAccount()


            If account Is Nothing Then

                Return False

            End If


            If Not account.Username.Equals(
                username,
                StringComparison.OrdinalIgnoreCase
            ) Then

                Return False

            End If


            Dim salt As Byte() =
                Convert.FromBase64String(
                    account.Salt
                )


            Dim storedHash As Byte() =
                Convert.FromBase64String(
                    account.PasswordHash
                )


            Dim enteredHash As Byte() =
                HashPassword(
                    password,
                    salt
                )


            If storedHash.Length <>
               enteredHash.Length Then

                Return False

            End If


            Return CryptographicOperations.FixedTimeEquals(
                storedHash,
                enteredHash
            )


        Catch

            Return False

        End Try

    End Function


    '====================================================
    ' CHANGE ADMIN PASSWORD
    '====================================================
    Public Shared Sub ChangePassword(
        currentPassword As String,
        newPassword As String
    )

        Dim account As AdminAccount =
            LoadUsableAccount()


        If account Is Nothing Then

            Throw New Exception(
                "Administrator security information could not be found."
            )

        End If


        If Not VerifyCredentials(
            account.Username,
            currentPassword
        ) Then

            Throw New Exception(
                "The current administrator password is incorrect."
            )

        End If


        ValidateNewPassword(
            newPassword
        )


        Dim newSalt As Byte() =
            RandomNumberGenerator.GetBytes(
                SaltSize
            )


        Dim newHash As Byte() =
            HashPassword(
                newPassword,
                newSalt
            )


        account.Salt =
            Convert.ToBase64String(
                newSalt
            )


        account.PasswordHash =
            Convert.ToBase64String(
                newHash
            )


        SaveAccount(
            account
        )

    End Sub


    '====================================================
    ' GET ADMIN USERNAME
    '====================================================
    Public Shared Function GetAdminUsername() As String

        Try

            Dim account As AdminAccount =
                LoadUsableAccount()


            If account Is Nothing Then

                Return ""

            End If


            Return If(
                account.Username,
                ""
            ).Trim()


        Catch

            Return ""

        End Try

    End Function


    '====================================================
    ' LOAD USABLE ACCOUNT
    '
    ' 1. Try primary admin.json
    ' 2. If missing/damaged, try security backup
    ' 3. Restore primary automatically from backup
    '====================================================
    Private Shared Function LoadUsableAccount() As AdminAccount

        '================================================
        ' IMPORT EXISTING WINDOWS ADMIN ON FIRST MAC RUN
        '================================================
        TryImportExistingAdmin()


        '================================================
        ' TRY PRIMARY
        '================================================
        Dim account As AdminAccount =
            TryLoadAccount(
                AdminFile
            )


        If IsValidAccount(
            account
        ) Then

            EnsureProtectionState(
                account
            )


            Return account

        End If


        '================================================
        ' TRY SECURITY BACKUP
        '================================================
        Dim backupAccount As AdminAccount =
            TryLoadAccount(
                AdminBackupFile
            )


        If IsValidAccount(
            backupAccount
        ) Then

            '============================================
            ' RESTORE DELETED/DAMAGED PRIMARY FILE
            '============================================
            WriteAccountFile(
                AdminFile,
                backupAccount
            )


            EnsureProtectionState(
                backupAccount
            )


            Return backupAccount

        End If


        Return Nothing

    End Function


    '====================================================
    ' IMPORT AN EXISTING ADMINISTRATOR CREDENTIAL SET
    '
    ' This supports migration from the Windows edition.
    ' The user may place admin.json and/or the Security
    ' backup beside the Excel database, or copy the old
    ' StudentManagement folder into Documents.
    '====================================================
    Private Shared Sub TryImportExistingAdmin()

        Try
            If IsValidAccountFile(AdminFile) OrElse
               IsValidAccountFile(AdminBackupFile) Then
                Exit Sub
            End If

            Dim candidates As New List(Of String)()

            Dim documents As String =
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

            Dim databaseFolder As String = ""
            Try
                databaseFolder = If(AppSettingsManager.Current.DatabaseFolder, "").Trim()
            Catch
                databaseFolder = ""
            End Try

            If databaseFolder <> "" Then
                candidates.Add(Path.Combine(databaseFolder, "admin.json"))
                candidates.Add(Path.Combine(databaseFolder, "Security", "admin.backup.json"))
                candidates.Add(Path.Combine(databaseFolder, "admin.backup.json"))
            End If

            If documents <> "" Then
                Dim defaultDb As String = Path.Combine(documents, "StudentsDatabase")
                candidates.Add(Path.Combine(defaultDb, "admin.json"))
                candidates.Add(Path.Combine(defaultDb, "Security", "admin.backup.json"))
                candidates.Add(Path.Combine(defaultDb, "admin.backup.json"))

                Dim legacyCopy As String = Path.Combine(documents, "StudentManagement")
                candidates.Add(Path.Combine(legacyCopy, "admin.json"))
                candidates.Add(Path.Combine(legacyCopy, "Security", "admin.backup.json"))
                candidates.Add(Path.Combine(legacyCopy, "admin.backup.json"))
            End If

            For Each candidate As String In candidates.Distinct(StringComparer.OrdinalIgnoreCase)
                Dim imported As AdminAccount = TryLoadAccount(candidate)
                If IsValidAccount(imported) Then
                    SaveAccount(imported)
                    Exit Sub
                End If
            Next

        Catch
            ' Migration is best-effort. Normal authentication
            ' behavior continues when no valid legacy file exists.
        End Try

    End Sub


    '====================================================
    ' MAKE SURE SECURITY PROTECTION EXISTS
    '
    ' This also automatically upgrades an administrator
    ' created using the previous version of the app.
    '====================================================
    Private Shared Sub EnsureProtectionState(
        account As AdminAccount
    )

        If account Is Nothing Then

            Exit Sub

        End If


        Directory.CreateDirectory(
            SecurityFolder
        )


        '================================================
        ' CREATE BACKUP IF REQUIRED
        '================================================
        If Not IsValidAccountFile(
            AdminBackupFile
        ) Then

            WriteAccountFile(
                AdminBackupFile,
                account
            )

        End If


        '================================================
        ' CREATE WINDOWS PROVISIONED MARKER
        '================================================
        SetProvisionedMarker()

    End Sub


    '====================================================
    ' VALIDATE PASSWORD
    '====================================================
    Private Shared Sub ValidateNewPassword(
        password As String
    )

        password =
            If(
                password,
                ""
            )


        If password.Length < 8 Then

            Throw New Exception(
                "Administrator password must contain at least 8 characters."
            )

        End If


        Dim hasUpper As Boolean =
            False

        Dim hasLower As Boolean =
            False

        Dim hasNumber As Boolean =
            False


        For Each ch As Char In password

            If Char.IsUpper(
                ch
            ) Then

                hasUpper =
                    True


            ElseIf Char.IsLower(
                ch
            ) Then

                hasLower =
                    True


            ElseIf Char.IsDigit(
                ch
            ) Then

                hasNumber =
                    True

            End If

        Next


        If Not hasUpper OrElse
           Not hasLower OrElse
           Not hasNumber Then

            Throw New Exception(
                "Administrator password must contain at least one uppercase letter, one lowercase letter and one number."
            )

        End If

    End Sub


    '====================================================
    ' PASSWORD HASH
    '====================================================
    Private Shared Function HashPassword(
        password As String,
        salt As Byte()
    ) As Byte()

        Return Rfc2898DeriveBytes.Pbkdf2(
            Encoding.UTF8.GetBytes(
                password
            ),
            salt,
            Iterations,
            HashAlgorithmName.SHA256,
            HashSize
        )

    End Function


    '====================================================
    ' SAVE ACCOUNT
    '
    ' BOTH PRIMARY AND SECURITY BACKUP ARE UPDATED.
    '====================================================
    Private Shared Sub SaveAccount(
        account As AdminAccount
    )

        Directory.CreateDirectory(
            AdminFolder
        )


        Directory.CreateDirectory(
            SecurityFolder
        )


        WriteAccountFile(
            AdminFile,
            account
        )


        WriteAccountFile(
            AdminBackupFile,
            account
        )


        SetProvisionedMarker()

    End Sub


    '====================================================
    ' WRITE ACCOUNT FILE
    '====================================================
    Private Shared Sub WriteAccountFile(
        filePath As String,
        account As AdminAccount
    )

        Dim folder As String =
            Path.GetDirectoryName(
                filePath
            )


        If Not String.IsNullOrWhiteSpace(
            folder
        ) Then

            Directory.CreateDirectory(
                folder
            )

        End If


        Dim options As New JsonSerializerOptions With {
            .WriteIndented = True
        }


        Dim json As String =
            JsonSerializer.Serialize(
                account,
                options
            )


        Dim tempFile As String =
            filePath &
            ".tmp"


        File.WriteAllText(
            tempFile,
            json,
            Encoding.UTF8
        )


        If File.Exists(
            filePath
        ) Then

            File.Delete(
                filePath
            )

        End If


        File.Move(
            tempFile,
            filePath
        )

    End Sub


    '====================================================
    ' TRY LOAD ACCOUNT FILE
    '====================================================
    Private Shared Function TryLoadAccount(
        filePath As String
    ) As AdminAccount

        Try

            If Not File.Exists(
                filePath
            ) Then

                Return Nothing

            End If


            Dim json As String =
                File.ReadAllText(
                    filePath,
                    Encoding.UTF8
                )


            If String.IsNullOrWhiteSpace(
                json
            ) Then

                Return Nothing

            End If


            Return JsonSerializer.Deserialize(
                Of AdminAccount
            )(
                json
            )


        Catch

            Return Nothing

        End Try

    End Function


    '====================================================
    ' VALID ACCOUNT FILE?
    '====================================================
    Private Shared Function IsValidAccountFile(
        filePath As String
    ) As Boolean

        Dim account As AdminAccount =
            TryLoadAccount(
                filePath
            )


        Return IsValidAccount(
            account
        )

    End Function


    '====================================================
    ' VALID ACCOUNT?
    '====================================================
    Private Shared Function IsValidAccount(
        account As AdminAccount
    ) As Boolean

        If account Is Nothing Then

            Return False

        End If


        If String.IsNullOrWhiteSpace(
            account.Username
        ) Then

            Return False

        End If


        If String.IsNullOrWhiteSpace(
            account.Salt
        ) Then

            Return False

        End If


        If String.IsNullOrWhiteSpace(
            account.PasswordHash
        ) Then

            Return False

        End If


        Try

            Dim salt As Byte() =
                Convert.FromBase64String(
                    account.Salt
                )


            Dim hash As Byte() =
                Convert.FromBase64String(
                    account.PasswordHash
                )


            If salt.Length = 0 OrElse
               hash.Length = 0 Then

                Return False