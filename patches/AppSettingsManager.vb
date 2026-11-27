Imports System.IO
Imports System.Text.Json


'========================================================
' APPLICATION SETTINGS MODEL
'========================================================
Public Class StudentAppSettings

    '====================================================
    ' SCHOOL / APPLICATION
    '====================================================
    Public Property SchoolName As String =
        "STUDENT MANAGEMENT SYSTEM"


    '====================================================
    ' DATABASE
    '====================================================
    Public Property DatabaseFolder As String =
        ""


    '====================================================
    ' BACKUPS
    '====================================================
    Public Property AutomaticBackupEnabled As Boolean =
        True


    Public Property BackupRetentionCount As Integer =
        30


    '====================================================
    ' EXPORTS
    '====================================================
    Public Property DefaultExportFolder As String =
        Environment.GetFolderPath(
            Environment.SpecialFolder.MyDocuments
        )

End Class



'========================================================
' SETTINGS MANAGER
'========================================================
Public Module AppSettingsManager

    '====================================================
    ' STORAGE LOCATION
    '
    ' Example:
    '
    ' C:\Users\<User>\AppData\Local\
    ' StudentManagement\settings.json
    '====================================================
    Private ReadOnly SettingsFolder As String =
        Path.Combine(
            Environment.GetFolderPath(
                Environment.SpecialFolder.LocalApplicationData
            ),
            "StudentManagement"
        )


    Private ReadOnly SettingsFile As String =
        Path.Combine(
            SettingsFolder,
            "settings.json"
        )


    '====================================================
    ' CURRENT SETTINGS
    '====================================================
    Private _current As StudentAppSettings =
        Nothing


    Public ReadOnly Property Current As StudentAppSettings

        Get

            If _current Is Nothing Then

                Load()

            End If


            Return _current

        End Get

    End Property


    '====================================================
    ' SETTINGS FILE PATH
    '
    ' Useful later in Settings screen.
    '====================================================
    Public ReadOnly Property SettingsFilePath As String

        Get

            Return SettingsFile

        End Get

    End Property


    '====================================================
    ' LOAD SETTINGS
    '====================================================
    Public Sub Load()

        Try

            EnsureSettingsFolder()


            '================================================
            ' FIRST RUN
            '================================================
            If Not File.Exists(
                SettingsFile
            ) Then

                _current =
                    CreateDefaultSettings()


                Save()


                Exit Sub

            End If


            '================================================
            ' READ JSON
            '================================================
            Dim json As String =
                File.ReadAllText(
                    SettingsFile
                )


            If String.IsNullOrWhiteSpace(
                json
            ) Then

                _current =
                    CreateDefaultSettings()


                Save()


                Exit Sub

            End If


            '================================================
            ' DESERIALIZE
            '================================================
            Dim options As New JsonSerializerOptions()


            options.PropertyNameCaseInsensitive =
                True


            Dim loadedSettings As StudentAppSettings =
                JsonSerializer.Deserialize(
                    Of StudentAppSettings
                )(
                    json,
                    options
                )


            If loadedSettings Is Nothing Then

                Throw New Exception(
                    "The settings file could not be read."
                )

            End If


            '================================================
            ' VALIDATE
            '================================================
            _current =
                ValidateSettings(
                    loadedSettings
                )


        Catch

            '================================================
            ' BAD / DAMAGED SETTINGS FILE
            '
            ' Keep a copy before restoring defaults.
            '================================================
            BackupDamagedSettingsFile()


            _current =
                CreateDefaultSettings()


            Try

                Save()

            Catch

                'Do not crash the whole application just
                'because the settings file could not be saved.

            End Try

        End Try

    End Sub


    '====================================================
    ' SAVE SETTINGS
    '====================================================
    Public Sub Save()

        If _current Is Nothing Then

            _current =
                CreateDefaultSettings()

        End If


        _current =
            ValidateSettings(
                _current
            )


        EnsureSettingsFolder()


        Dim options As New JsonSerializerOptions()


        options.WriteIndented =
            True


        Dim json As String =
            JsonSerializer.Serialize(
                _current,
                options
            )


        '================================================
        ' WRITE SAFELY THROUGH TEMP FILE
        '================================================
        Dim tempFile As String =
            SettingsFile &
            ".tmp"


        File.WriteAllText(
            tempFile,
            json
        )


        If File.Exists(
            SettingsFile
        ) Then

            File.Delete(
                SettingsFile
            )

        End If


        File.Move(
            tempFile,
            SettingsFile
        )

    End Sub


    '====================================================
    ' SAVE NEW VALUES
    '
    ' This will be called later by FrmSettings.
    '====================================================
    Public Sub UpdateSettings(
        schoolName As String,
        databaseFolder As String,
        automaticBackupEnabled As Boolean,
        backupRetentionCount As Integer,
        defaultExportFolder As String
    )

        Dim newSettings As New StudentAppSettings()


        newSettings.SchoolName =
            schoolName


        newSettings.DatabaseFolder =
            databaseFolder


        newSettings.AutomaticBackupEnabled =
            automaticBackupEnabled


        newSettings.BackupRetentionCount =
            backupRetentionCount


        newSettings.DefaultExportFolder =
            defaultExportFolder


        _current =
            ValidateSettings(
                newSettings
            )


        Save()

    End Sub


    '====================================================
    ' RESET SETTINGS
    '====================================================
    Public Sub ResetToDefaults()

        _current =
            CreateDefaultSettings()


        Save()

    End Sub


    '====================================================
    ' PLATFORM DEFAULT DATABASE FOLDER
    '====================================================
    Private Function DefaultDatabaseFolder() As String

        If OperatingSystem.IsWindows() Then
            Return "C:\StudentsDatabase"
        End If

        Return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "StudentsDatabase"
        )

    End Function


    '====================================================
    ' DEFAULT SETTINGS
    '====================================================
    Private Function CreateDefaultSettings() As StudentAppSettings

        Dim settings As New StudentAppSettings()


        settings.SchoolName =
            "STUDENT MANAGEMENT SYSTEM"


        settings.DatabaseFolder =
            DefaultDatabaseFolder()


        settings.AutomaticBackupEnabled =
            True


        settings.BackupRetentionCount =
            30


        settings.DefaultExportFolder =
            Environment.GetFolderPath(
                Environment.SpecialFolder.MyDocuments
            )


        Return settings

    End Function


    '====================================================
    ' VALIDATE SETTINGS
    '====================================================
    Private Function ValidateSettings(
        settings As StudentAppSettings
    ) As StudentAppSettings

        If settings Is Nothing Then

            Return CreateDefaultSettings()

        End If


        '================================================
        ' SCHOOL NAME
        '================================================
        settings.SchoolName =
            If(
                settings.SchoolName,
                ""
            ).Trim()


        If settings.SchoolName = "" Then

            settings.SchoolName =
                "STUDENT MANAGEMENT SYSTEM"

        End If


        '================================================
        ' DATABASE FOLDER
        '================================================
        settings.DatabaseFolder =
            NormalizeFolder(
                settings.DatabaseFolder
            )


        If settings.DatabaseFolder = "" Then

            settings.DatabaseFolder =
                DefaultDatabaseFolder()

        End If


        '================================================
        ' BACKUP RETENTION
        '================================================
        If settings.BackupRetentionCount < 1 Then

            settings.BackupRetentionCount =
                1

        End If


        If settings.BackupRetentionCount > 365 Then

            settings.BackupRetentionCount =
                365

        End If


        '================================================
        ' EXPORT FOLDER
        '================================================
        settings.DefaultExportFolder =
            NormalizeFolder(
                settings.DefaultExportFolder
            )


        If settings.DefaultExportFolder = "" Then

            settings.DefaultExportFolder =
                Environment.GetFolderPath(
                    Environment.SpecialFolder.MyDocuments
                )

        End If


        Return settings

    End Function


    '====================================================
    ' NORMALIZE FOLDER
    '====================================================
    Private Function NormalizeFolder(
        folder As String
    ) As String

        Dim value As String =
            If(
                folder,
                ""
            ).Trim()


        If value = "" Then

            Return ""

        End If


        Try

            value =
                Path.GetFullPath(
                    value
                )

        Catch

            Return ""

        End Try


        Return value.TrimEnd(
            Path.DirectorySeparatorChar,
            Path.AltDirectorySeparatorChar
        )

    End Function


    '====================================================
    ' SETTINGS DIRECTORY
    '====================================================
    Private Sub EnsureSettingsFolder()

        If Not Directory.Exists(
            SettingsFolder
        ) Then

            Directory.CreateDirectory(
                SettingsFolder
            )

        End If

    End Sub


    '====================================================
    ' BACKUP DAMAGED SETTINGS
    '====================================================
    Private Sub BackupDamagedSettingsFile()

        Try

            If Not File.Exists(
                SettingsFile
            ) Then

                Exit Sub

            End If


            EnsureSettingsFolder()


            Dim damagedFile As String =
                Path.Combine(
                    SettingsFolder,
                    "settings_corrupt_" &
                    DateTime.Now.ToString(
                        "yyyyMMdd_HHmmss"
                    ) &
                    ".json"
                )


            File.Copy(
                SettingsFile,
                damagedFile,
                True
            )

        Catch

            'Never let a damaged settings file stop
            'the application from starting.

        End Try

    End Sub

End Module