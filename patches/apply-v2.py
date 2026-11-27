from pathlib import Path

root = Path(".")
settings = root / "StudentManagement.Core" / "AppSettingsManager.vb"
repo = root / "StudentManagement.Core" / "ExcelRepository.vb"
admin = root / "StudentManagement.Core" / "AdminAuthManager.vb"

# --- AppSettingsManager: true platform default ---
s = settings.read_text(encoding="utf-8")
s = s.replace(
'''    Public Property DatabaseFolder As String =
        "C:\\StudentsDatabase"''',
'''    Public Property DatabaseFolder As String =
        ""''',
1)

marker = """    '====================================================
    ' DEFAULT SETTINGS
    '====================================================
"""
helper = """    '====================================================
    ' PLATFORM DEFAULT DATABASE FOLDER
    '====================================================
    Private Function DefaultDatabaseFolder() As String

        If OperatingSystem.IsWindows() Then
            Return "C:\\StudentsDatabase"
        End If

        Return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "StudentsDatabase"
        )

    End Function


"""
if "Private Function DefaultDatabaseFolder() As String" not in s:
    if marker not in s:
        raise SystemExit("AppSettings marker not found")
    s = s.replace(marker, helper + marker, 1)

old_default = '''        settings.DatabaseFolder =
            "C:\\StudentsDatabase"'''
if old_default not in s:
    raise SystemExit("CreateDefaultSettings database default not found")
s = s.replace(old_default, '''        settings.DatabaseFolder =
            DefaultDatabaseFolder()''', 1)

old_validate = '''        If settings.DatabaseFolder = "" Then

            settings.DatabaseFolder =
                "C:\\StudentsDatabase"

        End If'''
if old_validate not in s:
    raise SystemExit("ValidateSettings database fallback not found")
s = s.replace(old_validate, '''        If settings.DatabaseFolder = "" Then

            settings.DatabaseFolder =
                DefaultDatabaseFolder()

        End If''', 1)
settings.write_text(s, encoding="utf-8")

# --- ExcelRepository: platform fallback ---
s = repo.read_text(encoding="utf-8")
old = '''            If folder = "" Then

                folder =
                    "C:\\StudentsDatabase"

            End If'''
new = '''            If folder = "" Then

                folder =
                    If(
                        OperatingSystem.IsWindows(),
                        "C:\\StudentsDatabase",
                        Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                            "StudentsDatabase"
                        )
                    )

            End If'''
if old not in s:
    raise SystemExit("ExcelRepository database fallback not found")
s = s.replace(old, new, 1)
repo.write_text(s, encoding="utf-8")

# --- AdminAuthManager: import existing Windows credentials on first Mac launch ---
s = admin.read_text(encoding="utf-8")
old_load = '''    Private Shared Function LoadUsableAccount() As AdminAccount

        '================================================
        ' TRY PRIMARY'''
new_load = '''    Private Shared Function LoadUsableAccount() As AdminAccount

        '================================================
        ' IMPORT EXISTING WINDOWS ADMIN ON FIRST MAC RUN
        '================================================
        TryImportExistingAdmin()


        '================================================
        ' TRY PRIMARY'''
if old_load not in s:
    raise SystemExit("LoadUsableAccount insertion point not found")
s = s.replace(old_load, new_load, 1)

insert_marker = """    '====================================================
    ' MAKE SURE SECURITY PROTECTION EXISTS
"""
method = """    '====================================================
    ' IMPORT AN EXISTING ADMINISTRATOR CREDENTIAL SET
    '
    ' The Windows app stores admin.json separately from
    ' the Excel workbook. For migration, the user can
    ' place admin.json (or its Security backup) beside
    ' the database in Documents/StudentsDatabase.
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
                databaseFolder =
                    If(AppSettingsManager.Current.DatabaseFolder, "").Trim()
            Catch
                databaseFolder = ""
            End Try

            If databaseFolder <> "" Then
                candidates.Add(Path.Combine(databaseFolder, "admin.json"))
                candidates.Add(Path.Combine(databaseFolder, "Security", "admin.backup.json"))
                candidates.Add(Path.Combine(databaseFolder, "admin.backup.json"))
            End If

            If documents <> "" Then
                Dim defaultDb As String =
                    Path.Combine(documents, "StudentsDatabase")

                candidates.Add(Path.Combine(defaultDb, "admin.json"))
                candidates.Add(Path.Combine(defaultDb, "Security", "admin.backup.json"))
                candidates.Add(Path.Combine(defaultDb, "admin.backup.json"))

                Dim legacyCopy As String =
                    Path.Combine(documents, "StudentManagement")

                candidates.Add(Path.Combine(legacyCopy, "admin.json"))
                candidates.Add(Path.Combine(legacyCopy, "Security", "admin.backup.json"))
                candidates.Add(Path.Combine(legacyCopy, "admin.backup.json"))
            End If

            For Each candidate As String In
                candidates.Distinct(StringComparer.OrdinalIgnoreCase)

                Dim imported As AdminAccount =
                    TryLoadAccount(candidate)

                If IsValidAccount(imported) Then
                    SaveAccount(imported)
                    Exit Sub
                End If

            Next

        Catch
            ' Best-effort migration. If no valid legacy
            ' credential exists, normal login behavior continues.
        End Try

    End Sub


"""
if insert_marker not in s:
    raise SystemExit("Admin migration method insertion marker not found")
s = s.replace(insert_marker, method + insert_marker, 1)
admin.write_text(s, encoding="utf-8")

print("Mac V2 migration patches applied successfully")
