Imports ClosedXML.Excel
Imports System.IO
Imports System.Linq
Imports System.Collections.Generic
Imports System.Data
Imports System.Globalization


Public Module ExcelRepository

    '====================================================
    ' DATABASE FOLDER
    '====================================================
    Private ReadOnly Property DatabaseFolder As String

        Get

            Dim folder As String =
                If(
                    AppSettingsManager.Current.DatabaseFolder,
                    ""
                ).Trim()


            If folder = "" Then

                folder =
                    If(
                        OperatingSystem.IsWindows(),
                        "C:\StudentsDatabase",
                        Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                            "StudentsDatabase"
                        )
                    )

            End If


            Return Path.GetFullPath(
                folder
            )

        End Get

    End Property


    Private Const WorksheetName As String =
        "Main Sheet"


    Private Const DeletedWorksheetName As String =
        "Deleted"


    Private Const FirstDataRow As Integer =
        2


    Private Const LastDataRow As Integer =
        2000


    '====================================================
    ' STANDARD STUDENT FIELDS
    '====================================================
    Private ReadOnly StudentFields As String() = {
        "Student No",
        "Name",
        "Gender",
        "Iquama",
        "Date of Birth",
        "Grade",
        "Room",
        "Phone No",
        "Phone No2",
        "Phone No3",
        "Phone No4"
    }


    '====================================================
    ' AUTOMATICALLY FIND EXCEL DATABASE
    '====================================================
    Public ReadOnly Property WorkbookPath As String

        Get

            Dim currentDatabaseFolder As String =
                DatabaseFolder


            If Not Directory.Exists(
                currentDatabaseFolder
            ) Then

                Throw New DirectoryNotFoundException(
                    "Database folder not found:" &
                    Environment.NewLine &
                    currentDatabaseFolder
                )

            End If


            Dim files =
                Directory.GetFiles(
                    currentDatabaseFolder,
                    "*.xlsx",
                    SearchOption.TopDirectoryOnly
                ).
                Concat(
                    Directory.GetFiles(
                        currentDatabaseFolder,
                        "*.xlsm",
                        SearchOption.TopDirectoryOnly
                    )
                ).
                Where(
                    Function(f)

                        Return Not Path.
                            GetFileName(f).
                            StartsWith(
                                "~$",
                                StringComparison.OrdinalIgnoreCase
                            )

                    End Function
                ).
                ToArray()


            If files.Length = 0 Then

                Throw New FileNotFoundException(
                    "No Excel workbook was found in:" &
                    Environment.NewLine &
                    currentDatabaseFolder
                )

            End If


            If files.Length > 1 Then

                Throw New Exception(
                    "More than one Excel workbook was found in:" &
                    Environment.NewLine &
                    currentDatabaseFolder &
                    Environment.NewLine &
                    Environment.NewLine &
                    "Please keep only the student database workbook in this folder."
                )

            End If


            Return files(0)

        End Get

    End Property


    '====================================================
    ' GET COLUMN NUMBERS
    '====================================================
    Private Function GetHeaders(
        ws As IXLWorksheet
    ) As Dictionary(Of String, Integer)

        Dim headers As New Dictionary(
            Of String,
            Integer
        )(
            StringComparer.OrdinalIgnoreCase
        )


        Dim lastColumnUsed =
            ws.LastColumnUsed()


        If lastColumnUsed Is Nothing Then

            Throw New Exception(
                "No columns were found in worksheet '" &
                ws.Name &
                "'."
            )

        End If


        Dim lastColumn As Integer =
            lastColumnUsed.ColumnNumber()


        For columnNumber As Integer =
            1 To lastColumn

            Dim header As String =
                ws.Cell(
                    1,
                    columnNumber
                ).
                GetString().
                Trim()


            If header <> "" Then

                headers(header) =
                    columnNumber

            End If

        Next


        Return headers

    End Function


    '====================================================
    ' GET TABLE FIELD INDEXES
    '====================================================
    Private Function GetTableFieldIndexes(
        table As IXLTable
    ) As Dictionary(Of String, Integer)

        Dim result As New Dictionary(
            Of String,
            Integer
        )(
            StringComparer.OrdinalIgnoreCase
        )


        Dim fieldIndex As Integer =
            1


        For Each field As IXLTableField In table.Fields

            Dim fieldName As String =
                field.Name.Trim()


            If fieldName <> "" Then

                result(fieldName) =
                    fieldIndex

            End If


            fieldIndex +=
                1

        Next


        Return result

    End Function


    '====================================================
    ' CREATE STANDARD DATATABLE
    '====================================================
    Private Function CreateStudentDataTable() As DataTable

        Dim dt As New DataTable()


        For Each fieldName As String In StudentFields

            dt.Columns.Add(
                fieldName,
                GetType(String)
            )

        Next


        Return dt

    End Function


    '====================================================
    ' VALIDATE MAIN HEADERS
    '====================================================
    Private Sub ValidateMainHeaders(
        columns As Dictionary(Of String, Integer)
    )

        For Each fieldName As String In StudentFields

            If Not columns.ContainsKey(
                fieldName
            ) Then

                Throw New Exception(
                    "Column not found: " &
                    fieldName
                )

            End If

        Next

    End Sub


    '====================================================
    ' VALIDATE DELETED TABLE
    '====================================================
    Private Sub ValidateDeletedFields(
        fields As Dictionary(Of String, Integer)
    )

        For Each fieldName As String In StudentFields

            If fieldName.Equals(
                "Date of Birth",
                StringComparison.OrdinalIgnoreCase
            ) Then

                Continue For

            End If


            If Not fields.ContainsKey(
                fieldName
            ) Then

                Throw New Exception(
                    "Column '" &
                    fieldName &
                    "' was not found in the Deleted table."
                )

            End If

        Next

    End Sub


    '====================================================
    ' SAFETY BACKUP
    '====================================================
    Private Sub CreateSafetyBackup(
        filePath As String
    )

        Try

            DatabaseBackup.CreateBackup(
                filePath
            )

        Catch ex As Exception

            Throw New Exception(
                "The database operation was cancelled because the safety backup could not be created." &
                Environment.NewLine &
                Environment.NewLine &
                ex.Message,
                ex
            )

        End Try

    End Sub


    '====================================================
    ' NORMALIZE TEXT
    '====================================================
    Private Function NormalizeUpper(
        value As String
    ) As String

        Return If(
            value,
            ""
        ).
        Trim().
        ToUpperInvariant()

    End Function


    Private Function NormalizeText(
        value As String
    ) As String

        Return If(
            value,
            ""
        ).
        Trim()

    End Function


    '====================================================
    ' DATE OF BIRTH
    '====================================================
    Private Function TryNormalizeDate(
        value As String,
        ByRef normalizedDate As String
    ) As Boolean

        normalizedDate =
            ""


        Dim cleanValue As String =
            NormalizeText(
                value
            )


        If cleanValue = "" Then

            Return True

        End If


        Dim parsedDate As DateTime


        Dim acceptedFormats As String() = {
            "dd/MM/yyyy",
            "d/M/yyyy",
            "dd-MM-yyyy",
            "d-M-yyyy",
            "yyyy-MM-dd"
        }


        If DateTime.TryParseExact(
            cleanValue,
            acceptedFormats,
            CultureInfo.InvariantCulture,
            DateTimeStyles.None,
            parsedDate
        ) Then

            normalizedDate =
                parsedDate.ToString(
                    "dd/MM/yyyy",
                    CultureInfo.InvariantCulture
                )


            Return True

        End If


        Return False

    End Function


    Private Function NormalizeDateOrThrow(
        value As String
    ) As String

        Dim normalizedDate As String =
            ""


        If Not TryNormalizeDate(
            value,
            normalizedDate
        ) Then

            Throw New Exception(
                "Invalid Date of Birth: " &
                value &
                Environment.NewLine &
                Environment.NewLine &
                "Please use the format dd/MM/yyyy."
            )

        End If


        Return normalizedDate

    End Function


    Private Function GetDateCellValue(
        cell As IXLCell
    ) As String

        If cell Is Nothing OrElse
           cell.IsEmpty() Then

            Return ""

        End If


        Dim dateValue As DateTime


        If cell.TryGetValue(
            dateValue
        ) Then

            Return dateValue.ToString(
                "dd/MM/yyyy",
                CultureInfo.InvariantCulture
            )

        End If


        Dim formattedValue As String =
            cell.GetFormattedString().
            Trim()


        Dim normalizedDate As String =
            ""


        If TryNormalizeDate(
            formattedValue,
            normalizedDate
        ) Then

            Return normalizedDate

        End If


        Return formattedValue

    End Function


    Private Function GetDateCellObject(
        cell As IXLCell
    ) As Object

        If cell Is Nothing OrElse
           cell.IsEmpty() Then

            Return ""

        End If


        Dim dateValue As DateTime


        If cell.TryGetValue(
            dateValue
        ) Then

            Return dateValue.Date

        End If


        Dim dateText As String =
            GetDateCellValue(
                cell
            )


        If dateText = "" Then

            Return ""

        End If


        Dim normalizedDate As String =
            ""


        If TryNormalizeDate(
            dateText,
            normalizedDate
        ) Then

            Dim parsedDate As DateTime


            DateTime.TryParseExact(
                normalizedDate,
                "dd/MM/yyyy",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                parsedDate
            )


            Return parsedDate.Date

        End If


        Return dateText

    End Function


    Private Sub WriteDateOfBirth(
        cell As IXLCell,
        value As String
    )

        Dim normalizedDate As String =
            NormalizeDateOrThrow(
                value
            )


        If normalizedDate = "" Then

            cell.Clear(
                XLClearOptions.Contents
            )


            cell.Style.NumberFormat.Format =
                "dd/MM/yyyy"


            Exit Sub

        End If


        Dim parsedDate As DateTime


        If Not DateTime.TryParseExact(
            normalizedDate,
            "dd/MM/yyyy",
            CultureInfo.InvariantCulture,
            DateTimeStyles.None,
            parsedDate
        ) Then

            Throw New Exception(
                "Invalid Date of Birth."
            )

        End If


        cell.Value =
            parsedDate.Date


        cell.Style.NumberFormat.Format =
            "dd/MM/yyyy"

    End Sub


    '====================================================
    ' LOG DISPLAY VALUE
    '====================================================
    Private Function DisplayLogValue(
        value As String
    ) As String

        Dim cleanValue As String =
            NormalizeText(
                value
            )


        If cleanValue = "" Then

            Return "(blank)"

        End If


        Return cleanValue

    End Function


    '====================================================
    ' ADD CHANGED FIELD
    '====================================================
    Private Sub AddChangedField(
        changes As List(Of String),
        fieldName As String,
        oldValue As String,
        newValue As String
    )

        Dim oldClean As String =
            NormalizeText(
                oldValue
            )


        Dim newClean As String =
            NormalizeText(
                newValue
            )


        If oldClean.Equals(
            newClean,
            StringComparison.OrdinalIgnoreCase
        ) Then

            Exit Sub

        End If


        changes.Add(
            fieldName &
            ": " &
            DisplayLogValue(
                oldClean
            ) &
            " → " &
            DisplayLogValue(
                newClean
            )
        )

    End Sub


    '====================================================
    ' ADD STUDENT - OLD SIGNATURE
    '====================================================
    Public Sub AddStudent(
        studentNo As String,
        studentName As String,
        gender As String,
        iquama As String,
        grade As String,
        room As String,
        phone1 As String,
        phone2 As String,
        phone3 As String,
        phone4 As String
    )

        AddStudent(
            studentNo,
            studentName,
            gender,
            iquama,
            "",
            grade,
            room,
            phone1,
            phone2,
            phone3,
            phone4
        )

    End Sub


    '====================================================
    ' ADD STUDENT - DOB
    '====================================================
    Public Sub AddStudent(
        studentNo As String,
        studentName As String,
        gender As String,
        iquama As String,
        dateOfBirth As String,
        grade As String,
        room As String,
        phone1 As String,
        phone2 As String,
        phone3 As String,
        phone4 As String
    )

        studentNo =
            NormalizeText(
                studentNo
            )


        If studentNo = "" Then

            Throw New Exception(
                "Student No cannot be blank."
            )

        End If


        Dim filePath As String =
            WorkbookPath


        Dim normalizedBirthDate As String =
            NormalizeDateOrThrow(
                dateOfBirth
            )


        Using wb As New XLWorkbook(
            filePath
        )

            Dim ws As IXLWorksheet =
                wb.Worksheet(
                    WorksheetName
                )


            Dim columns =
                GetHeaders(
                    ws
                )


            ValidateMainHeaders(
                columns
            )


            For rowNumber As Integer =
                FirstDataRow To LastDataRow

                Dim existingStudentNo As String =
                    ws.Cell(
                        rowNumber,
                        columns("Student No")
                    ).
                    GetFormattedString().
                    Trim()


                If existingStudentNo <> "" AndAlso
                   existingStudentNo.Equals(
                       studentNo,
                       StringComparison.OrdinalIgnoreCase
                   ) Then

                    Throw New Exception(
                        "Student No " &
                        studentNo &
                        " already exists."
                    )

                End If

            Next


            Dim targetRow As Integer =
                FindFirstEmptyMainRow(
                    ws,
                    columns
                )


            If targetRow = 0 Then

                Throw New Exception(
                    "There are no empty rows available in Main Sheet."
                )

            End If


            CreateSafetyBackup(
                filePath
            )


            WriteStudentToMainSheet(
                ws,
                columns,
                targetRow,
                studentNo,
                studentName,
                gender,
                iquama,
                normalizedBirthDate,
                grade,
                room,
                phone1,
                phone2,
                phone3,
                phone4,
                True
            )


            Dim addDetails As String =
                "New student added" &
                " | Grade: " &
                DisplayLogValue(
                    NormalizeUpper(
                        grade
                    )
                ) &
                " | Room: " &
                DisplayLogValue(
                    NormalizeUpper(
                        room
                    )
                )


            If normalizedBirthDate <> "" Then

                addDetails &=
                    " | Date of Birth: " &
                    normalizedBirthDate

            End If


            ActivityLogger.AddEntry(
                wb,
                "ADDED",
                studentNo,
                studentName,
                addDetails
            )


            wb.Save()

        End Using

    End Sub


    '====================================================
    ' FIND FIRST EMPTY ROW
    '====================================================
    Private Function FindFirstEmptyMainRow(
        ws As IXLWorksheet,
        columns As Dictionary(Of String, Integer)
    ) As Integer

        For rowNumber As Integer =
            FirstDataRow To LastDataRow

            Dim currentStudentNo As String =
                ws.Cell(
                    rowNumber,
                    columns("Student No")
                ).
                GetFormattedString().
                Trim()


            If currentStudentNo = "" Then

                Return rowNumber

            End If

        Next


        Return 0

    End Function


    '====================================================
    ' WRITE STUDENT
    '====================================================
    Private Sub WriteStudentToMainSheet(
        ws As IXLWorksheet,
        columns As Dictionary(Of String, Integer),
        targetRow As Integer,
        studentNo As String,
        studentName As String,
        gender As String,