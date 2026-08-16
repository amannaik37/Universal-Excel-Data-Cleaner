Option Explicit

Sub Universal_Data_Cleaner()

    Dim targetWB As Workbook
    Dim rawWS As Worksheet
    Dim cleanWS As Worksheet
    Dim reportWS As Worksheet

    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long
    Dim c As Long

    Dim rawData As Variant
    Dim cleanData As Variant

    Dim header As String
    Dim txt As String
    Dim flag As String

    Dim missingCount As Long
    Dim duplicateCount As Long
    Dim flaggedRows As Long

    Dim rowDict As Object
    Dim rowKey As String

    Set targetWB = ActiveWorkbook

    If targetWB Is Nothing Then
        MsgBox "Please open an Excel data file first.", vbExclamation
        Exit Sub
    End If

    If UCase(targetWB.Name) = "PERSONAL.XLSB" Then
        MsgBox "Please select your data workbook first.", vbExclamation
        Exit Sub
    End If

    Set rawWS = ActiveSheet

    If Application.WorksheetFunction.CountA(rawWS.Cells) = 0 Then
        MsgBox "The active sheet is empty.", vbExclamation
        Exit Sub
    End If

    lastRow = rawWS.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = rawWS.Cells.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    If lastRow < 2 Then
        MsgBox "At least one header row and one data row are required.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    rawData = rawWS.Range(rawWS.Cells(1, 1), rawWS.Cells(lastRow, lastCol)).Value2
    ReDim cleanData(1 To lastRow, 1 To lastCol + 1)

    For c = 1 To lastCol
        header = CleanText(CStr(rawData(1, c)))
        If header = "" Then header = "Column_" & c
        cleanData(1, c) = header
    Next c

    cleanData(1, lastCol + 1) = "Data_Quality_Flag"

    For r = 2 To lastRow
        flag = ""

        For c = 1 To lastCol
            header = LCase(CStr(cleanData(1, c)))

            If IsError(rawData(r, c)) Then
                cleanData(r, c) = ""
                AddFlag flag, "Excel Error: " & cleanData(1, c)

            Else
                txt = CleanText(CStr(rawData(r, c)))

                If IsMissingValue(txt) Then
                    cleanData(r, c) = ""
                    missingCount = missingCount + 1
                    AddFlag flag, "Missing: " & cleanData(1, c)

                ElseIf InStr(header, "email") > 0 Then
                    txt = LCase(txt)
                    cleanData(r, c) = txt

                    If Not IsValidEmail(txt) Then
                        AddFlag flag, "Check Email: " & cleanData(1, c)
                    End If

                ElseIf InStr(header, "phone") > 0 _
                    Or InStr(header, "mobile") > 0 _
                    Or InStr(header, "contact") > 0 Then

                    txt = DigitsOnly(txt)
                    cleanData(r, c) = txt

                    If Len(txt) < 7 Or Len(txt) > 15 Then
                        AddFlag flag, "Check Phone: " & cleanData(1, c)
                    End If

                ElseIf InStr(header, "date") > 0 _
                    Or header = "dob" _
                    Or InStr(header, "birth") > 0 Then

                    If IsDate(rawData(r, c)) Then
                        cleanData(r, c) = CDate(rawData(r, c))
                    Else
                        cleanData(r, c) = txt
                        AddFlag flag, "Check Date: " & cleanData(1, c)
                    End If

                ElseIf IsFinancialColumn(header) Then

                    If IsNumeric(CleanNumber(txt)) Then
                        cleanData(r, c) = CDbl(CleanNumber(txt))
                    Else
                        cleanData(r, c) = txt
                        AddFlag flag, "Check Number: " & cleanData(1, c)
                    End If

                Else

                    If VarType(rawData(r, c)) = vbString Then
                        cleanData(r, c) = txt
                    Else
                        cleanData(r, c) = rawData(r, c)
                    End If

                End If
            End If
        Next c

        If flag = "" Then
            cleanData(r, lastCol + 1) = "Clean"
        Else
            cleanData(r, lastCol + 1) = flag
            flaggedRows = flaggedRows + 1
        End If
    Next r

    Set rowDict = CreateObject("Scripting.Dictionary")

    For r = 2 To lastRow
        rowKey = ""

        For c = 1 To lastCol
            rowKey = rowKey & "|" & CStr(cleanData(r, c))
        Next c

        If rowDict.Exists(rowKey) Then
            duplicateCount = duplicateCount + 1

            If cleanData(r, lastCol + 1) = "Clean" Then
                cleanData(r, lastCol + 1) = "Duplicate Record"
                flaggedRows = flaggedRows + 1
            Else
                cleanData(r, lastCol + 1) = _
                    cleanData(r, lastCol + 1) & " | Duplicate Record"
            End If
        Else
            rowDict.Add rowKey, True
        End If
    Next r

    On Error Resume Next
    targetWB.Worksheets("Clean_Data").Delete
    targetWB.Worksheets("Data_Quality_Report").Delete
    On Error GoTo ErrorHandler

    Set cleanWS = targetWB.Worksheets.Add( _
        After:=targetWB.Worksheets(targetWB.Worksheets.Count))

    cleanWS.Name = "Clean_Data"

    cleanWS.Range( _
        cleanWS.Cells(1, 1), _
        cleanWS.Cells(lastRow, lastCol + 1)).Value = cleanData

    With cleanWS
        .Rows(1).Font.Bold = True
        .Rows(1).AutoFilter
        .Columns.AutoFit
    End With

    For c = 1 To lastCol

        header = LCase(CStr(cleanData(1, c)))

        If InStr(header, "phone") > 0 _
            Or InStr(header, "mobile") > 0 _
            Or InStr(header, "contact") > 0 Then

            cleanWS.Columns(c).NumberFormat = "@"

        ElseIf InStr(header, "date") > 0 _
            Or header = "dob" _
            Or InStr(header, "birth") > 0 Then

            cleanWS.Columns(c).NumberFormat = "dd-mmm-yyyy"

        ElseIf IsFinancialColumn(header) Then

            cleanWS.Columns(c).NumberFormat = "#,##0.00"

        End If
    Next c

    Set reportWS = targetWB.Worksheets.Add( _
        After:=targetWB.Worksheets(targetWB.Worksheets.Count))

    reportWS.Name = "Data_Quality_Report"

    With reportWS

        .Range("A1").Value = "UNIVERSAL DATA QUALITY REPORT"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 16

        .Range("A3").Value = "Metric"
        .Range("B3").Value = "Result"

        .Range("A4").Value = "Source Workbook"
        .Range("B4").Value = targetWB.Name

        .Range("A5").Value = "Source Sheet"
        .Range("B5").Value = rawWS.Name

        .Range("A6").Value = "Rows Processed"
        .Range("B6").Value = lastRow - 1

        .Range("A7").Value = "Columns Processed"
        .Range("B7").Value = lastCol

        .Range("A8").Value = "Missing Values"
        .Range("B8").Value = missingCount

        .Range("A9").Value = "Duplicate Records"
        .Range("B9").Value = duplicateCount

        .Range("A10").Value = "Rows Requiring Review"
        .Range("B10").Value = flaggedRows

        .Range("A11").Value = "Clean Output Rows"
        .Range("B11").Value = lastRow - 1

        .Range("A3:B3").Font.Bold = True
        .Columns("A:B").AutoFit

    End With

    cleanWS.Activate

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox _
        "Universal Cleaning Completed!" & vbCrLf & vbCrLf & _
        "Rows Processed: " & Format(lastRow - 1, "#,##0") & vbCrLf & _
        "Columns: " & lastCol & vbCrLf & _
        "Missing Values: " & Format(missingCount, "#,##0") & vbCrLf & _
        "Duplicates: " & Format(duplicateCount, "#,##0") & vbCrLf & _
        "Rows for Review: " & Format(flaggedRows, "#,##0"), _
        vbInformation, _
        "Universal Data Cleaner"

    Exit Sub

ErrorHandler:

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox _
        "Error " & Err.Number & vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "Universal Data Cleaner"

End Sub

Function CleanText(ByVal txt As String) As String

    txt = Replace(txt, Chr(160), " ")
    txt = Replace(txt, vbTab, " ")
    txt = Replace(txt, vbCr, " ")
    txt = Replace(txt, vbLf, " ")

    On Error Resume Next
    txt = Application.WorksheetFunction.Trim(txt)
    On Error GoTo 0

    CleanText = txt

End Function

Function IsMissingValue(ByVal txt As String) As Boolean

    Select Case LCase(Trim(txt))

        Case "", "n/a", "na", "null", "none", "nil", "#n/a", "-"
            IsMissingValue = True

        Case Else
            IsMissingValue = False

    End Select

End Function

Function DigitsOnly(ByVal txt As String) As String

    Dim i As Long
    Dim ch As String
    Dim result As String

    For i = 1 To Len(txt)

        ch = Mid(txt, i, 1)

        If ch >= "0" And ch <= "9" Then
            result = result & ch
        End If

    Next i

    DigitsOnly = result

End Function

Function CleanNumber(ByVal txt As String) As String

    txt = Replace(txt, ",", "")
    txt = Replace(txt, "$", "")
    txt = Replace(txt, "₹", "")
    txt = Replace(txt, "€", "")
    txt = Replace(txt, "£", "")

    CleanNumber = Trim(txt)

End Function

Function IsFinancialColumn(ByVal header As String) As Boolean

    If InStr(header, "salary") > 0 _
        Or InStr(header, "revenue") > 0 _
        Or InStr(header, "amount") > 0 _
        Or InStr(header, "price") > 0 _
        Or InStr(header, "cost") > 0 _
        Or InStr(header, "income") > 0 _
        Or InStr(header, "wage") > 0 Then

        IsFinancialColumn = True

    Else

        IsFinancialColumn = False

    End If

End Function

Function IsValidEmail(ByVal email As String) As Boolean

    Dim atPos As Long
    Dim dotPos As Long

    atPos = InStr(1, email, "@")
    dotPos = InStrRev(email, ".")

    IsValidEmail = _
        (atPos > 1 And _
         dotPos > atPos + 1 And _
         dotPos < Len(email))

End Function

Sub AddFlag(ByRef flags As String, ByVal newFlag As String)

    If flags = "" Then

        flags = newFlag

    Else

        flags = flags & " | " & newFlag

    End If

End Sub
