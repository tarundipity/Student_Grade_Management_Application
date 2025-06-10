VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStudentMarks 
   Caption         =   "Student Grades"
   ClientHeight    =   3230
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "frmStudentMarks.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStudentMarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Courses As collection
Private grades As collection
Private Students As collection
Private Sub cancelCmdButton_Click()
    Unload Me
End Sub
Private Sub classAvgCmdButton_Click()

    On Error GoTo ErrorHandler
    
    ' check if grades collection is there (not empty)
    If grades Is Nothing Then
        MsgBox "No grade data available. Please import data first.", vbExclamation
        Exit Sub
    End If

    ' calc the class average and standard deviation using weighting
    Dim gradesCollected As collection
    Dim totalGradesWeighted As Double
    Dim totalWeight As Double
    Dim gradesArray() As Double
    Dim i As Long

    totalGradesWeighted = 0
    totalWeight = 0
    i = 1

    For Each gradesCollected In grades
        If IsNumeric(gradesCollected(1)) And IsNumeric(gradesCollected(4)) And IsNumeric(gradesCollected(5)) And IsNumeric(gradesCollected(6)) And IsNumeric(gradesCollected(7)) And IsNumeric(gradesCollected(8)) And IsNumeric(gradesCollected(9)) Then
            totalGradesWeighted = totalGradesWeighted + calculateWeight(gradesCollected)
            totalWeight = totalWeight + 1

            ' convert grade collection to array
            Dim totalTestsWeight As Double
            totalTestsWeight = calculateWeight(gradesCollected)
            ReDim Preserve gradesArray(1 To i)
            gradesArray(i) = totalTestsWeight
            i = i + 1
        End If
    Next gradesCollected

    Dim classAvg As Double
    Dim classStDev As Double

    If totalWeight > 0 Then
        classAvg = totalGradesWeighted / totalWeight
        classStDev = WorksheetFunction.StDev(gradesArray)
    Else
        classAvg = 0
        classStDev = 0
    End If

    ' show class average and standard deviation
    MsgBox "Class Average: " & Format(classAvg, "0.00") & vbCrLf & _
           "Standard Deviation: " & Format(classStDev, "0.00"), vbInformation, "Class Statistics"
CleanExit:
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    MsgBox "The " & Chr(34) & Err.Description & Chr(34) & " error occurred [" & Err.Number & "].", _
        vbCritical, "Error Handled"
    Err.Clear
    Resume CleanExit
    
End Sub
Private Function calculateWeight(gradesCollected As collection)

    Dim totalTestsWeight As Double
    totalTestsWeight = gradesCollected(4) * 0.05 + gradesCollected(5) * 0.05 + gradesCollected(6) * 0.05 + gradesCollected(7) * 0.05 + gradesCollected(8) * 0.3 + gradesCollected(9) * 0.3
    calculateWeight = totalTestsWeight
    
End Function
Private Sub continueCmdButton_Click()

    If importDataOptionButton.Value = True Then
        importFromDatabase ' call function to import data from the database
    ElseIf listCoursesOptionButton.Value = True Then
        showAvailableCourses ' call function to show message box with the available courses
    ElseIf courseEnrollOptionButton.Value = True Then
        studentsEnrolledInCourse ' call function to display students enrolled in the selected course
    ElseIf generateReportOptionButton.Value = True Then
        generateReportAndChart ' call function to generate the report and charts
    Else
        MsgBox "Please select a valid option.", vbExclamation
    End If
    
End Sub
Private Sub importFromDatabase()
    On Error GoTo ErrorHandler

    Dim filename As String
    Dim con As Object
    Dim rs As Object
    Dim strSql As String
    Dim fieldIndex As Integer

    ' if "Import Data" radio button is selected
    If importDataOptionButton.Value = True Then
        ' FileDialog
        Dim fd As FileDialog
        Set fd = Application.FileDialog(msoFileDialogOpen)
        fd.Title = "Select a file"
        fd.InitialFileName = ThisWorkbook.Path

        If fd.Show = -1 Then
            filename = fd.SelectedItems(1)

            ' conncect to the database
            Set con = CreateObject("ADODB.Connection")
            con.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filename & ";Persist Security Info=False;"

            ' using SQL for courses
            Set rs = CreateObject("ADODB.Recordset")
            strSql = "SELECT * FROM Courses"
            rs.Open strSql, con
            Set Courses = New collection
            If Not rs.EOF Then
                Do Until rs.EOF
                    Dim courseData As collection
                    Set courseData = New collection
                    For fieldIndex = 0 To rs.Fields.count - 1
                        courseData.Add rs.Fields(fieldIndex).Value, CStr(fieldIndex + 1)
                    Next fieldIndex
                    Courses.Add courseData
                    rs.MoveNext
                Loop
            End If
            rs.Close

            ' using SQL for grades
            strSql = "SELECT * FROM Grades"
            rs.Open strSql, con
            Set grades = New collection
            If Not rs.EOF Then
                Do Until rs.EOF
                    Dim grData As collection
                    Set grData = New collection
                    For fieldIndex = 0 To rs.Fields.count - 1
                        grData.Add rs.Fields(fieldIndex).Value, CStr(fieldIndex + 1)
                    Next fieldIndex
                    grades.Add grData
                    rs.MoveNext
                Loop
            End If
            rs.Close

            ' using SQL for students
            strSql = "SELECT * FROM Students"
            rs.Open strSql, con
            Set Students = New collection
            If Not rs.EOF Then
                Do Until rs.EOF
                    Dim studentData As collection
                    Set studentData = New collection
                    For fieldIndex = 0 To rs.Fields.count - 1
                        studentData.Add rs.Fields(fieldIndex).Value, CStr(fieldIndex + 1)
                    Next fieldIndex
                    Students.Add studentData
                    rs.MoveNext
                Loop
            End If
            rs.Close

            con.Close
            Set rs = Nothing
            Set con = Nothing

            ' success message
            MsgBox "Data imported successfully!", vbInformation
        End If
    End If
CleanExit:
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    MsgBox "The " & Chr(34) & Err.Description & Chr(34) & " error occurred [" & Err.Number & "].", _
        vbCritical, "Error Handled"
    Err.Clear
    Resume CleanExit
End Sub
Private Sub showAvailableCourses()
        ' check if courses collection is there (not empty)
    If Courses Is Nothing Then
        MsgBox "No courses are available. Please import data first.", vbExclamation
        Exit Sub
    End If

    ' available courses
    Dim courseList As String
    courseList = "Available Courses:" & vbCrLf

    ' loop through collection
    Dim coursesCollected As collection
    For Each coursesCollected In Courses
        Dim courseID As String
        Dim courseCode As String
        Dim courseName As String
        ' This is assuming the file imported has course id in column 1, course code in column 2 and course name in column 3
        courseID = coursesCollected(1)
        courseCode = coursesCollected(2)
        courseName = coursesCollected(3)
        courseList = courseList & " * " & courseID & " * " & courseCode & " * " & courseName & vbCrLf
    Next coursesCollected

    ' use MsgBox to show available courses
    MsgBox courseList, vbInformation, "Available Courses"
End Sub
Private Sub generateReportAndChart()
    On Error GoTo ErrorHandler

    If Courses Is Nothing Or grades Is Nothing Or Students Is Nothing Then
        MsgBox "Please import data first before generating the report.", vbExclamation
        Exit Sub
    End If

    ' gen report and charts here
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim chartSheet As Worksheet
    Dim wordRange As Object
    Dim grData As collection
    Dim grArray() As Double ' change data type
    Dim grSum As Double
    Dim count As Long
    Dim minGr As Double
    Dim maxGr As Double
    Dim meanGr As Double
    Dim modeGr As Variant
    Dim medGr As Double
    Dim strDevGr As Double
    Dim i As Long

    ' word application and document
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    Set wordDoc = wordApp.Documents.Add

    ' content
    Set wordRange = wordDoc.Content
    wordRange.Text = "Comprehensive Report of Students Grades" & vbCrLf & vbCrLf
    wordRange.Bold = True
    wordRange.Font.Underline = True
    wordRange.Collapse 0 ' cursor at the end of the document
    wordRange.Text = "Grade Statistics:" & vbCrLf & vbCrLf
    wordRange.Bold = False
    wordRange.Font.Underline = True
    wordRange.Collapse 0
    wordRange.Text = "These are the results of the data:" & vbCrLf & vbCrLf
    wordRange.Bold = False
    wordRange.Font.Underline = False
    wordRange.Collapse 0

    ' statistics from the imported data
    grSum = 0
    count = 0
    minGr = 999999
    maxGr = -999999
    Set grData = grades(1)

    ' statistical calculations in arrays
    ReDim grArray(1 To grData.count)
    For i = 1 To grData.count
        If IsNumeric(grData(i)) Then
            grArray(i) = grData(i)
            ' statistics update
            grSum = grSum + grArray(i)
            If grArray(i) < minGr Then
                minGr = grArray(i)
            End If
            If grArray(i) > maxGr Then
                maxGr = grArray(i)
            End If
            count = count + 1
        End If
    Next i

    ' mean (average)
    meanGr = grSum / count

    ' mode
    modeGr = GetMode(grArray)

    ' median
    ' sort array
    QuickSort grArray, 1, count
    If count Mod 2 = 0 Then
        medGr = (grArray(count \ 2) + grArray(count \ 2 + 1)) / 2
    Else
        medGr = grArray((count + 1) \ 2)
    End If

    ' standard deviation
    strDevGr = Application.WorksheetFunction.StDev(grArray)

    ' add stats to the document with blank lines between them
    wordRange.Text = "Minimum Grade: " & minGr & vbCrLf
    wordRange.Bold = False
    wordRange.Font.Underline = False
    wordRange.Collapse 0
    wordRange.Text = "Maximum Grade: " & maxGr & vbCrLf
    wordRange.Bold = False
    wordRange.Font.Underline = False
    wordRange.Collapse 0
    wordRange.Text = "Average Grade: " & meanGr & vbCrLf
    wordRange.Bold = False
    wordRange.Font.Underline = False
    wordRange.Collapse 0
    wordRange.Text = "Mode: " & modeGr & vbCrLf
    wordRange.Bold = False
    wordRange.Font.Underline = False
    wordRange.Collapse 0
    wordRange.Text = "Median: " & medGr & vbCrLf
    wordRange.Bold = False
    wordRange.Font.Underline = False
    wordRange.Collapse 0
    wordRange.Text = "Standard Deviation: " & strDevGr & vbCrLf & vbCrLf
    wordRange.Bold = False
    wordRange.Font.Underline = False
    wordRange.Collapse 0

    wordRange.Text = "Histogram with Finals Grades: " & vbCrLf & vbCrLf
    wordRange.Bold = True
    wordRange.Font.Underline = True
    wordRange.Collapse 0
    
    ' histogram chart to the Word document
    Set chartSheet = ThisWorkbook.Worksheets.Add
    chartSheet.Name = "HistogramChart1"

    ' bin size for the histogram
    Dim binSize As Double
    binSize = (maxGr - minGr) / 10

    ' histogram data in the worksheet based on student frequency within grade ranges
    Dim binStart As Double
    Dim binEnd As Double
    Dim binRange As Range
    Set binRange = chartSheet.Range("A1")
    Dim frequencyRange As Range
    Set frequencyRange = chartSheet.Range("B1")

    frequencyRange.Value = "Frequency"
    For i = 1 To 10 ' # of bins
        binStart = minGr + (i - 1) * binSize
        binEnd = binStart + binSize
        binRange.Offset(0, i - 1).Value = "Bin " & i
        binRange.Offset(1, i - 1).Value = binStart
        binRange.Offset(2, i - 1).Value = binEnd
        binRange.Offset(3, i - 1).Value = GetFrequencyInRange(grData, binStart, binEnd)
        frequencyRange.Offset(3, i - 1).FormulaR1C1 = "=SUM(R[-3]C:R[-1]C)"
    Next i

    ' histogram chart
    chartSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=chartSheet.Range("A2:K4"), PlotBy:=xlColumns
    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Text = "Grade Distribution"

    ' histogram chart to the document
    chartSheet.ChartObjects(1).Copy
    wordRange.Paste
    wordRange.Collapse 0

    ' delete chart
    Application.DisplayAlerts = False
    chartSheet.Delete
    Application.DisplayAlerts = True

    ' save
    wordDoc.SaveAs ThisWorkbook.Path & "\Comprehensive_Report.docx"

    ' success message
    MsgBox "Report generated successfully!", vbInformation
CleanExit:
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    MsgBox "The " & Chr(34) & Err.Description & Chr(34) & " error occurred [" & Err.Number & "].", _
        vbCritical, "Error Handled"
    Err.Clear
    Resume CleanExit
End Sub

Private Function GetMode(ByVal data As Variant) As Variant
    Dim dict As Object
    Dim val As Variant
    Dim count As Long
    Dim maxCount As Long
    Dim mode As Variant

    Set dict = CreateObject("Scripting.Dictionary")

    For Each val In data
        If Not IsEmpty(val) Then
            If dict.exists(val) Then
                count = dict(val) + 1
                dict(val) = count
                If count > maxCount Then
                    maxCount = count
                    mode = val
                End If
            Else
                dict.Add val, 1
                If maxCount = 0 Then
                    maxCount = 1
                    mode = val
                End If
            End If
        End If
    Next val

    If maxCount = 1 Then
        GetMode = "No mode"
    Else
        GetMode = mode
    End If
End Function
Private Sub QuickSort(arr As Variant, left As Long, right As Long)
    Dim i As Long
    Dim j As Long
    Dim pivot As Double
    Dim temp As Double

    i = left
    j = right
    pivot = arr((left + right) \ 2)

    Do While i <= j
        Do While arr(i) < pivot
            i = i + 1
        Loop
        Do While arr(j) > pivot
            j = j - 1
        Loop
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop

    If left < j Then QuickSort arr, left, j
    If i < right Then QuickSort arr, i, right
End Sub

Private Function GetFrequencyInRange(grData As collection, startValue As Double, endValue As Double) As Long
    Dim count As Long
    Dim grade As Variant

    For Each grade In grData
        If IsNumeric(grade) Then
            If grade >= startValue And grade <= endValue Then
                count = count + 1
            End If
        End If
    Next grade

    GetFrequencyInRange = count
End Function
Private Sub studentsEnrolledInCourse()
     ' Check if Courses collection is available
    If Courses Is Nothing Then
        MsgBox "No courses are available. Please import data first.", vbExclamation
        Exit Sub
    End If

    ' Ask user to input course code
    Dim courseCodeInput As String
    courseCodeInput = InputBox("Please enter the course code you want information on:", "Course Enrollment")
    
    On Error Resume Next
    Application.DisplayAlerts = False ' Disable the warning prompt
    Worksheets("CourseEnrollment").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Create a new worksheet to display information
    Dim enrollmentWorksheet As Worksheet
    Set enrollmentWorksheet = ThisWorkbook.Worksheets.Add
    enrollmentWorksheet.Name = "CourseEnrollment"

    enrollmentWorksheet.Range("A1").Value = "Course Code"
    enrollmentWorksheet.Range("B1").Value = courseCodeInput

    ' Add labels for student information
    enrollmentWorksheet.Range("A2").Value = "Student ID"
    enrollmentWorksheet.Range("B2").Value = "First Name"
    enrollmentWorksheet.Range("C2").Value = "Last Name"
    ' Loop through Grades and cross reference
    Dim grData As collection
    Dim rowIndex As Long
    rowIndex = 3
    For Each grData In grades
        If grData(3) = courseCodeInput Then ' Assuming course code is in column 3
            ' Find student data by matching Student ID
            Dim studentData As collection
            For Each studentData In Students
                If studentData(3) = grData(2) Then ' Assuming Student ID is in column 3
                    ' Add data to the worksheet
                    enrollmentWorksheet.Cells(rowIndex, 1).Value = studentData(3) ' Assuming Student ID is in column 3
                    enrollmentWorksheet.Cells(rowIndex, 2).Value = studentData(1) ' Assuming first name is in column 1
                    enrollmentWorksheet.Cells(rowIndex, 3).Value = studentData(2) ' Assuming last name is in column 2
                    rowIndex = rowIndex + 1
                    Exit For ' Exit the loop once a match is found
                End If
            Next studentData
        End If
    Next grData

    ' Autofit columns
    enrollmentWorksheet.Columns.AutoFit

    ' Successful enrollment information display message
    MsgBox "Student enrollment information for course code " & courseCodeInput & " has been displayed.", vbInformation

    ' Explicitly activate the worksheet to ensure it's visible
    enrollmentWorksheet.Activate
End Sub
