Attribute VB_Name = "Module_MainProcess"
Option Explicit

' --- �萔��` ---
' �V�[�g��
Private Const SHEET_NAME_TABLE As String = "table"

' �o�ȕ�V�[�g�̃Z���͈�
Private Const ROW_DATE As Long = 3
Private Const ROW_WEEKDAY As Long = 4
Private Const COL_START_DATE As Long = 7 ' G�� (���t�̊J�n��)

Private Const ROW_STUDENT_START As Long = 6
Private Const ROW_STUDENT_END As Long = 50
Private Const COL_STUDENT_NUMBER As Long = 1 ' A��
Private Const COL_STUDENT_NAME As Long = 2 ' B�� (��ԍ��Ŏw��)

' �W�v��̒�` (AL�񂩂�AR��)
Private Const COL_SUM_SHUTTEI As Long = 38 ' AL�� (�o��)
Private Const COL_SUM_KIBIKI As Long = 39 ' AM�� (������)
Private Const COL_SUM_KESSEKI As Long = 40 ' AN�� (����)
Private Const COL_SUM_KOUKETSU As Long = 41 ' AO�� (����)
Private Const COL_SUM_ATTENDANCE As Long = 42 ' AP�� (�o�ȓ���)
Private Const COL_SUM_CHIKOKU As Long = 43 ' AQ�� (�x��)
Private Const COL_SUM_SOUTAI As Long = 44 ' AR�� (����)


Private Const CELL_WAREKI As String = "A2"
Private Const CELL_MONTH As String = "E2"
Private Const CELL_GRADE As String = "S2"
Private Const CELL_CLASS As String = "W2"
Private Const CELL_TEACHER As String = "AD2"
Private Const CELL_SCHOOL_DAY_COUNT As String = "B52" ' ���Ɠ������i�[����Z��

' ���k����CSV�֘A
Private Const CSV_FILE_NAME As String = "�������k�o�^�p����.csv"
Private Const CSV_COL_CLASS As Long = 4  ' D��
Private Const CSV_COL_STUDENT_NUMBER As Long = 5 ' E��
Private Const CSV_COL_STUDENT_NAME As Long = 6 ' F��
Private Const CSV_COL_STATUS As Long = 12 ' L�� (1:�L��, 0:���p��~)

' ���ȘA��CSV�֘A
Private Const ABSENCE_FILE_NAME As String = "���ȘA��.csv"
Private Const ABSENCE_COL_DATE As Long = 1    ' A��
Private Const ABSENCE_COL_GRADE As Long = 2  ' B��
Private Const ABSENCE_COL_CLASS As Long = 3  ' C��
Private Const ABSENCE_COL_STUDENT_NUMBER As Long = 4 ' D��
Private Const ABSENCE_COL_CONTACT_TYPE As Long = 7 ' G��
Private Const ABSENCE_COL_REASON As Long = 10 ' J��

' ���̑�
Private Const MAX_STUDENTS_ON_SHEET As Long = 50 ' �o�ȕ�V�[�g�ɕ\������ő吶�k��

' --- ��ȃv���V�[�W�� ---

Public Sub �o�ȕ�ЂȌ`�쐬()
    Dim ws As Worksheet
    Dim teacherName As String
    Dim grade As String
    Dim classNum As String
    Dim currentYear As Integer
    Dim currentMonth As Integer
    Dim lastDay As Integer
    Dim forceSchoolDays() As Integer
    Dim forceHolidays() As Integer
    Dim vacationStartDate As Date ' �����x�ɊJ�n��
    Dim vacationEndDate As Date   ' �����x�ɏI����
    Dim studentCount As Integer

    ' 1. �V�[�g�̏����ƃN���A
    Set ws = PrepareAttendanceSheet()
    If ws Is Nothing Then Exit Sub ' �V�[�g�������s���͏I��

    ' 2. ���[�U�[���͂̎擾
    ' �����x�ɂ̊J�n���ƏI�������擾���������ǉ�
    If Not GetUserInputs(teacherName, grade, classNum, currentYear, currentMonth, _
                         forceSchoolDays, forceHolidays, vacationStartDate, vacationEndDate) Then
        MsgBox "���[�U�[���͂��L�����Z�����ꂽ���߁A�����𒆒f���܂��B", vbInformation, "�������f"
        Exit Sub ' ���̓L�����Z�����͏I��
    End If

    ' 3. ��{���̏������݁i�a��A���A�w�N�A�N���X�A�S�C���j
    Call WriteBasicInfo(ws, teacherName, grade, classNum, currentYear, currentMonth)

    ' 4. ���t�Ɨj���̐���
    lastDay = SetDatesAndWeekdays(ws, currentYear, currentMonth)

    ' 5. ���k����̓ǂݍ��݂Ɠ]�L
    studentCount = LoadAndWriteStudentList(ws, grade, classNum)
    Dim lastStudentDataRow As Long
    If studentCount > 0 Then
        lastStudentDataRow = ROW_STUDENT_START + studentCount - 1
    Else
        ' ���k�����Ȃ��ꍇ�ł��A���t�̌r���`��̂��߂ɍŒ���̍s�͈͂�ݒ�
        lastStudentDataRow = ROW_STUDENT_END
    End If

    ' 6. �x������A�r���Əj�����̕`��A���Ɠ����̃J�E���g
    Dim schoolDayCount As Long
    ' �� �C��: vacationStartDate �� vacationEndDate �������ɒǉ�
    schoolDayCount = DrawHolidayLinesAndLabels(ws, currentYear, currentMonth, _
                                             lastDay, lastStudentDataRow, _
                                             forceSchoolDays, forceHolidays, _
                                             vacationStartDate, vacationEndDate)
    ws.Range(CELL_SCHOOL_DAY_COUNT).Value = schoolDayCount ' ���Ɠ������Z���ɏ�������

    ' 7. �����x�ɂ̎΂ߐ��`��
    If vacationStartDate <> 0 And vacationEndDate <> 0 Then ' ���t���L���ȏꍇ�̂ݕ`��
        Call DrawSingleVacationLine(ws, currentYear, currentMonth, studentCount, vacationStartDate, vacationEndDate)
    End If

    ' 8. �W�v��̃N���A�Ə��� (���ȏ�񔽉f�O�ɃN���A���Ă���)
    Call ClearAndPrepareSummaryColumns(ws)

    ' 9. ���ȏ��̔��f�m�F
    If MsgBox("���ȘA������̏��𔽉f���܂����H", vbYesNo + vbQuestion, "���ȏ�񔽉f") = vbYes Then
        Call ���ȏ�񔽉f(grade, classNum, currentYear, currentMonth) ' �����v���V�[�W�����Ăяo��
    End If
    
    ' 10. �W�v�����s (���ȏ�񔽉f��Ɏ��s����)
    Call CalculateAttendanceSummary(ws, lastDay, studentCount, schoolDayCount)

    MsgBox "�o�ȕ�̂ЂȌ`���쐬����܂����B" & vbCrLf & _
           "���k���� " & studentCount & " ���o�^���܂����B" & vbCrLf & _
           "���Ɠ���: " & schoolDayCount & " ��", vbInformation, "����"
End Sub

' --- �v���C�x�[�g�֐��ƃv���V�[�W�� ---

' �� 1. �V�[�g�̏����ƃN���A���s���֐�
Private Function PrepareAttendanceSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME_TABLE)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = SHEET_NAME_TABLE
    End If

    ' �����̐}�`�����ׂč폜 (���A�e�L�X�g�{�b�N�X�Ȃ�)
    Dim i As Integer
    For i = ws.Shapes.count To 1 Step -1
        ws.Shapes(i).Delete
    Next i

    ' �N���A����͈͂��w��
    Dim clearRangeData As Range
    ' ���̌��̍ő�����i31���j���̗�܂ł��N���A�ΏۂƂ���
    Set clearRangeData = ws.Range(ws.Cells(ROW_STUDENT_START, COL_START_DATE), ws.Cells(ROW_STUDENT_END, COL_START_DATE + 31 - 1))
    
    Dim clearRangeSummary As Range
    Set clearRangeSummary = ws.Range(ws.Cells(ROW_STUDENT_START, COL_SUM_SHUTTEI), ws.Cells(ROW_STUDENT_END, COL_SUM_SOUTAI))

    ' �a��A�����A�w�N�A�N���X�A�S�C���A���Ɠ����Z�����N���A
    ws.Range(CELL_WAREKI).ClearContents
    ws.Range(CELL_MONTH).ClearContents
    ws.Range(CELL_GRADE).ClearContents
    ws.Range(CELL_CLASS).ClearContents
    ws.Range(CELL_TEACHER).ClearContents
    ws.Range(CELL_SCHOOL_DAY_COUNT).ClearContents

    ' �w���ԍ��Ǝ�������N���A
    ws.Range(ws.Cells(ROW_STUDENT_START, COL_STUDENT_NUMBER), ws.Cells(ROW_STUDENT_END, COL_STUDENT_NAME)).ClearContents

    ' ���t�s�Ɨj���s�̃R���e���c���N���A�i�t�H���g�J���[���܂ށj
    With ws.Range("G3:AK4")
        .ClearContents
        .Font.Color = RGB(0, 0, 0) ' ���F�Ƀ��Z�b�g
    End With
    
    ' �w��͈͂̃R���e���c���N���A���A�t�H���g�F�����Z�b�g
    With clearRangeData
        .ClearContents
        .Font.Color = RGB(0, 0, 0) ' ���F�Ƀ��Z�b�g
    End With
    
    With clearRangeSummary
        .ClearContents
        .Font.Color = RGB(0, 0, 0) ' ���F�Ƀ��Z�b�g
    End With

    ' �R�����g�i�����j���폜 (�f�[�^�͈͂ƏW�v�͈͂ɑ΂��Ă��ꂼ����s)
    Dim cell As Range
    On Error Resume Next
    For Each cell In clearRangeData
        If Not cell.Comment Is Nothing Then
            cell.Comment.Delete
        End If
    Next cell
    For Each cell In clearRangeSummary
        If Not cell.Comment Is Nothing Then
            cell.Comment.Delete
        End If
        Next cell
    On Error GoTo 0

    Set PrepareAttendanceSheet = ws
End Function

' �� 2. ���[�U�[���͂̎擾���s���֐� (�����x�ɂ̈�����ǉ�)
Private Function GetUserInputs(ByRef teacherName As String, ByRef grade As String, ByRef classNum As String, _
                               ByRef currentYear As Integer, ByRef currentMonth As Integer, _
                               ByRef forceSchoolDays() As Integer, ByRef forceHolidays() As Integer, _
                               ByRef vacationStartDate As Date, ByRef vacationEndDate As Date) As Boolean

    Dim inputString As String
    Dim dateInput As String
    
    ' �����l��ݒ�
    vacationStartDate = 0
    vacationEndDate = 0

    ' �S�C��
    inputString = SafeInputBox("�S�C�̖��O����͂��Ă��������B�����ŃL�����Z������ƒ��[���N���A���ďI�����܂��B", "�S�C���")
    If inputString = "" Then GetUserInputs = False: Exit Function
    teacherName = inputString

    ' �w�N
    inputString = SafeInputBox("�w�N����͂��Ă��������B(���p����:1,2�Ȃ�)", "�w�N���")
    If inputString = "" Then GetUserInputs = False: Exit Function
    grade = inputString

    ' �N���X
    inputString = SafeInputBox("�N���X����͂��Ă��������B(���p�p����:1,2,B�Ȃ�)", "�N���X���")
    If inputString = "" Then GetUserInputs = False: Exit Function
    classNum = inputString

    ' �N�x
    currentYear = year(Date)
    inputString = SafeInputBox("�N�x����͂��Ă��������B" & vbCrLf & _
                                 "���݂̔N�x�� " & currentYear & " �N�ł��B" & vbCrLf & _
                                 "�ύX���K�v�ȏꍇ�͐V�����l����͂��Ă��������B", _
                                 "�N�x�m�F", CStr(currentYear))
    If IsNumeric(inputString) Then
        If CInt(inputString) >= 2000 And CInt(inputString) <= 2100 Then
            currentYear = CInt(inputString)
        Else
            MsgBox "�L���ȔN�x�i2000�`2100�j����͂��Ă��������B����l�� " & currentYear & " �N���g�p���܂��B", vbExclamation
        End If
    ElseIf inputString <> "" Then
        MsgBox "���l�ȊO�����͂���܂����B����l�� " & currentYear & " �N���g�p���܂��B", vbExclamation
    End If

    ' ��
    currentMonth = month(Date)
    inputString = SafeInputBox("�m�F: ���݂̌��� " & currentMonth & " ���ł��B" & vbCrLf & _
                                 "�ύX���K�v�ȏꍇ�͐V�����l����͂��Ă��������B", _
                                 "���̊m�F", CStr(currentMonth))
    If IsNumeric(inputString) Then
        If CInt(inputString) >= 1 And CInt(inputString) <= 12 Then
            currentMonth = CInt(inputString)
        Else
            MsgBox "�L���Ȍ��i1�`12�j����͂��Ă��������B����l�� " & currentMonth & " �����g�p���܂��B", vbExclamation
        End If
    ElseIf inputString <> "" Then
        MsgBox "���l�ȊO�����͂���܂����B����l�� " & currentMonth + " �����g�p���܂��B", vbExclamation
    End If

    ' �������Ɠ�
    Dim forceDatesInput As String
    If MsgBox("�j����y�j���ł����Ɠ��Ƃ��Ĉ������t�͂���܂����H�i�^����A���ƎQ�ςȂǁj", vbYesNo + vbQuestion, "���Ɠ��Ƃ��Ĉ������t") = vbYes Then
        forceDatesInput = SafeInputBox("���Ɠ��Ƃ��Ĉ������t���J���}��؂�œ��͂��Ă��������B(���p����: 1,8,22�Ȃ�)", "�������Ɠ�")
        forceSchoolDays = ParseCommaSeparatedNumbers(forceDatesInput)
    Else
        ReDim forceSchoolDays(0) ' ��̔z���������
    End If

    ' �����x��
    If MsgBox("�����ł��x���Ƃ��Ĉ������t�͂���܂����H�i�x�Z�Ȃǁj", vbYesNo + vbQuestion, "�x���Ƃ��Ĉ������t") = vbYes Then
        forceDatesInput = SafeInputBox("�x���Ƃ��Ĉ������t���J���}��؂�œ��͂��Ă��������B(���p����: 10,12�Ȃ�)", "�����x��")
        forceHolidays = ParseCommaSeparatedNumbers(forceDatesInput)
    Else
        ReDim forceHolidays(0) ' ��̔z���������
    End If
    
    ' �������x�ɂ̓��̓v�����v�g��ǉ���
    If MsgBox("�����x�ɂ͂���܂����H�i�ċG�A�~�G�Ȃǁj" & vbCrLf & _
              "�����x�Ɋ��Ԃ̊Y���Z���Ɏ΂ߐ���������܂��B", vbYesNo + vbQuestion, "�����x�ɐݒ�") = vbYes Then
        dateInput = SafeInputBox("�����x�ɂ̊J�n������͂��Ă��������B(��: " & Format(Date, "yyyy/m/d") & ")", "�����x�ɊJ�n��")
        If IsDate(dateInput) Then
            vacationStartDate = CDate(dateInput)
        ElseIf dateInput <> "" Then
            MsgBox "�L���ȓ��t�`���ł͂���܂���B�����x�ɂ͐ݒ肳��܂���B", vbExclamation
        End If
        
        If vacationStartDate <> 0 Then ' �J�n�����L���ȏꍇ�̂ݏI������q�˂�
            dateInput = SafeInputBox("�����x�ɂ̏I��������͂��Ă��������B(��: " & Format(Date, "yyyy/m/d") & ")", "�����x�ɏI����")
            If IsDate(dateInput) Then
                vacationEndDate = CDate(dateInput)
                If vacationEndDate < vacationStartDate Then
                    MsgBox "�I�������J�n�����O�̓��t�ł��B�����x�ɂ͐ݒ肳��܂���B", vbExclamation
                    vacationStartDate = 0 ' �����ȓ��t�Ƃ��ă��Z�b�g
                    vacationEndDate = 0
                End If
            ElseIf dateInput <> "" Then
                MsgBox "�L���ȓ��t�`���ł͂���܂���B�����x�ɂ͐ݒ肳��܂���B", vbExclamation
                vacationStartDate = 0 ' �����ȓ��t�Ƃ��ă��Z�b�g
                vacationEndDate = 0
            End If
        End If
    End If

    GetUserInputs = True
End Function

' �� ���[�U�[���͕⏕�֐��i�L�����Z�����̋󕶎����Ԃ��j
Private Function SafeInputBox(prompt As String, title As String, Optional defaultValue As Variant) As String
    On Error Resume Next
    If IsMissing(defaultValue) Then
        SafeInputBox = InputBox(prompt, title)
    Else
        SafeInputBox = InputBox(prompt, title, defaultValue)
    End If
    On Error GoTo 0
End Function

' �� �J���}��؂�̐���������𐮐��z��ɕϊ�����֐�
Private Function ParseCommaSeparatedNumbers(ByVal inputString As String) As Integer()
    Dim tempArray() As String
    Dim resultArray() As Integer
    Dim count As Long
    Dim i As Long

    If Trim(inputString) = "" Then
        ReDim resultArray(0) ' ��̔z��
        ParseCommaSeparatedNumbers = resultArray
        Exit Function
    End If

    tempArray = Split(inputString, ",")
    count = 0
    For i = LBound(tempArray) To UBound(tempArray)
        If IsNumeric(Trim(tempArray(i))) Then
            count = count + 1
        End If
    Next i

    If count > 0 Then
        ReDim resultArray(count - 1)
        count = 0
        For i = LBound(tempArray) To UBound(tempArray)
            If IsNumeric(Trim(tempArray(i))) Then
                resultArray(count) = CInt(Trim(tempArray(i)))
                count = count + 1
            End If
        Next i
    Else
        ReDim resultArray(0) ' �L���Ȑ������Ȃ��ꍇ
    End If

    ParseCommaSeparatedNumbers = resultArray
End Function

' �� 3. ��{�����V�[�g�ɏ������ރv���V�[�W��
Private Sub WriteBasicInfo(ws As Worksheet, teacherName As String, grade As String, classNum As String, _
                           currentYear As Integer, currentMonth As Integer)
    ws.Range(CELL_GRADE).Value = grade
    ws.Range(CELL_CLASS).Value = classNum
    ws.Range(CELL_TEACHER).Value = teacherName

    Dim wareki As String
    ' �ߘa�̊J�n�N��2019�N�Ƃ��Čv�Z (2019�N = �ߘa1�N)
    wareki = "�ߘa" & (currentYear - 2018) & "�N"
    ws.Range(CELL_WAREKI).Value = wareki

    ws.Range(CELL_MONTH).Value = currentMonth
End Sub

' �� 4. ���t�Ɨj�����V�[�g�ɐݒ肷��֐�
Private Function SetDatesAndWeekdays(ws As Worksheet, currentYear As Integer, currentMonth As Integer) As Integer
    Dim lastDay As Integer
    lastDay = day(DateSerial(currentYear, currentMonth + 1, 0)) ' ���̌��̍ŏI�����v�Z

    Dim i As Integer
    Dim currentDate As Date

    ' ���t�����
    For i = 1 To lastDay
        ws.Cells(ROW_DATE, COL_START_DATE + i - 1).Value = i
    Next i

    ' �j�������
    For i = 1 To lastDay
        currentDate = DateSerial(currentYear, currentMonth, i)
        ws.Cells(ROW_WEEKDAY, COL_START_DATE + i - 1).Value = WeekdayName(Weekday(currentDate), True)
    Next i

    SetDatesAndWeekdays = lastDay
End Function

' �� 5. ���k�����ǂݍ��݁A�V�[�g�ɓ]�L����֐�
Private Function LoadAndWriteStudentList(ws As Worksheet, grade As String, classNum As String) As Integer
    Dim csvFile As String
    Dim csvWB As Workbook
    Dim csvWS As Worksheet
    Dim lastRow As Long
    Dim targetClass As String
    Dim studentCounter As Integer
    Dim i As Long

    ' ��: "1�NA�g" �̌`���ɍ��킹��
    targetClass = grade & "�N" & classNum & "�g"
    csvFile = ThisWorkbook.Path & "\" & CSV_FILE_NAME
    studentCounter = 0

    If Dir(csvFile) = "" Then
        MsgBox "���k����CSV�t�@�C����������܂���: " & csvFile, vbExclamation, "�t�@�C�������o"
        LoadAndWriteStudentList = 0
        Exit Function
    End If

    On Error Resume Next
    Set csvWB = Workbooks.Open(csvFile)
    If Err.Number <> 0 Then
        MsgBox "CSV�t�@�C�����J���܂���ł���: " & Err.Description, vbCritical, "�G���["
        LoadAndWriteStudentList = 0
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    Set csvWS = csvWB.Worksheets(1)
    lastRow = csvWS.Cells(csvWS.Rows.count, "A").End(xlUp).row ' CSV�t�@�C���̍ŏI�s���擾

    For i = 2 To lastRow ' �w�b�_�[�s���X�L�b�v
        ' �N���X����v���A���X�e�[�^�X���u�L��(1)�v�̐��k�݂̂𒊏o
        If csvWS.Cells(i, CSV_COL_CLASS).Value = targetClass And csvWS.Cells(i, CSV_COL_STATUS).Value = 1 Then
            studentCounter = studentCounter + 1
            If studentCounter <= MAX_STUDENTS_ON_SHEET Then ' �V�[�g�ɕ\������ő吶�k���𒴂��Ȃ��ꍇ
                ws.Cells(ROW_STUDENT_START + studentCounter - 1, COL_STUDENT_NUMBER).Value = csvWS.Cells(i, CSV_COL_STUDENT_NUMBER).Value
                ws.Cells(ROW_STUDENT_START + studentCounter - 1, COL_STUDENT_NAME).Value = csvWS.Cells(i, CSV_COL_STUDENT_NAME).Value
            End If
        End If
    Next i

    csvWB.Close False ' CSV�t�@�C����ۑ������ɕ���
    LoadAndWriteStudentList = studentCounter ' �o�^�������k����Ԃ�
End Function

' �� 6. �x������A�r���Əj�����̕`��A���Ɠ����̃J�E���g���s���֐�
Private Function DrawHolidayLinesAndLabels(ws As Worksheet, currentYear As Integer, currentMonth As Integer, _
                                            lastDay As Integer, lastStudentDataRow As Long, _
                                            forceSchoolDays() As Integer, forceHolidays() As Integer, _
                                            vacationStartDate As Date, vacationEndDate As Date) As Long ' �� �ǉ�����
    Dim schoolDayCount As Long
    schoolDayCount = 0

    Dim i As Integer
    Dim currentDate As Date
    Dim isHoliday As Boolean
    Dim holidayName As String
    Dim isForceHoliday As Boolean
    Dim isForceSchoolDay As Boolean
    Dim j As Integer
    Dim lineShape As Shape
    Dim leftPosition As Double
    Dim topPosition As Double
    Dim heightOfLine As Double
    Dim txtBox As Shape
    Dim nationalHolidayName As String

    For i = 1 To lastDay
        currentDate = DateSerial(currentYear, currentMonth, i)
        isHoliday = False
        holidayName = ""

        ' �� ��������ǉ�: �����x�Ɋ��ԓ��̃`�F�b�N ��
        If vacationStartDate <> 0 And vacationEndDate <> 0 Then ' �����x�ɂ��ݒ肳��Ă���ꍇ�̂�
            If currentDate >= vacationStartDate And currentDate <= vacationEndDate Then
                ' �����x�Ɋ��ԓ��̓��t�́A�ʂ̏c���Əj������`�悵�Ȃ�
                ' �������A���t�Ɨj���̐Ԏ��\���͒����x�ɂ̏����ōs�����߁A�����ł͉������Ȃ�
                ' ���Ɠ������J�E���g���Ȃ��i�����x�ɒ��͎��Ɠ��ł͂Ȃ����߁j
                GoTo NextDayLoop ' ���̓��̏����փX�L�b�v
            End If
        End If
        ' �� �����܂Œǉ� ��

        ' �����x�����X�g�Ɋ܂܂�邩�`�F�b�N�i�ŗD��j
        isForceHoliday = False
        For j = LBound(forceHolidays) To UBound(forceHolidays)
            If forceHolidays(j) = i Then
                isForceHoliday = True
                Exit For
            End If
        Next j
        If isForceHoliday Then
            isHoliday = True
            holidayName = "�x�Z" ' �����x���̏ꍇ�A�x�������u�x�Z�v�Ƃ���
        End If

        ' �������Ɠ����X�g�Ɋ܂܂�邩�`�F�b�N�i�����x�����D�悳��邽�߁AisForceHoliday��False�̏ꍇ�̂ݓK�p�j
        isForceSchoolDay = False
        For j = LBound(forceSchoolDays) To UBound(forceSchoolDays)
            If forceSchoolDays(j) = i Then
                isForceSchoolDay = True
                Exit For
            End If
        Next j
        If isForceSchoolDay And Not isForceHoliday Then ' �����x���łȂ���΋������Ɠ���K�p
            isHoliday = False
            holidayName = "" ' �������Ɠ��̏ꍇ�A�x�����Ȃ�
        End If

        ' �����ݒ肪�K�p����Ă��Ȃ��ꍇ�A�ʏ�̋x��������s��
        If Not isForceHoliday And Not isForceSchoolDay Then
            ' ���j���̏ꍇ
            If Weekday(currentDate) = vbSunday Then
                isHoliday = True
            End If

            ' �y�j���̏ꍇ
            If Weekday(currentDate) = vbSaturday Then
                isHoliday = True
            End If

            ' �j���̏ꍇ
            nationalHolidayName = GetNationalHolidayName(currentDate)
            If nationalHolidayName <> "" Then
                isHoliday = True
                holidayName = nationalHolidayName
            End If
        End If

        If isHoliday Then
            ' ���t�Z����Ԏ��� (�����x�Ɋ��ԊO�̋x���̂�)
            ws.Cells(ROW_DATE, COL_START_DATE + i - 1).Font.Color = RGB(255, 0, 0)
            ' �j���Z����Ԏ��� (�����x�Ɋ��ԊO�̋x���̂�)
            ws.Cells(ROW_WEEKDAY, COL_START_DATE + i - 1).Font.Color = RGB(255, 0, 0)

            ' �c��������
            leftPosition = ws.Cells(ROW_DATE, COL_START_DATE + i - 1).Left + ws.Cells(ROW_DATE, COL_START_DATE + i - 1).Width / 2
            topPosition = ws.Rows(ROW_STUDENT_START).Top
            ' ���k�f�[�^�̍ŏI�s�܂Ő���`��
            heightOfLine = (ws.Cells(lastStudentDataRow, COL_STUDENT_NAME).Top + ws.Cells(lastStudentDataRow, COL_STUDENT_NAME).Height) - topPosition

            Set lineShape = ws.Shapes.AddLine( _
                leftPosition, topPosition, _
                leftPosition, topPosition + heightOfLine)

            With lineShape.Line
                .Weight = 1.5
                .DashStyle = msoLineSolid
                .ForeColor.RGB = RGB(255, 0, 0)
            End With

            ' �j�������ݒ肳��Ă���ꍇ�̂݃e�L�X�g�{�b�N�X��ǉ�
            If holidayName <> "" Then
                On Error Resume Next
                Set txtBox = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                                                    leftPosition - 5, topPosition + 10, _
                                                    20, heightOfLine - 20)
                On Error GoTo 0

                If Not txtBox Is Nothing Then
                    With txtBox
                        On Error Resume Next
                        .TextFrame.Characters.Text = holidayName
                        .TextFrame.Characters.Font.Size = 8
                        .TextFrame.Characters.Font.Color.RGB = RGB(255, 0, 0)
                        .TextFrame.Characters.Font.Bold = True
                        .Fill.Visible = msoFalse
                        .Line.Visible = msoFalse
                        .TextFrame.AutoSize = True
                        .TextFrame.VerticalAlignment = xlVAlignCenter
                        .TextFrame.HorizontalAlignment = xlHAlignCenter
                        .TextFrame.Orientation = msoTextOrientationUpward
                        On Error GoTo 0
                        .Left = leftPosition - (.Width / 2)
                        .Top = topPosition + (heightOfLine / 2) - (.Height / 2)
                    End With
                End If
            End If
        Else
            ' �x���ł͂Ȃ��ꍇ�A���Ɠ������J�E���g
            schoolDayCount = schoolDayCount + 1
            ' ���t�Z���Ɨj���Z���������Ƀ��Z�b�g (�����x�Ɋ��ԊO�̕����̂�)
            ' �����x�Ɋ��ԓ��̕�����DrawSingleVacationLine�ŐԎ��ɂ��Ă���̂ŁA�����ł͏������Ȃ�
            ' �� �C��: ���ɐԎ��ɂȂ��Ă��Ȃ����m�F���Ă��獕���Ƀ��Z�b�g ��
            If ws.Cells(ROW_DATE, COL_START_DATE + i - 1).Font.Color <> RGB(255, 0, 0) Then
                ws.Cells(ROW_DATE, COL_START_DATE + i - 1).Font.Color = RGB(0, 0, 0)
                ws.Cells(ROW_WEEKDAY, COL_START_DATE + i - 1).Font.Color = RGB(0, 0, 0)
            End If
        End If
NextDayLoop: ' GoTo �X�e�[�g�����g�̃^�[�Q�b�g
    Next i

    DrawHolidayLinesAndLabels = schoolDayCount
End Function

' �� �j�����𔻒肷��֐�
Private Function GetNationalHolidayName(targetDate As Date) As String
    Dim targetYear As Integer
    Dim targetMonth As Integer
    Dim targetDay As Integer
    Dim holidayName As String

    targetYear = year(targetDate)
    targetMonth = month(targetDate)
    targetDay = day(targetDate)

    holidayName = ""

    Select Case targetMonth
        Case 1
            If targetDay = 1 Then holidayName = "����"
            If GetNthMonday(targetYear, 1, 2) = targetDay Then holidayName = "���l�̓�"

        Case 2
            If targetDay = 11 Then holidayName = "�����L�O�̓�"
            If targetDay = 23 Then holidayName = "�V�c�a����"

        Case 3
            If targetDay = GetShunbun(targetYear) Then holidayName = "�t���̓�"

        Case 4
            If targetDay = 29 Then holidayName = "���a�̓�"

        Case 5
            If targetDay = 3 Then holidayName = "���@�L�O��"
            If targetDay = 4 Then holidayName = "�݂ǂ�̓�"
            If targetDay = 5 Then holidayName = "���ǂ��̓�"
            ' �U�֋x��
            If targetYear >= 2007 And targetMonth = 5 And targetDay = 6 And Weekday(DateSerial(targetYear, 5, 3), vbSunday) = 1 And Weekday(DateSerial(targetYear, 5, 4), vbSunday) <> 1 And Weekday(DateSerial(targetYear, 5, 5), vbSunday) = 1 Then
                holidayName = "�U�֋x��"
            End If

        Case 7
            If targetYear = 2020 And targetDay = 23 Then
                holidayName = "�C�̓�"
            ElseIf targetYear = 2021 And targetDay = 22 Then
                holidayName = "�C�̓�"
            ElseIf GetNthMonday(targetYear, 7, 3) = targetDay Then
                holidayName = "�C�̓�"
            End If

        Case 8
            If targetYear = 2020 And targetDay = 10 Then
                holidayName = "�R�̓�"
            ElseIf targetYear = 2021 And targetDay = 8 Then
                holidayName = "�R�̓�"
            ElseIf targetDay = 11 Then
                holidayName = "�R�̓�"
            End If

        Case 9
            If GetNthMonday(targetYear, 9, 3) = targetDay Then holidayName = "�h�V�̓�"
            If targetDay = GetShubun(targetYear) Then holidayName = "�H���̓�"
            ' �����̋x�� (�h�V�̓��ƏH���̓��ɋ��܂ꂽ����)
            If targetYear >= 1986 Then
                Dim kRDate As Date
                Dim sBDate As Date
                kRDate = DateSerial(targetYear, 9, GetNthMonday(targetYear, 9, 3))
                sBDate = DateSerial(targetYear, 9, GetShubun(targetYear))
                If targetDate > kRDate And targetDate < sBDate And Weekday(targetDate, vbSunday) <> 1 And Weekday(targetDate, vbSaturday) <> 7 Then
                    holidayName = "�����̋x��"
                End If
            End If

        Case 10
            If targetYear = 2020 And targetDay = 24 Then
                holidayName = "�X�|�[�c�̓�"
            ElseIf targetYear = 2021 And targetDay = 18 Then
                holidayName = "�X�|�[�c�̓�"
            ElseIf targetYear >= 2020 And GetNthMonday(targetYear, 10, 2) = targetDay Then
                holidayName = "�X�|�[�c�̓�"
            End If

        Case 11
            If targetDay = 3 Then holidayName = "�����̓�"
            If targetDay = 23 Then holidayName = "�ΘJ���ӂ̓�"
    End Select

    ' �����ŐU�֋x�����ă`�F�b�N (�d����`������邽��)
    If holidayName = "" Then
        If IsSubstituteHoliday(targetDate) Then
            holidayName = "�U�֋x��"
        End If
    End If

    GetNationalHolidayName = holidayName
End Function

' �� ��N���j�����v�Z����֐�
Private Function GetNthMonday(targetYear As Integer, targetMonth As Integer, n As Integer) As Integer
    Dim firstDay As Date
    Dim firstWeekday As Integer ' 1=��, 2=��, ..., 7=�y (vbMonday�̏ꍇ)
    Dim firstMonday As Integer

    firstDay = DateSerial(targetYear, targetMonth, 1)
    firstWeekday = Weekday(firstDay, vbMonday) ' �T�̎n�܂�����j���ɐݒ�

    If firstWeekday = 1 Then ' 1�������j���������ꍇ
        firstMonday = 1
    Else ' 1�������j���ȊO�������ꍇ
        firstMonday = 8 - firstWeekday + 1
    End If

    GetNthMonday = firstMonday + (n - 1) * 7
End Function

' �� �t���̓����v�Z����֐�
Private Function GetShunbun(targetYear As Integer) As Integer
    ' �T�Z�� (���m�ȓV���w�I�v�Z�ł͂Ȃ����AVBA�p�r�Ƃ��Ă͏\��)
    GetShunbun = Int(20.8431 + 0.242194 * (targetYear - 1851) - Int((targetYear - 1851) / 4))
End Function

' �� �H���̓����v�Z����֐�
Private Function GetShubun(targetYear As Integer) As Integer
    ' �T�Z�� (���m�ȓV���w�I�v�Z�ł͂Ȃ����AVBA�p�r�Ƃ��Ă͏\��)
    GetShubun = Int(23.2488 + 0.242194 * (targetYear - 1851) - Int((targetYear - 1851) / 4))
End Function

' �� �U�֋x���𔻒肷��֐�
Private Function IsSubstituteHoliday(targetDate As Date) As Boolean
    Dim checkDate As Date
    Dim i As Integer

    IsSubstituteHoliday = False

    If Weekday(targetDate) <> vbSunday Then Exit Function ' ���j���łȂ���ΐU�֋x���ł͂Ȃ�

    ' �ߋ�7���Ԃ�k���āA���j���Əd�Ȃ�j�������邩�`�F�b�N
    For i = 1 To 7
        checkDate = targetDate - i
        ' �U�֋x�����l�����Ȃ������ȏj������֐����g�p
        If GetNationalHolidayNameWithoutSubstitute(checkDate) <> "" And Weekday(checkDate) = vbSunday Then
            IsSubstituteHoliday = True
            Exit Function
        End If
    Next i
End Function

' �� �U�֋x�����l�����Ȃ������ȏj������ (GetNationalHolidayName����U�֋x����������O)
Private Function GetNationalHolidayNameWithoutSubstitute(targetDate As Date) As String
    Dim targetYear As Integer
    Dim targetMonth As Integer
    Dim targetDay As Integer
    Dim holidayName As String

    targetYear = year(targetDate)
    targetMonth = month(targetDate)
    targetDay = day(targetDate)

    holidayName = ""

    ' �T�x���i�y���j�͂����ł͏j���Ƃ��Ĉ���Ȃ�
    If Weekday(targetDate) <> vbSunday And Weekday(targetDate) <> vbSaturday Then
        Select Case targetMonth
            Case 1
                If targetDay = 1 Then holidayName = "����"
                If GetNthMonday(targetYear, 1, 2) = targetDay Then holidayName = "���l�̓�"
            Case 2
                If targetDay = 11 Then holidayName = "�����L�O�̓�"
                If targetDay = 23 Then holidayName = "�V�c�a����"
            Case 3
                If targetDay = GetShunbun(targetYear) Then holidayName = "�t���̓�"
            Case 4
                If targetDay = 29 Then holidayName = "���a�̓�"
            Case 5
                If targetDay = 3 Then holidayName = "���@�L�O��"
                If targetDay = 4 Then holidayName = "�݂ǂ�̓�"
                If targetDay = 5 Then holidayName = "���ǂ��̓�"
            Case 7
                If targetYear = 2020 And targetDay = 23 Then
                    holidayName = "�C�̓�"
                ElseIf targetYear = 2021 And targetDay = 22 Then
                    holidayName = "�C�̓�"
                ElseIf GetNthMonday(targetYear, 7, 3) = targetDay Then
                    holidayName = "�C�̓�"
                End If
            Case 8
                If targetYear = 2020 And targetDay = 10 Then
                    holidayName = "�R�̓�"
                ElseIf targetYear = 2021 And targetDay = 8 Then
                    holidayName = "�R�̓�"
                ElseIf targetDay = 11 Then
                    holidayName = "�R�̓�"
                End If
            Case 9
                If GetNthMonday(targetYear, 9, 3) = targetDay Then holidayName = "�h�V�̓�"
                If targetDay = GetShubun(targetYear) Then holidayName = "�H���̓�"
                If targetYear >= 1986 Then
                    Dim kRDate As Date, sBDate As Date
                    kRDate = DateSerial(targetYear, 9, GetNthMonday(targetYear, 9, 3))
                    sBDate = DateSerial(targetYear, 9, GetShubun(targetYear))
                    If targetDate > kRDate And targetDate < sBDate And Weekday(targetDate, vbSunday) <> 1 And Weekday(targetDate, vbSaturday) <> 7 Then
                        holidayName = "�����̋x��"
                    End If
                End If
            Case 10
                If targetYear = 2020 And targetDay = 24 Then
                    holidayName = "�X�|�[�c�̓�"
                ElseIf targetYear = 2021 And targetDay = 18 Then
                    holidayName = "�X�|�[�c�̓�"
                ElseIf targetYear >= 2020 And GetNthMonday(targetYear, 10, 2) = targetDay Then
                    holidayName = "�X�|�[�c�̓�"
                End If
            Case 11
                If targetDay = 3 Then holidayName = "�����̓�"
                If targetDay = 23 Then holidayName = "�ΘJ���ӂ̓�"
        End Select
    End If
    
    GetNationalHolidayNameWithoutSubstitute = holidayName
End Function

' �� �����x�ɂ̃Z���͈͂�1�{�̎΂ߐ��������v���V�[�W��
Private Sub DrawSingleVacationLine(ws As Worksheet, currentYear As Integer, currentMonth As Integer, _
                                  studentCount As Integer, vacationStartDate As Date, vacationEndDate As Date)
    Dim startDay As Integer
    Dim endDay As Integer
    Dim firstDateOfSheet As Date
    Dim lastDateOfSheet As Date
    
    ' ���݂̃V�[�g�������N��
    firstDateOfSheet = DateSerial(currentYear, currentMonth, 1)
    lastDateOfSheet = DateSerial(currentYear, currentMonth, day(DateSerial(currentYear, currentMonth + 1, 0)))

    ' �����x�ɂ̊J�n���ƏI���������݂̃V�[�g�̌��ɂǂ̒��x�܂܂�邩�v�Z
    If vacationStartDate < firstDateOfSheet Then
        startDay = 1
    ElseIf year(vacationStartDate) = currentYear And month(vacationStartDate) = currentMonth Then
        startDay = day(vacationStartDate)
    Else
        Exit Sub ' �x�ɊJ�n�������݂̃V�[�g�̌�����̏ꍇ�A���̃V�[�g�ł͕`�悵�Ȃ�
    End If

    If vacationEndDate > lastDateOfSheet Then
        endDay = day(lastDateOfSheet)
    ElseIf year(vacationEndDate) = currentYear And month(vacationEndDate) = currentMonth Then
        endDay = day(vacationEndDate)
    Else
        Exit Sub ' �x�ɏI���������݂̃V�[�g�̌����O�̏ꍇ�A���̃V�[�g�ł͕`�悵�Ȃ�
    End If

    ' �`�悷��͈͂��L�����`�F�b�N
    If startDay > endDay Then Exit Sub
    If studentCount = 0 Then Exit Sub ' ���k�����Ȃ��ꍇ�͕`�悵�Ȃ�

    Dim topLeftCell As Range
    Dim bottomRightCell As Range
    Dim lineShape As Shape
    
    ' �`��̊J�n�Z���i����j
    Set topLeftCell = ws.Cells(ROW_STUDENT_START, COL_START_DATE + startDay - 1)
    
    ' �`��̏I���Z���i�E���j
    Set bottomRightCell = ws.Cells(ROW_STUDENT_START + studentCount - 1, COL_START_DATE + endDay - 1)
    
    ' 1�{�̎΂ߐ���`�� (���ォ��E����)
    Set lineShape = ws.Shapes.AddLine( _
        topLeftCell.Left, topLeftCell.Top, _
        bottomRightCell.Left + bottomRightCell.Width, bottomRightCell.Top + bottomRightCell.Height)
    
    With lineShape.Line
        .Weight = 1.5 ' ���̑���
        .DashStyle = msoLineSolid ' ����
        .ForeColor.RGB = RGB(255, 150, 150) ' �����D�F
    End With
    
    ' ����w�ʂɈړ����āA���ȋL���Ȃǂ�������悤�ɂ���
    lineShape.ZOrder msoSendToBack
    
    ' �� ��������ǉ��E�C��: ���Ԃ̗j����Ԏ��� ��
    Dim dayIndex As Integer
    For dayIndex = startDay To endDay
        Dim targetCol As Long
        targetCol = COL_START_DATE + dayIndex - 1
        
        ' ���t�Z����Ԏ���
        ws.Cells(ROW_DATE, targetCol).Font.Color = RGB(255, 0, 0)
        ' �j���Z����Ԏ���
        ws.Cells(ROW_WEEKDAY, targetCol).Font.Color = RGB(255, 0, 0)
    Next dayIndex
    ' �� �����܂Œǉ��E�C�� ��

    ' �����x�ɂł��邱�Ƃ������e�L�X�g�{�b�N�X��ǉ�
    Dim txtBox As Shape
    Dim middleDay As Integer
    
    ' �e�L�X�g�{�b�N�X��z�u����Z���̖ڈ� (���Ԃ̒���������)
    middleDay = Int((startDay + endDay) / 2)
    If middleDay = 0 Then middleDay = startDay ' 1�������̏ꍇ�̑΍�
    
    Dim targetCellForText As Range
    Set targetCellForText = ws.Cells(ROW_STUDENT_START + Int(studentCount / 2), COL_START_DATE + middleDay - 1)
    
    If targetCellForText Is Nothing Then
        Debug.Print "Debug: targetCellForText��Nothing�ł��B�e�L�X�g�{�b�N�X�̍쐬���X�L�b�v���܂��B"
        Exit Sub
    End If

    On Error GoTo AddTextboxErrorHandler
    
    Set txtBox = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                                        targetCellForText.Left, targetCellForText.Top, _
                                        50, 20)
    
    On Error GoTo 0

    If Not txtBox Is Nothing Then
        With txtBox
            On Error Resume Next
            .TextFrame.Characters.Text = "�����x��"
            .TextFrame.Characters.Font.Size = 10
            .TextFrame.Characters.Font.Color.RGB = RGB(100, 100, 100)
            .TextFrame.Characters.Font.Bold = True
            .Fill.Visible = msoFalse
            .Line.Visible = msoFalse
            .TextFrame.AutoSize = True
            .TextFrame.VerticalAlignment = xlVAlignCenter
            .TextFrame.HorizontalAlignment = xlHAlignCenter
            
            If .Width > 0 And .Height > 0 Then
                .Left = topLeftCell.Left + (bottomRightCell.Left + bottomRightCell.Width - topLeftCell.Left - .Width) / 2
                .Top = topLeftCell.Top + (bottomRightCell.Top + bottomRightCell.Height - topLeftCell.Top - .Height) / 2
            Else
                Debug.Print "Debug: �e�L�X�g�{�b�N�X��Width�܂���Height���s���ł��B�z�u���X�L�b�v�B"
            End If

            .ZOrder msoSendToBack
            On Error GoTo 0
        End With
    Else
        Debug.Print "Debug: txtBox��Nothing�ł��BAddTextbox�����s���܂����B"
        Debug.Print "Debug: targetCellForText.Left: " & targetCellForText.Left & ", targetCellForText.Top: " & targetCellForText.Top
    End If
    Exit Sub
    
AddTextboxErrorHandler:
    Debug.Print "Debug: AddTextbox�ŃG���[���������܂����B�G���[�R�[�h: " & Err.Number & ", " & Err.Description
    Err.Clear
End Sub

' �� �W�v����N���A����v���V�[�W��
Private Sub ClearAndPrepareSummaryColumns(ws As Worksheet)
    ' AL�񂩂�AR��܂ł̏W�v�͈͂��N���A
    ws.Range(ws.Cells(ROW_STUDENT_START, COL_SUM_SHUTTEI), ws.Cells(ROW_STUDENT_END, COL_SUM_SOUTAI)).ClearContents
End Sub


' �� ���ȏ�񔽉f�v���V�[�W��
Public Sub ���ȏ�񔽉f(Optional ByVal inputGrade As String = "", Optional ByVal inputClass As String = "", _
                         Optional ByVal inputYear As Integer = 0, Optional ByVal inputMonth As Integer = 0)
    Dim ws As Worksheet
    Dim absenceFile As String
    Dim absenceWB As Workbook
    Dim absenceWS As Worksheet
    Dim lastRow As Long
    Dim i As Integer, j As Integer
    Dim currentYear As Integer
    Dim currentMonth As Integer
    Dim targetGrade As String
    Dim targetClass As String
    Dim studentNumber As Integer
    Dim absenceDate As Date
    Dim absenceDay As Integer
    Dim attendanceSymbol As String
    Dim reasonText As String
    Dim foundStudent As Boolean
    Dim studentRow As Integer
    Dim contactType As String
    
    ' table�V�[�g��I��
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME_TABLE)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "�o�ȕ�V�[�g(" & SHEET_NAME_TABLE & ")��������܂���B��ɏo�ȕ�ЂȌ`���쐬���Ă��������B", vbExclamation
        Exit Sub
    End If
    
    ' �v���V�[�W���̈������ȗ����ꂽ�ꍇ�A�V�[�g�̊����l���g�p
    If inputGrade = "" Then
        targetGrade = Replace(ws.Range(CELL_GRADE).Value, "�w�N", "")
    Else
        targetGrade = inputGrade
    End If
    
    If inputClass = "" Then
        targetClass = ws.Range(CELL_CLASS).Value
    Else
        targetClass = inputClass
    End If
    
    If inputYear = 0 Then
        currentYear = year(Date) ' �f�t�H���g�l�Ƃ��Č��݂̔N���g�p
        If ws.Range(CELL_WAREKI).Value <> "" Then ' �V�[�g�ɔN�x��񂪂���΂�����g��
            Dim warekiStr As String
            warekiStr = ws.Range(CELL_WAREKI).Value
            If InStr(warekiStr, "�ߘa") > 0 Then
                Dim reiwaYear As String
                reiwaYear = Split(warekiStr, "�ߘa")(1)
                reiwaYear = Split(reiwaYear, "�N")(0)
                If IsNumeric(reiwaYear) Then
                    currentYear = CInt(reiwaYear) + 2018 ' �ߘa1�N = 2019�N
                End If
            End If
        End If
    Else
        currentYear = inputYear
    End If
    
    If inputMonth = 0 Then
        currentMonth = ws.Range(CELL_MONTH).Value
    Else
        currentMonth = inputMonth
    End If
    
    ' CSV���̃N���X�\�L (��: "1�NA�g") �ɍ��킹��
    Dim fullTargetClass As String
    fullTargetClass = targetGrade & "�N" & targetClass & "�g"
    
    absenceFile = ThisWorkbook.Path & "\" & ABSENCE_FILE_NAME
    
    If Dir(absenceFile) = "" Then
        MsgBox "���ȘA��CSV�t�@�C����������܂���: " & absenceFile, vbExclamation, "�t�@�C�������o"
        Exit Sub
    End If ' �� �������C�����܂����I
    
    On Error Resume Next
    Set absenceWB = Workbooks.Open(absenceFile)
    If Err.Number <> 0 Then
        MsgBox "CSV�t�@�C�����J���܂���ł���: " & Err.Description, vbCritical, "�G���["
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    Set absenceWS = absenceWB.Worksheets(1)
    lastRow = absenceWS.Cells(absenceWS.Rows.count, "A").End(xlUp).row
    
    Dim processedCount As Long
    processedCount = 0
    
    For i = 2 To lastRow ' �w�b�_�[�s���X�L�b�v
        Dim dateString As String
        Dim dateParts As Variant
        
        dateString = absenceWS.Cells(i, ABSENCE_COL_DATE).Value
        ' ��: "2023/04/01(�y)" �̌`��������t�����݂̂𒊏o
        dateParts = Split(dateString, "(")
        If UBound(dateParts) >= 0 Then
            On Error Resume Next
            absenceDate = CDate(dateParts(0)) ' ���t�������Date�^�ɕϊ�
            On Error GoTo 0
            
            If absenceDate = 0 Then ' ���t�ϊ��Ɏ��s�����ꍇ (CDate�͖����ȓ��t��0�ɕϊ�����)
                Debug.Print "���t�ϊ��G���[: " & dateString + " (�s: " & i & ")"
            Else
                ' �Ώۂ̔N�x�ƌ��ɍ��v���邩�m�F
                If year(absenceDate) = currentYear And month(absenceDate) = currentMonth Then
                    Dim csvGrade As String
                    Dim csvClass As String
                    
                    csvGrade = absenceWS.Cells(i, ABSENCE_COL_GRADE).Value
                    csvClass = absenceWS.Cells(i, ABSENCE_COL_CLASS).Value
                    
                    ' �w�N�ƃN���X����v���邩�m�F (csvGrade��"1�N��"�̂悤�Ȍ`����z��)
                    If InStr(csvGrade, targetGrade) > 0 And csvClass = fullTargetClass Then
                        studentNumber = CInt(absenceWS.Cells(i, ABSENCE_COL_STUDENT_NUMBER).Value)
                        contactType = Trim(absenceWS.Cells(i, ABSENCE_COL_CONTACT_TYPE).Value)
                        
                        ' �A����ʂɉ������L����ݒ�
                        Select Case contactType
                            Case "����"
                                attendanceSymbol = "�~"
                            Case "�x��"
                                attendanceSymbol = "�`"
                            Case "����"
                                attendanceSymbol = "�n"
                            Case "�x��/����"
                                attendanceSymbol = "�`�n"
                            Case "������"
                                attendanceSymbol = "�L"
                            Case "�o��"
                                attendanceSymbol = "�e"
                            Case "/"
                                attendanceSymbol = "/" ' �����̋L���Ƃ���
                            Case Else
                                attendanceSymbol = "�~" ' ����`�̏ꍇ�͌���
                        End Select
                        
                        reasonText = absenceWS.Cells(i, ABSENCE_COL_REASON).Value ' ���R���擾

                        foundStudent = False
                        ' �o�ȕ�V�[�g����Y�����k��T��
                        For j = ROW_STUDENT_START To ROW_STUDENT_END
                            If ws.Cells(j, COL_STUDENT_NUMBER).Value = studentNumber Then
                                studentRow = j
                                foundStudent = True
                                Exit For
                            End If
                        Next j

                        If foundStudent Then
                            absenceDay = day(absenceDate)
                            ' �o�ȕ�V�[�g�̊Y���Z���ɋL�����L��
                            ws.Cells(studentRow, COL_START_DATE + absenceDay - 1).Value = attendanceSymbol
                            
                            ' ���R������΃R�����g�Ƃ��Ēǉ�
                            If reasonText <> "" Then
                                On Error Resume Next
                                If Not ws.Cells(studentRow, COL_START_DATE + absenceDay - 1).Comment Is Nothing Then
                                    ws.Cells(studentRow, COL_START_DATE + absenceDay - 1).Comment.Delete ' �����R�����g���폜
                                End If
                                ' �V�����R�����g��ǉ����A�e�L�X�g��ݒ�
                                ws.Cells(studentRow, COL_START_DATE + absenceDay - 1).AddComment reasonText
                                ' �R�����g�̎����T�C�Y����
                                ws.Cells(studentRow, COL_START_DATE + absenceDay - 1).Comment.Shape.TextFrame.AutoSize = True
                                On Error GoTo 0
                            End If
                            
                            processedCount = processedCount + 1
                        End If
                    End If
                End If
            End If
        End If
    Next i
    
    absenceWB.Close False ' CSV�t�@�C����ۑ������ɕ���
    
    If processedCount > 0 Then
        MsgBox "���ȏ��̔��f���������܂����B" & vbCrLf & _
               processedCount & " ���̏��𔽉f���܂����B", vbInformation, "����"
    Else
        MsgBox "�Y�����錇�ȏ��͂���܂���ł����B", vbInformation, "����"
    End If
End Sub

' �� �o�ȕ�̏W�v���s���v���V�[�W��
Private Sub CalculateAttendanceSummary(ws As Worksheet, lastDayOfMonth As Integer, studentCount As Integer, schoolDayCount As Long)
    Dim row As Long
    Dim col As Long
    Dim studentDataRange As Range
    Dim cellValue As Variant
    Dim cell As Range ' �ϐ��錾��ǉ�
    
    Dim countShuttei As Long
    Dim countKibiki As Long
    Dim countKesseki As Long
    Dim countKouketsu As Long
    Dim countChikoku As Long
    Dim countSoutai As Long
    Dim countAttendance As Long ' �o�ȓ���
    
    ' �e���k�ɂ��ă��[�v
    For row = ROW_STUDENT_START To ROW_STUDENT_START + studentCount - 1
        ' �e���k�̏W�v�����Z�b�g
        countShuttei = 0
        countKibiki = 0
        countKesseki = 0
        countKouketsu = 0
        countChikoku = 0
        countSoutai = 0
        countAttendance = 0
        
        ' ���̐��k�̓��t�͈� (COL_START_DATE ���炻�̌��̍ŏI���܂�)
        Set studentDataRange = ws.Range(ws.Cells(row, COL_START_DATE), ws.Cells(row, COL_START_DATE + lastDayOfMonth - 1))
        
        ' ���t�͈͓��̊e�Z�����`�F�b�N
        For Each cell In studentDataRange
            cellValue = Trim(CStr(cell.Value)) ' �Z���̒l���擾���A������Ƃ��ăg����
            
            Select Case cellValue
                Case "�e"
                    countShuttei = countShuttei + 1
                Case "�L"
                    countKibiki = countKibiki + 1
                Case "�~"
                    countKesseki = countKesseki + 1
                Case "/"
                    countKouketsu = countKouketsu + 1
                Case "�`"
                    countChikoku = countChikoku + 1
                Case "�n"
                    countSoutai = countSoutai + 1
                Case "�`�n"
                    countChikoku = countChikoku + 1 ' �x���ɃJ�E���g
                    countSoutai = countSoutai + 1   ' ���ނɃJ�E���g
                Case "" ' �󔒂̏ꍇ�͏o�ȂƂ݂Ȃ�
                    ' ���̓��̓��t�Z�����Ԏ��i�x���j�łȂ���Ώo�ȂƃJ�E���g
                    Dim dayCol As Long
                    dayCol = cell.Column
                    
                    If ws.Cells(ROW_DATE, dayCol).Font.Color <> RGB(255, 0, 0) Then
                        countAttendance = countAttendance + 1
                    End If
            End Select
        Next cell
        
        ' �W�v���ʂ��V�[�g�ɏ�������
        ws.Cells(row, COL_SUM_SHUTTEI).Value = countShuttei
        ws.Cells(row, COL_SUM_KIBIKI).Value = countKibiki
        ws.Cells(row, COL_SUM_KESSEKI).Value = countKesseki
        ws.Cells(row, COL_SUM_KOUKETSU).Value = countKouketsu
        ws.Cells(row, COL_SUM_ATTENDANCE).Value = countAttendance
        ws.Cells(row, COL_SUM_CHIKOKU).Value = countChikoku
        ws.Cells(row, COL_SUM_SOUTAI).Value = countSoutai
    Next row
End Sub

