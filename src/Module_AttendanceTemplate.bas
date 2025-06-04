Attribute VB_Name = "Module_AttendanceTemplate"
Option Explicit

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
        lastStudentDataRow = Module_Constants.ROW_STUDENT_START + studentCount - 1 ' �萔�Q�Ƃ��C��
    Else
        ' ���k�����Ȃ��ꍇ�ł��A���t�̌r���`��̂��߂ɍŒ���̍s�͈͂�ݒ�
        lastStudentDataRow = Module_Constants.ROW_STUDENT_END ' �萔�Q�Ƃ��C��
    End If

    ' 6. �x������A�r���Əj�����̕`��A���Ɠ����̃J�E���g
    Dim schoolDayCount As Long
    ' �� �C��: vacationStartDate �� vacationEndDate �������ɒǉ�
    schoolDayCount = DrawHolidayLinesAndLabels(ws, currentYear, currentMonth, _
                                             lastDay, lastStudentDataRow, _
                                             forceSchoolDays, forceHolidays, _
                                             vacationStartDate, vacationEndDate)
    ws.Range(Module_Constants.CELL_SCHOOL_DAY_COUNT).Value = schoolDayCount ' ���Ɠ������Z���ɏ������� (�萔�Q�Ƃ��C��)

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