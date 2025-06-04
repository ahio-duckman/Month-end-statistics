Attribute VB_Name = "Module_Constants" 
Option Explicit

' --- �萔��` ---
' �V�[�g��
Public Const SHEET_NAME_TABLE As String = "table"

' �o�ȕ�V�[�g�̃Z���͈�
Public Const ROW_DATE As Long = 3
Public Const ROW_WEEKDAY As Long = 4
Public Const COL_START_DATE As Long = 7 ' G�� (���t�̊J�n��)

Public Const ROW_STUDENT_START As Long = 6
Public Const ROW_STUDENT_END As Long = 50
Public Const COL_STUDENT_NUMBER As Long = 1 ' A��
Public Const COL_STUDENT_NAME As Long = 2 ' B�� (��ԍ��Ŏw��)

' �W�v��̒�` (AL�񂩂�AR��)
Public Const COL_SUM_SHUTTEI As Long = 38 ' AL�� (�o��)
Public Const COL_SUM_KIBIKI As Long = 39 ' AM�� (������)
Public Const COL_SUM_KESSEKI As Long = 40 ' AN�� (����)
Public Const COL_SUM_KOUKETSU As Long = 41 ' AO�� (����)
Public Const COL_SUM_ATTENDANCE As Long = 42 ' AP�� (�o�ȓ���)
Public Const COL_SUM_CHIKOKU As Long = 43 ' AQ�� (�x��)
Public Const COL_SUM_SOUTAI As Long = 44 ' AR�� (����)


Public Const CELL_WAREKI As String = "A2"
Public Const CELL_MONTH As String = "E2"
Public Const CELL_GRADE As String = "S2"
Public Const CELL_CLASS As String = "W2"
Public Const CELL_TEACHER As String = "AD2"
Public Const CELL_SCHOOL_DAY_COUNT As String = "B52" ' ���Ɠ������i�[����Z��

' ���k����CSV�֘A
Public Const CSV_FILE_NAME As String = "�������k�o�^�p����.csv"
Public Const CSV_COL_CLASS As Long = 4   ' D��
Public Const CSV_COL_STUDENT_NUMBER As Long = 5 ' E��
Public Const CSV_COL_STUDENT_NAME As Long = 6 ' F��
Public Const CSV_COL_STATUS As Long = 12 ' L�� (1:�L��, 0:���p��~)

' ���ȘA��CSV�֘A
Public Const ABSENCE_FILE_NAME As String = "���ȘA��.csv"
Public Const ABSENCE_COL_DATE As Long = 1    ' A��
Public Const ABSENCE_COL_GRADE As Long = 2   ' B��
Public Const ABSENCE_CLASS As Long = 3   ' C��
Public Const ABSENCE_COL_STUDENT_NUMBER As Long = 4 ' D��
Public Const ABSENCE_COL_CONTACT_TYPE As Long = 7 ' G��
Public Const ABSENCE_COL_REASON As Long = 10 ' J��

' ���̑�
Public Const MAX_STUDENTS_ON_SHEET As Long = 50 ' �o�ȕ�V�[�g�ɕ\������ő吶�k��