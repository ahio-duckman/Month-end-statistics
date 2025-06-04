Attribute VB_Name = "Module_Constants" 
Option Explicit

' --- 定数定義 ---
' シート名
Public Const SHEET_NAME_TABLE As String = "table"

' 出席簿シートのセル範囲
Public Const ROW_DATE As Long = 3
Public Const ROW_WEEKDAY As Long = 4
Public Const COL_START_DATE As Long = 7 ' G列 (日付の開始列)

Public Const ROW_STUDENT_START As Long = 6
Public Const ROW_STUDENT_END As Long = 50
Public Const COL_STUDENT_NUMBER As Long = 1 ' A列
Public Const COL_STUDENT_NAME As Long = 2 ' B列 (列番号で指定)

' 集計列の定義 (AL列からAR列)
Public Const COL_SUM_SHUTTEI As Long = 38 ' AL列 (出停)
Public Const COL_SUM_KIBIKI As Long = 39 ' AM列 (忌引き)
Public Const COL_SUM_KESSEKI As Long = 40 ' AN列 (欠席)
Public Const COL_SUM_KOUKETSU As Long = 41 ' AO列 (公欠)
Public Const COL_SUM_ATTENDANCE As Long = 42 ' AP列 (出席日数)
Public Const COL_SUM_CHIKOKU As Long = 43 ' AQ列 (遅刻)
Public Const COL_SUM_SOUTAI As Long = 44 ' AR列 (早退)


Public Const CELL_WAREKI As String = "A2"
Public Const CELL_MONTH As String = "E2"
Public Const CELL_GRADE As String = "S2"
Public Const CELL_CLASS As String = "W2"
Public Const CELL_TEACHER As String = "AD2"
Public Const CELL_SCHOOL_DAY_COUNT As String = "B52" ' 授業日数を格納するセル

' 生徒名簿CSV関連
Public Const CSV_FILE_NAME As String = "児童生徒登録用名簿.csv"
Public Const CSV_COL_CLASS As Long = 4   ' D列
Public Const CSV_COL_STUDENT_NUMBER As Long = 5 ' E列
Public Const CSV_COL_STUDENT_NAME As Long = 6 ' F列
Public Const CSV_COL_STATUS As Long = 12 ' L列 (1:有効, 0:利用停止)

' 欠席連絡CSV関連
Public Const ABSENCE_FILE_NAME As String = "欠席連絡.csv"
Public Const ABSENCE_COL_DATE As Long = 1    ' A列
Public Const ABSENCE_COL_GRADE As Long = 2   ' B列
Public Const ABSENCE_CLASS As Long = 3   ' C列
Public Const ABSENCE_COL_STUDENT_NUMBER As Long = 4 ' D列
Public Const ABSENCE_COL_CONTACT_TYPE As Long = 7 ' G列
Public Const ABSENCE_COL_REASON As Long = 10 ' J列

' その他
Public Const MAX_STUDENTS_ON_SHEET As Long = 50 ' 出席簿シートに表示する最大生徒数