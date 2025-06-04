Attribute VB_Name = "Module_AttendanceTemplate"
Option Explicit

' --- 主なプロシージャ ---

Public Sub 出席簿ひな形作成()
    Dim ws As Worksheet
    Dim teacherName As String
    Dim grade As String
    Dim classNum As String
    Dim currentYear As Integer
    Dim currentMonth As Integer
    Dim lastDay As Integer
    Dim forceSchoolDays() As Integer
    Dim forceHolidays() As Integer
    Dim vacationStartDate As Date ' 長期休暇開始日
    Dim vacationEndDate As Date   ' 長期休暇終了日
    Dim studentCount As Integer

    ' 1. シートの準備とクリア
    Set ws = PrepareAttendanceSheet()
    If ws Is Nothing Then Exit Sub ' シート準備失敗時は終了

    ' 2. ユーザー入力の取得
    ' 長期休暇の開始日と終了日を取得する引数を追加
    If Not GetUserInputs(teacherName, grade, classNum, currentYear, currentMonth, _
                         forceSchoolDays, forceHolidays, vacationStartDate, vacationEndDate) Then
        MsgBox "ユーザー入力がキャンセルされたため、処理を中断します。", vbInformation, "処理中断"
        Exit Sub ' 入力キャンセル時は終了
    End If

    ' 3. 基本情報の書き込み（和暦、月、学年、クラス、担任名）
    Call WriteBasicInfo(ws, teacherName, grade, classNum, currentYear, currentMonth)

    ' 4. 日付と曜日の生成
    lastDay = SetDatesAndWeekdays(ws, currentYear, currentMonth)

    ' 5. 生徒名簿の読み込みと転記
    studentCount = LoadAndWriteStudentList(ws, grade, classNum)
    Dim lastStudentDataRow As Long
    If studentCount > 0 Then
        lastStudentDataRow = Module_Constants.ROW_STUDENT_START + studentCount - 1 ' 定数参照を修正
    Else
        ' 生徒がいない場合でも、日付の罫線描画のために最低限の行範囲を設定
        lastStudentDataRow = Module_Constants.ROW_STUDENT_END ' 定数参照を修正
    End If

    ' 6. 休日判定、罫線と祝日名の描画、授業日数のカウント
    Dim schoolDayCount As Long
    ' ★ 修正: vacationStartDate と vacationEndDate を引数に追加
    schoolDayCount = DrawHolidayLinesAndLabels(ws, currentYear, currentMonth, _
                                             lastDay, lastStudentDataRow, _
                                             forceSchoolDays, forceHolidays, _
                                             vacationStartDate, vacationEndDate)
    ws.Range(Module_Constants.CELL_SCHOOL_DAY_COUNT).Value = schoolDayCount ' 授業日数をセルに書き込み (定数参照を修正)

    ' 7. 長期休暇の斜め線描画
    If vacationStartDate <> 0 And vacationEndDate <> 0 Then ' 日付が有効な場合のみ描画
        Call DrawSingleVacationLine(ws, currentYear, currentMonth, studentCount, vacationStartDate, vacationEndDate)
    End If

    ' 8. 集計列のクリアと準備 (欠席情報反映前にクリアしておく)
    Call ClearAndPrepareSummaryColumns(ws)

    ' 9. 欠席情報の反映確認
    If MsgBox("欠席連絡からの情報を反映しますか？", vbYesNo + vbQuestion, "欠席情報反映") = vbYes Then
        Call 欠席情報反映(grade, classNum, currentYear, currentMonth) ' 既存プロシージャを呼び出し
    End If
    
    ' 10. 集計を実行 (欠席情報反映後に実行する)
    Call CalculateAttendanceSummary(ws, lastDay, studentCount, schoolDayCount)

    MsgBox "出席簿のひな形が作成されました。" & vbCrLf & _
           "生徒名を " & studentCount & " 件登録しました。" & vbCrLf & _
           "授業日数: " & schoolDayCount & " 日", vbInformation, "完了"
End Sub