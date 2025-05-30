Attribute VB_Name = "Module_MainProcess"
Option Explicit

' --- 定数定義 ---
' シート名
Private Const SHEET_NAME_TABLE As String = "table"

' 出席簿シートのセル範囲
Private Const ROW_DATE As Long = 3
Private Const ROW_WEEKDAY As Long = 4
Private Const COL_START_DATE As Long = 7 ' G列 (日付の開始列)

Private Const ROW_STUDENT_START As Long = 6
Private Const ROW_STUDENT_END As Long = 50
Private Const COL_STUDENT_NUMBER As Long = 1 ' A列
Private Const COL_STUDENT_NAME As Long = 2 ' B列 (列番号で指定)

' 集計列の定義 (AL列からAR列)
Private Const COL_SUM_SHUTTEI As Long = 38 ' AL列 (出停)
Private Const COL_SUM_KIBIKI As Long = 39 ' AM列 (忌引き)
Private Const COL_SUM_KESSEKI As Long = 40 ' AN列 (欠席)
Private Const COL_SUM_KOUKETSU As Long = 41 ' AO列 (公欠)
Private Const COL_SUM_ATTENDANCE As Long = 42 ' AP列 (出席日数)
Private Const COL_SUM_CHIKOKU As Long = 43 ' AQ列 (遅刻)
Private Const COL_SUM_SOUTAI As Long = 44 ' AR列 (早退)


Private Const CELL_WAREKI As String = "A2"
Private Const CELL_MONTH As String = "E2"
Private Const CELL_GRADE As String = "S2"
Private Const CELL_CLASS As String = "W2"
Private Const CELL_TEACHER As String = "AD2"
Private Const CELL_SCHOOL_DAY_COUNT As String = "B52" ' 授業日数を格納するセル

' 生徒名簿CSV関連
Private Const CSV_FILE_NAME As String = "児童生徒登録用名簿.csv"
Private Const CSV_COL_CLASS As Long = 4  ' D列
Private Const CSV_COL_STUDENT_NUMBER As Long = 5 ' E列
Private Const CSV_COL_STUDENT_NAME As Long = 6 ' F列
Private Const CSV_COL_STATUS As Long = 12 ' L列 (1:有効, 0:利用停止)

' 欠席連絡CSV関連
Private Const ABSENCE_FILE_NAME As String = "欠席連絡.csv"
Private Const ABSENCE_COL_DATE As Long = 1    ' A列
Private Const ABSENCE_COL_GRADE As Long = 2  ' B列
Private Const ABSENCE_COL_CLASS As Long = 3  ' C列
Private Const ABSENCE_COL_STUDENT_NUMBER As Long = 4 ' D列
Private Const ABSENCE_COL_CONTACT_TYPE As Long = 7 ' G列
Private Const ABSENCE_COL_REASON As Long = 10 ' J列

' その他
Private Const MAX_STUDENTS_ON_SHEET As Long = 50 ' 出席簿シートに表示する最大生徒数

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
        lastStudentDataRow = ROW_STUDENT_START + studentCount - 1
    Else
        ' 生徒がいない場合でも、日付の罫線描画のために最低限の行範囲を設定
        lastStudentDataRow = ROW_STUDENT_END
    End If

    ' 6. 休日判定、罫線と祝日名の描画、授業日数のカウント
    Dim schoolDayCount As Long
    ' ★ 修正: vacationStartDate と vacationEndDate を引数に追加
    schoolDayCount = DrawHolidayLinesAndLabels(ws, currentYear, currentMonth, _
                                             lastDay, lastStudentDataRow, _
                                             forceSchoolDays, forceHolidays, _
                                             vacationStartDate, vacationEndDate)
    ws.Range(CELL_SCHOOL_DAY_COUNT).Value = schoolDayCount ' 授業日数をセルに書き込み

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

' --- プライベート関数とプロシージャ ---

' ★ 1. シートの準備とクリアを行う関数
Private Function PrepareAttendanceSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME_TABLE)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = SHEET_NAME_TABLE
    End If

    ' 既存の図形をすべて削除 (線、テキストボックスなど)
    Dim i As Integer
    For i = ws.Shapes.count To 1 Step -1
        ws.Shapes(i).Delete
    Next i

    ' クリアする範囲を指定
    Dim clearRangeData As Range
    ' その月の最大日数（31日）分の列までをクリア対象とする
    Set clearRangeData = ws.Range(ws.Cells(ROW_STUDENT_START, COL_START_DATE), ws.Cells(ROW_STUDENT_END, COL_START_DATE + 31 - 1))
    
    Dim clearRangeSummary As Range
    Set clearRangeSummary = ws.Range(ws.Cells(ROW_STUDENT_START, COL_SUM_SHUTTEI), ws.Cells(ROW_STUDENT_END, COL_SUM_SOUTAI))

    ' 和暦、当月、学年、クラス、担任名、授業日数セルをクリア
    ws.Range(CELL_WAREKI).ClearContents
    ws.Range(CELL_MONTH).ClearContents
    ws.Range(CELL_GRADE).ClearContents
    ws.Range(CELL_CLASS).ClearContents
    ws.Range(CELL_TEACHER).ClearContents
    ws.Range(CELL_SCHOOL_DAY_COUNT).ClearContents

    ' 学生番号と氏名列もクリア
    ws.Range(ws.Cells(ROW_STUDENT_START, COL_STUDENT_NUMBER), ws.Cells(ROW_STUDENT_END, COL_STUDENT_NAME)).ClearContents

    ' 日付行と曜日行のコンテンツをクリア（フォントカラーも含む）
    With ws.Range("G3:AK4")
        .ClearContents
        .Font.Color = RGB(0, 0, 0) ' 黒色にリセット
    End With
    
    ' 指定範囲のコンテンツをクリアし、フォント色をリセット
    With clearRangeData
        .ClearContents
        .Font.Color = RGB(0, 0, 0) ' 黒色にリセット
    End With
    
    With clearRangeSummary
        .ClearContents
        .Font.Color = RGB(0, 0, 0) ' 黒色にリセット
    End With

    ' コメント（メモ）を削除 (データ範囲と集計範囲に対してそれぞれ実行)
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

' ★ 2. ユーザー入力の取得を行う関数 (長期休暇の引数を追加)
Private Function GetUserInputs(ByRef teacherName As String, ByRef grade As String, ByRef classNum As String, _
                               ByRef currentYear As Integer, ByRef currentMonth As Integer, _
                               ByRef forceSchoolDays() As Integer, ByRef forceHolidays() As Integer, _
                               ByRef vacationStartDate As Date, ByRef vacationEndDate As Date) As Boolean

    Dim inputString As String
    Dim dateInput As String
    
    ' 初期値を設定
    vacationStartDate = 0
    vacationEndDate = 0

    ' 担任名
    inputString = SafeInputBox("担任の名前を入力してください。ここでキャンセルすると帳票をクリアして終了します。", "担任情報")
    If inputString = "" Then GetUserInputs = False: Exit Function
    teacherName = inputString

    ' 学年
    inputString = SafeInputBox("学年を入力してください。(半角数字:1,2など)", "学年情報")
    If inputString = "" Then GetUserInputs = False: Exit Function
    grade = inputString

    ' クラス
    inputString = SafeInputBox("クラスを入力してください。(半角英数字:1,2,Bなど)", "クラス情報")
    If inputString = "" Then GetUserInputs = False: Exit Function
    classNum = inputString

    ' 年度
    currentYear = year(Date)
    inputString = SafeInputBox("年度を入力してください。" & vbCrLf & _
                                 "現在の年度は " & currentYear & " 年です。" & vbCrLf & _
                                 "変更が必要な場合は新しい値を入力してください。", _
                                 "年度確認", CStr(currentYear))
    If IsNumeric(inputString) Then
        If CInt(inputString) >= 2000 And CInt(inputString) <= 2100 Then
            currentYear = CInt(inputString)
        Else
            MsgBox "有効な年度（2000～2100）を入力してください。既定値の " & currentYear & " 年を使用します。", vbExclamation
        End If
    ElseIf inputString <> "" Then
        MsgBox "数値以外が入力されました。既定値の " & currentYear & " 年を使用します。", vbExclamation
    End If

    ' 月
    currentMonth = month(Date)
    inputString = SafeInputBox("確認: 現在の月は " & currentMonth & " 月です。" & vbCrLf & _
                                 "変更が必要な場合は新しい値を入力してください。", _
                                 "月の確認", CStr(currentMonth))
    If IsNumeric(inputString) Then
        If CInt(inputString) >= 1 And CInt(inputString) <= 12 Then
            currentMonth = CInt(inputString)
        Else
            MsgBox "有効な月（1～12）を入力してください。既定値の " & currentMonth & " 月を使用します。", vbExclamation
        End If
    ElseIf inputString <> "" Then
        MsgBox "数値以外が入力されました。既定値の " & currentMonth + " 月を使用します。", vbExclamation
    End If

    ' 強制授業日
    Dim forceDatesInput As String
    If MsgBox("祝日や土曜日でも授業日として扱う日付はありますか？（運動会、授業参観など）", vbYesNo + vbQuestion, "授業日として扱う日付") = vbYes Then
        forceDatesInput = SafeInputBox("授業日として扱う日付をカンマ区切りで入力してください。(半角数字: 1,8,22など)", "強制授業日")
        forceSchoolDays = ParseCommaSeparatedNumbers(forceDatesInput)
    Else
        ReDim forceSchoolDays(0) ' 空の配列を初期化
    End If

    ' 強制休日
    If MsgBox("平日でも休日として扱う日付はありますか？（休校など）", vbYesNo + vbQuestion, "休日として扱う日付") = vbYes Then
        forceDatesInput = SafeInputBox("休日として扱う日付をカンマ区切りで入力してください。(半角数字: 10,12など)", "強制休日")
        forceHolidays = ParseCommaSeparatedNumbers(forceDatesInput)
    Else
        ReDim forceHolidays(0) ' 空の配列を初期化
    End If
    
    ' ★長期休暇の入力プロンプトを追加★
    If MsgBox("長期休暇はありますか？（夏季、冬季など）" & vbCrLf & _
              "長期休暇期間の該当セルに斜め線が引かれます。", vbYesNo + vbQuestion, "長期休暇設定") = vbYes Then
        dateInput = SafeInputBox("長期休暇の開始日を入力してください。(例: " & Format(Date, "yyyy/m/d") & ")", "長期休暇開始日")
        If IsDate(dateInput) Then
            vacationStartDate = CDate(dateInput)
        ElseIf dateInput <> "" Then
            MsgBox "有効な日付形式ではありません。長期休暇は設定されません。", vbExclamation
        End If
        
        If vacationStartDate <> 0 Then ' 開始日が有効な場合のみ終了日を尋ねる
            dateInput = SafeInputBox("長期休暇の終了日を入力してください。(例: " & Format(Date, "yyyy/m/d") & ")", "長期休暇終了日")
            If IsDate(dateInput) Then
                vacationEndDate = CDate(dateInput)
                If vacationEndDate < vacationStartDate Then
                    MsgBox "終了日が開始日より前の日付です。長期休暇は設定されません。", vbExclamation
                    vacationStartDate = 0 ' 無効な日付としてリセット
                    vacationEndDate = 0
                End If
            ElseIf dateInput <> "" Then
                MsgBox "有効な日付形式ではありません。長期休暇は設定されません。", vbExclamation
                vacationStartDate = 0 ' 無効な日付としてリセット
                vacationEndDate = 0
            End If
        End If
    End If

    GetUserInputs = True
End Function

' ★ ユーザー入力補助関数（キャンセル時の空文字列を返す）
Private Function SafeInputBox(prompt As String, title As String, Optional defaultValue As Variant) As String
    On Error Resume Next
    If IsMissing(defaultValue) Then
        SafeInputBox = InputBox(prompt, title)
    Else
        SafeInputBox = InputBox(prompt, title, defaultValue)
    End If
    On Error GoTo 0
End Function

' ★ カンマ区切りの数字文字列を整数配列に変換する関数
Private Function ParseCommaSeparatedNumbers(ByVal inputString As String) As Integer()
    Dim tempArray() As String
    Dim resultArray() As Integer
    Dim count As Long
    Dim i As Long

    If Trim(inputString) = "" Then
        ReDim resultArray(0) ' 空の配列
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
        ReDim resultArray(0) ' 有効な数字がない場合
    End If

    ParseCommaSeparatedNumbers = resultArray
End Function

' ★ 3. 基本情報をシートに書き込むプロシージャ
Private Sub WriteBasicInfo(ws As Worksheet, teacherName As String, grade As String, classNum As String, _
                           currentYear As Integer, currentMonth As Integer)
    ws.Range(CELL_GRADE).Value = grade
    ws.Range(CELL_CLASS).Value = classNum
    ws.Range(CELL_TEACHER).Value = teacherName

    Dim wareki As String
    ' 令和の開始年を2019年として計算 (2019年 = 令和1年)
    wareki = "令和" & (currentYear - 2018) & "年"
    ws.Range(CELL_WAREKI).Value = wareki

    ws.Range(CELL_MONTH).Value = currentMonth
End Sub

' ★ 4. 日付と曜日をシートに設定する関数
Private Function SetDatesAndWeekdays(ws As Worksheet, currentYear As Integer, currentMonth As Integer) As Integer
    Dim lastDay As Integer
    lastDay = day(DateSerial(currentYear, currentMonth + 1, 0)) ' その月の最終日を計算

    Dim i As Integer
    Dim currentDate As Date

    ' 日付を入力
    For i = 1 To lastDay
        ws.Cells(ROW_DATE, COL_START_DATE + i - 1).Value = i
    Next i

    ' 曜日を入力
    For i = 1 To lastDay
        currentDate = DateSerial(currentYear, currentMonth, i)
        ws.Cells(ROW_WEEKDAY, COL_START_DATE + i - 1).Value = WeekdayName(Weekday(currentDate), True)
    Next i

    SetDatesAndWeekdays = lastDay
End Function

' ★ 5. 生徒名簿を読み込み、シートに転記する関数
Private Function LoadAndWriteStudentList(ws As Worksheet, grade As String, classNum As String) As Integer
    Dim csvFile As String
    Dim csvWB As Workbook
    Dim csvWS As Worksheet
    Dim lastRow As Long
    Dim targetClass As String
    Dim studentCounter As Integer
    Dim i As Long

    ' 例: "1年A組" の形式に合わせる
    targetClass = grade & "年" & classNum & "組"
    csvFile = ThisWorkbook.Path & "\" & CSV_FILE_NAME
    studentCounter = 0

    If Dir(csvFile) = "" Then
        MsgBox "生徒名簿CSVファイルが見つかりません: " & csvFile, vbExclamation, "ファイル未検出"
        LoadAndWriteStudentList = 0
        Exit Function
    End If

    On Error Resume Next
    Set csvWB = Workbooks.Open(csvFile)
    If Err.Number <> 0 Then
        MsgBox "CSVファイルを開けませんでした: " & Err.Description, vbCritical, "エラー"
        LoadAndWriteStudentList = 0
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    Set csvWS = csvWB.Worksheets(1)
    lastRow = csvWS.Cells(csvWS.Rows.count, "A").End(xlUp).row ' CSVファイルの最終行を取得

    For i = 2 To lastRow ' ヘッダー行をスキップ
        ' クラスが一致し、かつステータスが「有効(1)」の生徒のみを抽出
        If csvWS.Cells(i, CSV_COL_CLASS).Value = targetClass And csvWS.Cells(i, CSV_COL_STATUS).Value = 1 Then
            studentCounter = studentCounter + 1
            If studentCounter <= MAX_STUDENTS_ON_SHEET Then ' シートに表示する最大生徒数を超えない場合
                ws.Cells(ROW_STUDENT_START + studentCounter - 1, COL_STUDENT_NUMBER).Value = csvWS.Cells(i, CSV_COL_STUDENT_NUMBER).Value
                ws.Cells(ROW_STUDENT_START + studentCounter - 1, COL_STUDENT_NAME).Value = csvWS.Cells(i, CSV_COL_STUDENT_NAME).Value
            End If
        End If
    Next i

    csvWB.Close False ' CSVファイルを保存せずに閉じる
    LoadAndWriteStudentList = studentCounter ' 登録した生徒数を返す
End Function

' ★ 6. 休日判定、罫線と祝日名の描画、授業日数のカウントを行う関数
Private Function DrawHolidayLinesAndLabels(ws As Worksheet, currentYear As Integer, currentMonth As Integer, _
                                            lastDay As Integer, lastStudentDataRow As Long, _
                                            forceSchoolDays() As Integer, forceHolidays() As Integer, _
                                            vacationStartDate As Date, vacationEndDate As Date) As Long ' ★ 追加引数
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

        ' ★ ここから追加: 長期休暇期間内のチェック ★
        If vacationStartDate <> 0 And vacationEndDate <> 0 Then ' 長期休暇が設定されている場合のみ
            If currentDate >= vacationStartDate And currentDate <= vacationEndDate Then
                ' 長期休暇期間内の日付は、個別の縦線と祝日名を描画しない
                ' ただし、日付と曜日の赤字表示は長期休暇の処理で行うため、ここでは何もしない
                ' 授業日数もカウントしない（長期休暇中は授業日ではないため）
                GoTo NextDayLoop ' 次の日の処理へスキップ
            End If
        End If
        ' ★ ここまで追加 ★

        ' 強制休日リストに含まれるかチェック（最優先）
        isForceHoliday = False
        For j = LBound(forceHolidays) To UBound(forceHolidays)
            If forceHolidays(j) = i Then
                isForceHoliday = True
                Exit For
            End If
        Next j
        If isForceHoliday Then
            isHoliday = True
            holidayName = "休校" ' 強制休日の場合、休日名を「休校」とする
        End If

        ' 強制授業日リストに含まれるかチェック（強制休日が優先されるため、isForceHolidayがFalseの場合のみ適用）
        isForceSchoolDay = False
        For j = LBound(forceSchoolDays) To UBound(forceSchoolDays)
            If forceSchoolDays(j) = i Then
                isForceSchoolDay = True
                Exit For
            End If
        Next j
        If isForceSchoolDay And Not isForceHoliday Then ' 強制休日でなければ強制授業日を適用
            isHoliday = False
            holidayName = "" ' 強制授業日の場合、休日名なし
        End If

        ' 強制設定が適用されていない場合、通常の休日判定を行う
        If Not isForceHoliday And Not isForceSchoolDay Then
            ' 日曜日の場合
            If Weekday(currentDate) = vbSunday Then
                isHoliday = True
            End If

            ' 土曜日の場合
            If Weekday(currentDate) = vbSaturday Then
                isHoliday = True
            End If

            ' 祝日の場合
            nationalHolidayName = GetNationalHolidayName(currentDate)
            If nationalHolidayName <> "" Then
                isHoliday = True
                holidayName = nationalHolidayName
            End If
        End If

        If isHoliday Then
            ' 日付セルを赤字に (長期休暇期間外の休日のみ)
            ws.Cells(ROW_DATE, COL_START_DATE + i - 1).Font.Color = RGB(255, 0, 0)
            ' 曜日セルを赤字に (長期休暇期間外の休日のみ)
            ws.Cells(ROW_WEEKDAY, COL_START_DATE + i - 1).Font.Color = RGB(255, 0, 0)

            ' 縦線を引く
            leftPosition = ws.Cells(ROW_DATE, COL_START_DATE + i - 1).Left + ws.Cells(ROW_DATE, COL_START_DATE + i - 1).Width / 2
            topPosition = ws.Rows(ROW_STUDENT_START).Top
            ' 生徒データの最終行まで線を描画
            heightOfLine = (ws.Cells(lastStudentDataRow, COL_STUDENT_NAME).Top + ws.Cells(lastStudentDataRow, COL_STUDENT_NAME).Height) - topPosition

            Set lineShape = ws.Shapes.AddLine( _
                leftPosition, topPosition, _
                leftPosition, topPosition + heightOfLine)

            With lineShape.Line
                .Weight = 1.5
                .DashStyle = msoLineSolid
                .ForeColor.RGB = RGB(255, 0, 0)
            End With

            ' 祝日名が設定されている場合のみテキストボックスを追加
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
            ' 休日ではない場合、授業日数をカウント
            schoolDayCount = schoolDayCount + 1
            ' 日付セルと曜日セルを黒字にリセット (長期休暇期間外の平日のみ)
            ' 長期休暇期間内の平日はDrawSingleVacationLineで赤字にしているので、ここでは処理しない
            ' ★ 修正: 既に赤字になっていないか確認してから黒字にリセット ★
            If ws.Cells(ROW_DATE, COL_START_DATE + i - 1).Font.Color <> RGB(255, 0, 0) Then
                ws.Cells(ROW_DATE, COL_START_DATE + i - 1).Font.Color = RGB(0, 0, 0)
                ws.Cells(ROW_WEEKDAY, COL_START_DATE + i - 1).Font.Color = RGB(0, 0, 0)
            End If
        End If
NextDayLoop: ' GoTo ステートメントのターゲット
    Next i

    DrawHolidayLinesAndLabels = schoolDayCount
End Function

' ★ 祝日名を判定する関数
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
            If targetDay = 1 Then holidayName = "元日"
            If GetNthMonday(targetYear, 1, 2) = targetDay Then holidayName = "成人の日"

        Case 2
            If targetDay = 11 Then holidayName = "建国記念の日"
            If targetDay = 23 Then holidayName = "天皇誕生日"

        Case 3
            If targetDay = GetShunbun(targetYear) Then holidayName = "春分の日"

        Case 4
            If targetDay = 29 Then holidayName = "昭和の日"

        Case 5
            If targetDay = 3 Then holidayName = "憲法記念日"
            If targetDay = 4 Then holidayName = "みどりの日"
            If targetDay = 5 Then holidayName = "こどもの日"
            ' 振替休日
            If targetYear >= 2007 And targetMonth = 5 And targetDay = 6 And Weekday(DateSerial(targetYear, 5, 3), vbSunday) = 1 And Weekday(DateSerial(targetYear, 5, 4), vbSunday) <> 1 And Weekday(DateSerial(targetYear, 5, 5), vbSunday) = 1 Then
                holidayName = "振替休日"
            End If

        Case 7
            If targetYear = 2020 And targetDay = 23 Then
                holidayName = "海の日"
            ElseIf targetYear = 2021 And targetDay = 22 Then
                holidayName = "海の日"
            ElseIf GetNthMonday(targetYear, 7, 3) = targetDay Then
                holidayName = "海の日"
            End If

        Case 8
            If targetYear = 2020 And targetDay = 10 Then
                holidayName = "山の日"
            ElseIf targetYear = 2021 And targetDay = 8 Then
                holidayName = "山の日"
            ElseIf targetDay = 11 Then
                holidayName = "山の日"
            End If

        Case 9
            If GetNthMonday(targetYear, 9, 3) = targetDay Then holidayName = "敬老の日"
            If targetDay = GetShubun(targetYear) Then holidayName = "秋分の日"
            ' 国民の休日 (敬老の日と秋分の日に挟まれた平日)
            If targetYear >= 1986 Then
                Dim kRDate As Date
                Dim sBDate As Date
                kRDate = DateSerial(targetYear, 9, GetNthMonday(targetYear, 9, 3))
                sBDate = DateSerial(targetYear, 9, GetShubun(targetYear))
                If targetDate > kRDate And targetDate < sBDate And Weekday(targetDate, vbSunday) <> 1 And Weekday(targetDate, vbSaturday) <> 7 Then
                    holidayName = "国民の休日"
                End If
            End If

        Case 10
            If targetYear = 2020 And targetDay = 24 Then
                holidayName = "スポーツの日"
            ElseIf targetYear = 2021 And targetDay = 18 Then
                holidayName = "スポーツの日"
            ElseIf targetYear >= 2020 And GetNthMonday(targetYear, 10, 2) = targetDay Then
                holidayName = "スポーツの日"
            End If

        Case 11
            If targetDay = 3 Then holidayName = "文化の日"
            If targetDay = 23 Then holidayName = "勤労感謝の日"
    End Select

    ' ここで振替休日を再チェック (重複定義を避けるため)
    If holidayName = "" Then
        If IsSubstituteHoliday(targetDate) Then
            holidayName = "振替休日"
        End If
    End If

    GetNationalHolidayName = holidayName
End Function

' ★ 第N月曜日を計算する関数
Private Function GetNthMonday(targetYear As Integer, targetMonth As Integer, n As Integer) As Integer
    Dim firstDay As Date
    Dim firstWeekday As Integer ' 1=日, 2=月, ..., 7=土 (vbMondayの場合)
    Dim firstMonday As Integer

    firstDay = DateSerial(targetYear, targetMonth, 1)
    firstWeekday = Weekday(firstDay, vbMonday) ' 週の始まりを月曜日に設定

    If firstWeekday = 1 Then ' 1日が月曜日だった場合
        firstMonday = 1
    Else ' 1日が月曜日以外だった場合
        firstMonday = 8 - firstWeekday + 1
    End If

    GetNthMonday = firstMonday + (n - 1) * 7
End Function

' ★ 春分の日を計算する関数
Private Function GetShunbun(targetYear As Integer) As Integer
    ' 概算式 (正確な天文学的計算ではないが、VBA用途としては十分)
    GetShunbun = Int(20.8431 + 0.242194 * (targetYear - 1851) - Int((targetYear - 1851) / 4))
End Function

' ★ 秋分の日を計算する関数
Private Function GetShubun(targetYear As Integer) As Integer
    ' 概算式 (正確な天文学的計算ではないが、VBA用途としては十分)
    GetShubun = Int(23.2488 + 0.242194 * (targetYear - 1851) - Int((targetYear - 1851) / 4))
End Function

' ★ 振替休日を判定する関数
Private Function IsSubstituteHoliday(targetDate As Date) As Boolean
    Dim checkDate As Date
    Dim i As Integer

    IsSubstituteHoliday = False

    If Weekday(targetDate) <> vbSunday Then Exit Function ' 日曜日でなければ振替休日ではない

    ' 過去7日間を遡って、日曜日と重なる祝日があるかチェック
    For i = 1 To 7
        checkDate = targetDate - i
        ' 振替休日を考慮しない純粋な祝日判定関数を使用
        If GetNationalHolidayNameWithoutSubstitute(checkDate) <> "" And Weekday(checkDate) = vbSunday Then
            IsSubstituteHoliday = True
            Exit Function
        End If
    Next i
End Function

' ★ 振替休日を考慮しない純粋な祝日判定 (GetNationalHolidayNameから振替休日判定を除外)
Private Function GetNationalHolidayNameWithoutSubstitute(targetDate As Date) As String
    Dim targetYear As Integer
    Dim targetMonth As Integer
    Dim targetDay As Integer
    Dim holidayName As String

    targetYear = year(targetDate)
    targetMonth = month(targetDate)
    targetDay = day(targetDate)

    holidayName = ""

    ' 週休日（土日）はここでは祝日として扱わない
    If Weekday(targetDate) <> vbSunday And Weekday(targetDate) <> vbSaturday Then
        Select Case targetMonth
            Case 1
                If targetDay = 1 Then holidayName = "元日"
                If GetNthMonday(targetYear, 1, 2) = targetDay Then holidayName = "成人の日"
            Case 2
                If targetDay = 11 Then holidayName = "建国記念の日"
                If targetDay = 23 Then holidayName = "天皇誕生日"
            Case 3
                If targetDay = GetShunbun(targetYear) Then holidayName = "春分の日"
            Case 4
                If targetDay = 29 Then holidayName = "昭和の日"
            Case 5
                If targetDay = 3 Then holidayName = "憲法記念日"
                If targetDay = 4 Then holidayName = "みどりの日"
                If targetDay = 5 Then holidayName = "こどもの日"
            Case 7
                If targetYear = 2020 And targetDay = 23 Then
                    holidayName = "海の日"
                ElseIf targetYear = 2021 And targetDay = 22 Then
                    holidayName = "海の日"
                ElseIf GetNthMonday(targetYear, 7, 3) = targetDay Then
                    holidayName = "海の日"
                End If
            Case 8
                If targetYear = 2020 And targetDay = 10 Then
                    holidayName = "山の日"
                ElseIf targetYear = 2021 And targetDay = 8 Then
                    holidayName = "山の日"
                ElseIf targetDay = 11 Then
                    holidayName = "山の日"
                End If
            Case 9
                If GetNthMonday(targetYear, 9, 3) = targetDay Then holidayName = "敬老の日"
                If targetDay = GetShubun(targetYear) Then holidayName = "秋分の日"
                If targetYear >= 1986 Then
                    Dim kRDate As Date, sBDate As Date
                    kRDate = DateSerial(targetYear, 9, GetNthMonday(targetYear, 9, 3))
                    sBDate = DateSerial(targetYear, 9, GetShubun(targetYear))
                    If targetDate > kRDate And targetDate < sBDate And Weekday(targetDate, vbSunday) <> 1 And Weekday(targetDate, vbSaturday) <> 7 Then
                        holidayName = "国民の休日"
                    End If
                End If
            Case 10
                If targetYear = 2020 And targetDay = 24 Then
                    holidayName = "スポーツの日"
                ElseIf targetYear = 2021 And targetDay = 18 Then
                    holidayName = "スポーツの日"
                ElseIf targetYear >= 2020 And GetNthMonday(targetYear, 10, 2) = targetDay Then
                    holidayName = "スポーツの日"
                End If
            Case 11
                If targetDay = 3 Then holidayName = "文化の日"
                If targetDay = 23 Then holidayName = "勤労感謝の日"
        End Select
    End If
    
    GetNationalHolidayNameWithoutSubstitute = holidayName
End Function

' ★ 長期休暇のセル範囲に1本の斜め線を引くプロシージャ
Private Sub DrawSingleVacationLine(ws As Worksheet, currentYear As Integer, currentMonth As Integer, _
                                  studentCount As Integer, vacationStartDate As Date, vacationEndDate As Date)
    Dim startDay As Integer
    Dim endDay As Integer
    Dim firstDateOfSheet As Date
    Dim lastDateOfSheet As Date
    
    ' 現在のシートが示す年月
    firstDateOfSheet = DateSerial(currentYear, currentMonth, 1)
    lastDateOfSheet = DateSerial(currentYear, currentMonth, day(DateSerial(currentYear, currentMonth + 1, 0)))

    ' 長期休暇の開始日と終了日が現在のシートの月にどの程度含まれるか計算
    If vacationStartDate < firstDateOfSheet Then
        startDay = 1
    ElseIf year(vacationStartDate) = currentYear And month(vacationStartDate) = currentMonth Then
        startDay = day(vacationStartDate)
    Else
        Exit Sub ' 休暇開始日が現在のシートの月より後の場合、このシートでは描画しない
    End If

    If vacationEndDate > lastDateOfSheet Then
        endDay = day(lastDateOfSheet)
    ElseIf year(vacationEndDate) = currentYear And month(vacationEndDate) = currentMonth Then
        endDay = day(vacationEndDate)
    Else
        Exit Sub ' 休暇終了日が現在のシートの月より前の場合、このシートでは描画しない
    End If

    ' 描画する範囲が有効かチェック
    If startDay > endDay Then Exit Sub
    If studentCount = 0 Then Exit Sub ' 生徒がいない場合は描画しない

    Dim topLeftCell As Range
    Dim bottomRightCell As Range
    Dim lineShape As Shape
    
    ' 描画の開始セル（左上）
    Set topLeftCell = ws.Cells(ROW_STUDENT_START, COL_START_DATE + startDay - 1)
    
    ' 描画の終了セル（右下）
    Set bottomRightCell = ws.Cells(ROW_STUDENT_START + studentCount - 1, COL_START_DATE + endDay - 1)
    
    ' 1本の斜め線を描画 (左上から右下へ)
    Set lineShape = ws.Shapes.AddLine( _
        topLeftCell.Left, topLeftCell.Top, _
        bottomRightCell.Left + bottomRightCell.Width, bottomRightCell.Top + bottomRightCell.Height)
    
    With lineShape.Line
        .Weight = 1.5 ' 線の太さ
        .DashStyle = msoLineSolid ' 実線
        .ForeColor.RGB = RGB(255, 150, 150) ' 薄い灰色
    End With
    
    ' 線を背面に移動して、欠席記号などが見えるようにする
    lineShape.ZOrder msoSendToBack
    
    ' ★ ここから追加・修正: 期間の曜日を赤字に ★
    Dim dayIndex As Integer
    For dayIndex = startDay To endDay
        Dim targetCol As Long
        targetCol = COL_START_DATE + dayIndex - 1
        
        ' 日付セルを赤字に
        ws.Cells(ROW_DATE, targetCol).Font.Color = RGB(255, 0, 0)
        ' 曜日セルを赤字に
        ws.Cells(ROW_WEEKDAY, targetCol).Font.Color = RGB(255, 0, 0)
    Next dayIndex
    ' ★ ここまで追加・修正 ★

    ' 長期休暇であることを示すテキストボックスを追加
    Dim txtBox As Shape
    Dim middleDay As Integer
    
    ' テキストボックスを配置するセルの目安 (期間の中央あたり)
    middleDay = Int((startDay + endDay) / 2)
    If middleDay = 0 Then middleDay = startDay ' 1日だけの場合の対策
    
    Dim targetCellForText As Range
    Set targetCellForText = ws.Cells(ROW_STUDENT_START + Int(studentCount / 2), COL_START_DATE + middleDay - 1)
    
    If targetCellForText Is Nothing Then
        Debug.Print "Debug: targetCellForTextがNothingです。テキストボックスの作成をスキップします。"
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
            .TextFrame.Characters.Text = "長期休暇"
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
                Debug.Print "Debug: テキストボックスのWidthまたはHeightが不正です。配置をスキップ。"
            End If

            .ZOrder msoSendToBack
            On Error GoTo 0
        End With
    Else
        Debug.Print "Debug: txtBoxがNothingです。AddTextboxが失敗しました。"
        Debug.Print "Debug: targetCellForText.Left: " & targetCellForText.Left & ", targetCellForText.Top: " & targetCellForText.Top
    End If
    Exit Sub
    
AddTextboxErrorHandler:
    Debug.Print "Debug: AddTextboxでエラーが発生しました。エラーコード: " & Err.Number & ", " & Err.Description
    Err.Clear
End Sub

' ★ 集計列をクリアするプロシージャ
Private Sub ClearAndPrepareSummaryColumns(ws As Worksheet)
    ' AL列からAR列までの集計範囲をクリア
    ws.Range(ws.Cells(ROW_STUDENT_START, COL_SUM_SHUTTEI), ws.Cells(ROW_STUDENT_END, COL_SUM_SOUTAI)).ClearContents
End Sub


' ★ 欠席情報反映プロシージャ
Public Sub 欠席情報反映(Optional ByVal inputGrade As String = "", Optional ByVal inputClass As String = "", _
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
    
    ' tableシートを選択
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME_TABLE)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "出席簿シート(" & SHEET_NAME_TABLE & ")が見つかりません。先に出席簿ひな形を作成してください。", vbExclamation
        Exit Sub
    End If
    
    ' プロシージャの引数が省略された場合、シートの既存値を使用
    If inputGrade = "" Then
        targetGrade = Replace(ws.Range(CELL_GRADE).Value, "学年", "")
    Else
        targetGrade = inputGrade
    End If
    
    If inputClass = "" Then
        targetClass = ws.Range(CELL_CLASS).Value
    Else
        targetClass = inputClass
    End If
    
    If inputYear = 0 Then
        currentYear = year(Date) ' デフォルト値として現在の年を使用
        If ws.Range(CELL_WAREKI).Value <> "" Then ' シートに年度情報があればそれを使う
            Dim warekiStr As String
            warekiStr = ws.Range(CELL_WAREKI).Value
            If InStr(warekiStr, "令和") > 0 Then
                Dim reiwaYear As String
                reiwaYear = Split(warekiStr, "令和")(1)
                reiwaYear = Split(reiwaYear, "年")(0)
                If IsNumeric(reiwaYear) Then
                    currentYear = CInt(reiwaYear) + 2018 ' 令和1年 = 2019年
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
    
    ' CSV内のクラス表記 (例: "1年A組") に合わせる
    Dim fullTargetClass As String
    fullTargetClass = targetGrade & "年" & targetClass & "組"
    
    absenceFile = ThisWorkbook.Path & "\" & ABSENCE_FILE_NAME
    
    If Dir(absenceFile) = "" Then
        MsgBox "欠席連絡CSVファイルが見つかりません: " & absenceFile, vbExclamation, "ファイル未検出"
        Exit Sub
    End If ' ★ ここを修正しました！
    
    On Error Resume Next
    Set absenceWB = Workbooks.Open(absenceFile)
    If Err.Number <> 0 Then
        MsgBox "CSVファイルを開けませんでした: " & Err.Description, vbCritical, "エラー"
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    Set absenceWS = absenceWB.Worksheets(1)
    lastRow = absenceWS.Cells(absenceWS.Rows.count, "A").End(xlUp).row
    
    Dim processedCount As Long
    processedCount = 0
    
    For i = 2 To lastRow ' ヘッダー行をスキップ
        Dim dateString As String
        Dim dateParts As Variant
        
        dateString = absenceWS.Cells(i, ABSENCE_COL_DATE).Value
        ' 例: "2023/04/01(土)" の形式から日付部分のみを抽出
        dateParts = Split(dateString, "(")
        If UBound(dateParts) >= 0 Then
            On Error Resume Next
            absenceDate = CDate(dateParts(0)) ' 日付文字列をDate型に変換
            On Error GoTo 0
            
            If absenceDate = 0 Then ' 日付変換に失敗した場合 (CDateは無効な日付を0に変換する)
                Debug.Print "日付変換エラー: " & dateString + " (行: " & i & ")"
            Else
                ' 対象の年度と月に合致するか確認
                If year(absenceDate) = currentYear And month(absenceDate) = currentMonth Then
                    Dim csvGrade As String
                    Dim csvClass As String
                    
                    csvGrade = absenceWS.Cells(i, ABSENCE_COL_GRADE).Value
                    csvClass = absenceWS.Cells(i, ABSENCE_COL_CLASS).Value
                    
                    ' 学年とクラスが一致するか確認 (csvGradeは"1年生"のような形式を想定)
                    If InStr(csvGrade, targetGrade) > 0 And csvClass = fullTargetClass Then
                        studentNumber = CInt(absenceWS.Cells(i, ABSENCE_COL_STUDENT_NUMBER).Value)
                        contactType = Trim(absenceWS.Cells(i, ABSENCE_COL_CONTACT_TYPE).Value)
                        
                        ' 連絡種別に応じた記号を設定
                        Select Case contactType
                            Case "欠席"
                                attendanceSymbol = "×"
                            Case "遅刻"
                                attendanceSymbol = "チ"
                            Case "早退"
                                attendanceSymbol = "ハ"
                            Case "遅刻/早退"
                                attendanceSymbol = "チハ"
                            Case "忌引き"
                                attendanceSymbol = "キ"
                            Case "出停"
                                attendanceSymbol = "テ"
                            Case "/"
                                attendanceSymbol = "/" ' 公欠の記号として
                            Case Else
                                attendanceSymbol = "×" ' 未定義の場合は欠席
                        End Select
                        
                        reasonText = absenceWS.Cells(i, ABSENCE_COL_REASON).Value ' 理由を取得

                        foundStudent = False
                        ' 出席簿シートから該当生徒を探す
                        For j = ROW_STUDENT_START To ROW_STUDENT_END
                            If ws.Cells(j, COL_STUDENT_NUMBER).Value = studentNumber Then
                                studentRow = j
                                foundStudent = True
                                Exit For
                            End If
                        Next j

                        If foundStudent Then
                            absenceDay = day(absenceDate)
                            ' 出席簿シートの該当セルに記号を記入
                            ws.Cells(studentRow, COL_START_DATE + absenceDay - 1).Value = attendanceSymbol
                            
                            ' 理由があればコメントとして追加
                            If reasonText <> "" Then
                                On Error Resume Next
                                If Not ws.Cells(studentRow, COL_START_DATE + absenceDay - 1).Comment Is Nothing Then
                                    ws.Cells(studentRow, COL_START_DATE + absenceDay - 1).Comment.Delete ' 既存コメントを削除
                                End If
                                ' 新しいコメントを追加し、テキストを設定
                                ws.Cells(studentRow, COL_START_DATE + absenceDay - 1).AddComment reasonText
                                ' コメントの自動サイズ調整
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
    
    absenceWB.Close False ' CSVファイルを保存せずに閉じる
    
    If processedCount > 0 Then
        MsgBox "欠席情報の反映が完了しました。" & vbCrLf & _
               processedCount & " 件の情報を反映しました。", vbInformation, "完了"
    Else
        MsgBox "該当する欠席情報はありませんでした。", vbInformation, "完了"
    End If
End Sub

' ★ 出席簿の集計を行うプロシージャ
Private Sub CalculateAttendanceSummary(ws As Worksheet, lastDayOfMonth As Integer, studentCount As Integer, schoolDayCount As Long)
    Dim row As Long
    Dim col As Long
    Dim studentDataRange As Range
    Dim cellValue As Variant
    Dim cell As Range ' 変数宣言を追加
    
    Dim countShuttei As Long
    Dim countKibiki As Long
    Dim countKesseki As Long
    Dim countKouketsu As Long
    Dim countChikoku As Long
    Dim countSoutai As Long
    Dim countAttendance As Long ' 出席日数
    
    ' 各生徒についてループ
    For row = ROW_STUDENT_START To ROW_STUDENT_START + studentCount - 1
        ' 各生徒の集計をリセット
        countShuttei = 0
        countKibiki = 0
        countKesseki = 0
        countKouketsu = 0
        countChikoku = 0
        countSoutai = 0
        countAttendance = 0
        
        ' その生徒の日付範囲 (COL_START_DATE からその月の最終日まで)
        Set studentDataRange = ws.Range(ws.Cells(row, COL_START_DATE), ws.Cells(row, COL_START_DATE + lastDayOfMonth - 1))
        
        ' 日付範囲内の各セルをチェック
        For Each cell In studentDataRange
            cellValue = Trim(CStr(cell.Value)) ' セルの値を取得し、文字列としてトリム
            
            Select Case cellValue
                Case "テ"
                    countShuttei = countShuttei + 1
                Case "キ"
                    countKibiki = countKibiki + 1
                Case "×"
                    countKesseki = countKesseki + 1
                Case "/"
                    countKouketsu = countKouketsu + 1
                Case "チ"
                    countChikoku = countChikoku + 1
                Case "ハ"
                    countSoutai = countSoutai + 1
                Case "チハ"
                    countChikoku = countChikoku + 1 ' 遅刻にカウント
                    countSoutai = countSoutai + 1   ' 早退にカウント
                Case "" ' 空白の場合は出席とみなす
                    ' その日の日付セルが赤字（休日）でなければ出席とカウント
                    Dim dayCol As Long
                    dayCol = cell.Column
                    
                    If ws.Cells(ROW_DATE, dayCol).Font.Color <> RGB(255, 0, 0) Then
                        countAttendance = countAttendance + 1
                    End If
            End Select
        Next cell
        
        ' 集計結果をシートに書き込み
        ws.Cells(row, COL_SUM_SHUTTEI).Value = countShuttei
        ws.Cells(row, COL_SUM_KIBIKI).Value = countKibiki
        ws.Cells(row, COL_SUM_KESSEKI).Value = countKesseki
        ws.Cells(row, COL_SUM_KOUKETSU).Value = countKouketsu
        ws.Cells(row, COL_SUM_ATTENDANCE).Value = countAttendance
        ws.Cells(row, COL_SUM_CHIKOKU).Value = countChikoku
        ws.Cells(row, COL_SUM_SOUTAI).Value = countSoutai
    Next row
End Sub

