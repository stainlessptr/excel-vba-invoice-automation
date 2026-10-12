Attribute VB_Name = "InvoiceAutomation"
Sub AddNewRow()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim targetRow As Long
    Dim btn As Object ' 使用 Object 型でフォームコントロールボタンを参照
    Dim btnPrint As Object ' 印刷ボタン
    Dim targetCell As Range
    Dim printCell As Range


    ' 一覧ページを指定
    Set ws = ThisWorkbook.Sheets("list")
    
    ' 最終行と新しい行の位置を決定
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    targetRow = lastRow + 1

    ' 新しい行を挿入
    ws.Rows(targetRow).Insert Shift:=xlDown

    ' 新しい行データを初期化 申請日
    ws.Cells(targetRow, 1).Value = Date ' 現在の日付を1列目に挿入（フォーマットはYYYY/MM/DD）
    
    ' 新しい行データを初期化 名前列
    ws.Cells(targetRow, 2).Value = "" ' 名前列 (B列) を空白に設定
    
    ' 新しい行データを初期化 御中/様の選択列
    With ws.Cells(targetRow, 3).Validation
        .Delete ' 既存のデータ検証を削除
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="御中,様" ' ドロップダウンリストの選択肢を設定
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
    End With
    ws.Cells(targetRow, 3).Value = "御中" ' デフォルト値を「御中」に設定

    ' 申請総額の式を設定：各行の9、13、17、21列の合計
    ws.Cells(targetRow, 4).Formula = BuildTotalAmountFormula(targetRow) ' J、M、P、S、V、Y列の数値合計に10%を加算し、源泉所得税がある場合は加算してD列へ表示

    ' G列：源泉所得税の有無を0/1から選択
    With ws.Cells(targetRow, 7).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="0,1"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
    End With
    ws.Cells(targetRow, 7).Value = 0

    ' 明細入力欄を空白にする（現在は数量×単価の自動計算を使用しない）
    ws.Cells(targetRow, 8).Value = ""  ' H列：内容
    ws.Cells(targetRow, 9).Value = ""  ' I列：期間
    ws.Cells(targetRow, 10).Value = "" ' J列：金額

    ws.Cells(targetRow, 13).Value = "" ' M列：金額
    ws.Cells(targetRow, 16).Value = "" ' P列：金額
    ws.Cells(targetRow, 19).Value = "" ' S列：金額
    ws.Cells(targetRow, 22).Value = "" ' V列：金額
    ws.Cells(targetRow, 25).Value = "" ' Y列：金額
    

    ' 新しい行に上下の枠線を設定（実線(細)）
    With ws.Rows(targetRow).Borders(xlEdgeTop)
        .LineStyle = xlContinuous '実線(細)
        .Weight = xlThin ' 線の太さ
    End With
    With ws.Rows(targetRow).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous '実線(細)
        .Weight = xlThin ' 線の太さ
    End With
    
    
    ' 新しい行に「申請書作成」ボタンを追加
    Set targetCell = ws.Cells(targetRow, 5)  ' ボタンをE列に配置
    Set btn = ws.Shapes.AddFormControl(xlButtonControl, _
                                       targetCell.Left + 3, _
                                       targetCell.Top + 3, _
                                       targetCell.Width - 6, _
                                       targetCell.Height - 6)
    btn.OnAction = "GenerateRequestSheet" ' ボタンイベントをマクロにバインド
    btn.OLEFormat.Object.Caption = "申請書を作成" ' ボタンのテキストを設定

    ' 新しい行に「印刷」ボタンを追加
    Set printCell = ws.Cells(targetRow, 6) ' ボタンをF列に配置
    Set btnPrint = ws.Shapes.AddFormControl(xlButtonControl, _
                                            printCell.Left + 3, _
                                            printCell.Top + 3, _
                                            printCell.Width - 6, _
                                            printCell.Height - 6)
    btnPrint.OnAction = "PrintRequestSheet" ' 印刷ボタンイベントをマクロにバインド
    btnPrint.OLEFormat.Object.Caption = "印刷" ' 印刷ボタンのテキストを設定
    

    ' ボタンのフォントスタイルを設定
    With btn.OLEFormat.Object
        .Font.Bold = True ' フォントを太字に設定
        .Font.Color = RGB(0, 0, 255) ' フォントカラーを青に設定
    End With

    ' ボタンイベントをマクロにバインド
    btn.OnAction = "GenerateRequestSheet"
    
    ' ボタンのフォントスタイルを設定
    With btnPrint.OLEFormat.Object
        .Font.Bold = True ' フォントを太字に設定
        .Font.Color = RGB(0, 0, 255) ' フォントカラーを青に設定
    End With

    ' ボタンイベントをマクロにバインド
    btnPrint.OnAction = "PrintRequestSheet"
    
    MsgBox "新しい行が追加されました！", vbInformation
End Sub



Sub GenerateRequestSheet()
    Dim ws As Worksheet
    Dim templateWs As Worksheet
    Dim targetCell As Range
    Dim targetRow As Long
    Dim btn As Object
    Dim newRequestSheet As Worksheet
    Dim sheetName As String
    Dim requestDate As Variant
    Dim customerName As Variant
    Dim totalAmount As Variant
    Dim rowContent As Variant, rowQty As Variant, rowUnitPrice As Variant, rowAmount As Variant
    Dim totalCalculated As Double ' totalCalculated の宣言を追加
    Dim linkAddress As String

    ' 現在のワークシートとテンプレートページを取得
    Set ws = ThisWorkbook.Sheets("list")
    Set templateWs = ThisWorkbook.Sheets("template")
    
    ' ボタンが押されたセルを取得
    Set targetCell = ActiveSheet.Buttons(Application.Caller).TopLeftCell
    targetRow = targetCell.Row
    
    ' 一覧ページ上のデータを取得
    requestDate = ws.Cells(targetRow, 1).Value ' 申請日
    customerName = ws.Cells(targetRow, 2).Value ' 名前
    dear = ws.Cells(targetRow, 3).Value ' 御中/様
    totalAmount = ws.Cells(targetRow, 4).Value ' 合計金額
    
    ' データ検証
    If Not IsValidField(requestDate) Or Not IsValidField(customerName) Then
        MsgBox "必須項目（申請日、名前）が入力され、データ形式が正しいことを確認してください。", vbCritical
        Exit Sub
    End If

    ' データ形式の検証と変換
    If Not IsDate(requestDate) Then
        MsgBox "申請日の形式が正しくありません。再入力してください。", vbCritical
        Exit Sub
    End If
    
    ' J、M、P、S、V、Y列の金額を合計し、税込金額をD列へ反映
    totalCalculated = 0

    totalCalculated = totalCalculated + SafeNumber(ws.Cells(targetRow, 10).Value) ' J列
    totalCalculated = totalCalculated + SafeNumber(ws.Cells(targetRow, 13).Value) ' M列
    totalCalculated = totalCalculated + SafeNumber(ws.Cells(targetRow, 16).Value) ' P列
    totalCalculated = totalCalculated + SafeNumber(ws.Cells(targetRow, 19).Value) ' S列
    totalCalculated = totalCalculated + SafeNumber(ws.Cells(targetRow, 22).Value) ' V列
    totalCalculated = totalCalculated + SafeNumber(ws.Cells(targetRow, 25).Value) ' Y列

    If totalCalculated <= 0 Then
        MsgBox "明細金額（J、M、P、S、V、Y列）を1つ以上入力してください。", vbExclamation
        Exit Sub
    End If

    totalCalculated = WorksheetFunction.Round(totalCalculated, 0)

    ws.Cells(targetRow, 4).Formula = BuildTotalAmountFormula(targetRow)
    totalAmount = WorksheetFunction.Round(totalCalculated * 1.1, 0)

    ' ワークシート名が既に存在するかを確認
    sheetName = customerName & "_" & Format(requestDate, "yyyymmdd") ' 御中と申請日を新しいシート名として使用
    
    On Error Resume Next
    Set newRequestSheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If Not newRequestSheet Is Nothing Then
        If MsgBox("シート " & sheetName & " は既に存在しています。置き換えますか？", vbYesNo + vbQuestion) = vbYes Then
            Application.DisplayAlerts = False
            newRequestSheet.Delete
            Application.DisplayAlerts = True
            Set newRequestSheet = Nothing
            DoEvents

            If WorksheetExists(sheetName) Then
                MsgBox "既存のシート「" & sheetName & "」を削除できませんでした。シート保護やブック保護を確認してください。", vbCritical
                Exit Sub
            End If
        Else
            MsgBox "申請書の生成をキャンセルしました。", vbInformation
            Exit Sub
        End If
    End If
    
    ' テンプレートから新しいシートを作成
    templateWs.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set newRequestSheet = ActiveSheet
    newRequestSheet.Name = sheetName ' 御中と申請日を新しいシート名として使用
    
    ' テンプレートに申請書データを入力
    newRequestSheet.Range("H2").ClearContents
    newRequestSheet.Range("F4").Value = requestDate
    newRequestSheet.Range("B5").Value = customerName
    newRequestSheet.Range("D5").Value = dear
    newRequestSheet.Range("D21").Value = "小計"
    newRequestSheet.Range("E21").Value = WorksheetFunction.Round(totalCalculated, 0)
    ClearMergedCell newRequestSheet.Range("B22")
    ClearMergedCell newRequestSheet.Range("D22")
    ClearMergedCell newRequestSheet.Range("E22")
    ClearMergedCell newRequestSheet.Range("I22")
    If ws.Cells(targetRow, 7).Value = 1 Then
        SetMergedCellValue newRequestSheet.Range("B22"), ws.Cells(targetRow, 26).Value ' list Z列
        SetMergedCellValue newRequestSheet.Range("D22"), "税理士報酬源泉所得税"
        SetMergedCellValue newRequestSheet.Range("I22"), ws.Cells(targetRow, 28).Value ' list AB列
        SetMergedCellFormula newRequestSheet.Range("E22"), "=ROUND(I22*10.21%,0)"
    End If
    ClearMergedCell newRequestSheet.Range("F25")
    ClearMergedCell newRequestSheet.Range("F26")
    ClearMergedCell newRequestSheet.Range("F27")
    newRequestSheet.Range("E25").Formula = "=ROUND(SUM(E21:F22),0)"
    newRequestSheet.Range("E26").Formula = "=ROUND(E21*0.1,0)"
    newRequestSheet.Range("E27").Formula = "=ROUND(SUM(E25:F26),0)"
    newRequestSheet.Range("C11").Formula = "=ROUND(E27,0)"

    ' テンプレートに残っている明細金額をクリア
    For i = 1 To 6
        newRequestSheet.Cells(14 + i, 5).MergeArea.ClearContents
    Next i

    
    ' 明細行データを入力（内容、期間、金額）
    For i = 1 To 6
        Dim perRowAmount As Variant
        Dim perRowContent As Variant
        Dim perRowPeriod As Variant
        Dim baseCol As Long

        baseCol = 8 + (i - 1) * 3 ' H/K/N/Q/T/W列
        perRowContent = ws.Cells(targetRow, baseCol).Value ' 内容
        perRowPeriod = ws.Cells(targetRow, baseCol + 1).Value ' 期間
        perRowAmount = ws.Cells(targetRow, baseCol + 2).Value ' 金額

        If SafeText(perRowContent) <> "" Or SafeText(perRowPeriod) <> "" Or SafeText(perRowAmount) <> "" Then
            newRequestSheet.Cells(14 + i, 2).Value = SafeText(perRowContent) ' 内容
            newRequestSheet.Cells(14 + i, 4).Value = SafeText(perRowPeriod) ' 期間
            If SafeText(perRowAmount) <> "" And SafeNumber(perRowAmount) <> 0 Then
                newRequestSheet.Cells(14 + i, 5).Value = WorksheetFunction.Round(SafeNumber(perRowAmount), 0) ' 金額
            Else
                newRequestSheet.Cells(14 + i, 5).MergeArea.ClearContents
            End If
        End If
    Next i
    
    ' リンクのアドレスを作成
    linkAddress = ws.Name & "!A" & targetRow

    ' J1セルにリンクを設定
    With newRequestSheet.Hyperlinks.Add(Anchor:=newRequestSheet.Cells(1, 10), _
                                     Address:="", _
                                     SubAddress:=linkAddress, _
                                     TextToDisplay:="listに戻る")
        ' ハイパーリンクのフォーマットを設定
        .Range.Font.Color = RGB(0, 0, 255)
        .Range.Font.Underline = True
    End With

    MsgBox "申請書が生成されました！", vbInformation
End Sub

' フィールド検証
Function IsValidField(fieldValue As Variant) As Boolean
    ' フィールドが有効かをチェック
    If IsEmpty(fieldValue) Then
        IsValidField = False
    ElseIf VarType(fieldValue) = vbString And Trim(fieldValue) = "" Then
        IsValidField = False
    ElseIf VarType(fieldValue) = vbDouble And fieldValue = 0 Then
        IsValidField = False
    ElseIf VarType(fieldValue) = vbDate And Not IsDate(fieldValue) Then
        IsValidField = False
    Else
        IsValidField = True
    End If
End Function

Function SafeText(fieldValue As Variant) As String
    If IsError(fieldValue) Or IsEmpty(fieldValue) Then
        SafeText = ""
    Else
        SafeText = CStr(fieldValue)
    End If
End Function

Function BuildTotalAmountFormula(targetRow As Long) As String
    Dim amountSum As String

    amountSum = "SUM(" & _
                BuildAmountCellFormula("J", targetRow) & "," & _
                BuildAmountCellFormula("M", targetRow) & "," & _
                BuildAmountCellFormula("P", targetRow) & "," & _
                BuildAmountCellFormula("S", targetRow) & "," & _
                BuildAmountCellFormula("V", targetRow) & "," & _
                BuildAmountCellFormula("Y", targetRow) & ")"

    BuildTotalAmountFormula = "=IF(" & amountSum & "=0,"""",ROUND(" & amountSum & "*1.1+IF(AND(G" & targetRow & "=1,ISNUMBER(AB" & targetRow & ")),AB" & targetRow & "*10.21%,0),0))"
End Function

Function BuildAmountCellFormula(columnLetter As String, targetRow As Long) As String
    Dim address As String

    address = columnLetter & targetRow
    BuildAmountCellFormula = "IFERROR(VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(" & address & ",""－"",""-""),""−"",""-""),"","","""")),0)"
End Function

Function SafeNumber(fieldValue As Variant) As Double
    Dim textValue As String

    If IsError(fieldValue) Or IsEmpty(fieldValue) Then
        SafeNumber = 0
        Exit Function
    End If

    If IsNumeric(fieldValue) Then
        SafeNumber = CDbl(fieldValue)
        Exit Function
    End If

    textValue = CStr(fieldValue)
    textValue = Replace(textValue, ",", "")
    textValue = Replace(textValue, "，", "")
    textValue = Replace(textValue, "¥", "")
    textValue = Replace(textValue, "￥", "")
    textValue = Replace(textValue, "円", "")
    textValue = Replace(textValue, " ", "")
    textValue = Replace(textValue, ChrW(12288), "")
    textValue = Replace(textValue, "－", "-")
    textValue = Replace(textValue, "−", "-")

    If IsNumeric(textValue) Then
        SafeNumber = CDbl(textValue)
    Else
        SafeNumber = 0
    End If
End Function

Function WorksheetExists(sheetName As String) As Boolean
    Dim checkWs As Worksheet

    On Error Resume Next
    Set checkWs = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    WorksheetExists = Not checkWs Is Nothing
End Function

Sub ClearMergedCell(target As Range)
    target.MergeArea.ClearContents
End Sub

Sub SetMergedCellValue(target As Range, value As Variant)
    target.MergeArea.Cells(1, 1).Value = value
End Sub

Sub SetMergedCellFormula(target As Range, formulaText As String)
    target.MergeArea.Cells(1, 1).Formula = formulaText
End Sub

Sub PrintRequestSheet()
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim targetRow As Long
    Dim requestSheetName As String
    Dim customerName As String
    Dim requestDate As Variant
    Dim targetSheet As Worksheet

    ' 押されたボタンのセルがある行を取得
    Set ws = ThisWorkbook.Sheets("list")
    Set targetCell = ws.Buttons(Application.Caller).TopLeftCell
    targetRow = targetCell.Row

    ' 名前と申請日を取得し、申請書の名前を生成
    customerName = ws.Cells(targetRow, 2).Value ' 名前（御中/様）
    requestDate = ws.Cells(targetRow, 1).Value ' 申請日
    If Not IsDate(requestDate) Then
        MsgBox "申請日の形式が正しくありません。", vbExclamation
        Exit Sub
    End If
    requestSheetName = customerName & "_" & Format(CDate(requestDate), "yyyymmdd")

    ' 対応する申請書のシートが存在するか確認
    On Error Resume Next
    Set targetSheet = ThisWorkbook.Sheets(requestSheetName)
    On Error GoTo 0

    If targetSheet Is Nothing Then
        MsgBox "申請書「" & requestSheetName & "」が存在しません。先に作成してください！", vbExclamation
    Else
        ' 申請書を印刷
        On Error Resume Next
        targetSheet.PrintPreview
        On Error GoTo 0
    End If
End Sub

