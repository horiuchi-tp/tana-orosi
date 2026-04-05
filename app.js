Option Explicit

' =====================================================================================
' マクロ1：棚卸データ取り込み（自動判定統合版）
' 概要: スマホから送信されたモード（全体棚卸 か 列変更）を自動で読み取り、処理を分岐する
' =====================================================================================
Sub 棚卸データ取り込み()
    Dim ws As Worksheet
    Dim http As Object
    Dim apiUrl As String
    Dim responseText As String
    Dim spDataRows() As String, spDataCols() As String
    Dim i As Long, lastRow As Long, targetRow As Long
    Dim dictMaster As Object, dictDuplicate As Object, dictProcessed As Object
    Dim key As String
    Dim orderNo As String, customerName As String, rowNo As String
    Dim areaVal As String, qtyVal As String, excelQty As String, modeVal As String
    Dim numQtyVal As Long, numExcelQty As Long
    Dim newTargetRow As Long, updateCount As Long
    Dim isYellow As Boolean
    Dim bxMsg As String
    Dim k As Variant
    
    ' 指定された新しいGASのURL [cite: 24]
    apiUrl = "https://script.google.com/macros/s/AKfycbyo1t1sX5GyrCx22yfJOFNUJv6CesapJ7xGoFk947IDFF01glOPJLU5S3X3bizQE3tYBw/exec"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("日延べ")
    On Error GoTo ErrorHandler
    If ws Is Nothing Then
        MsgBox "「日延べ」シートが見つかりません。", vbCritical [cite: 25]
        Exit Sub
    End If
    
    ' --- 処理の高速化（画面更新等の停止） --- [cite: 25]
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' 辞書オブジェクトの初期化 [cite: 25]
    Set dictMaster = CreateObject("Scripting.Dictionary")
    Set dictDuplicate = CreateObject("Scripting.Dictionary")
    Set dictProcessed = CreateObject("Scripting.Dictionary")
    
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    
    ' 1. マスターデータのインデックス作成（7行目以降、黄色背景スキップ） [cite: 25, 26]
    For i = 7 To lastRow
        isYellow = False
        ' A列の背景色が黄色(vbYellow または ColorIndex=6)か判定 [cite: 26]
        If ws.Cells(i, 1).Interior.Color = vbYellow Or ws.Cells(i, 1).Interior.ColorIndex = 6 Then
            isYellow = True
        End If
        
        If Not isYellow Then
            ' キー: [受注番号_行番号] [cite: 26]
            key = Trim(CStr(ws.Cells(i, 12).Value)) & "_" & Trim(CStr(ws.Cells(i, 13).Value))
            
            If Not dictMaster.Exists(key) Then
                dictMaster.Add key, i [cite: 27]
            Else
                ' 無色行の重複（ダブリ）を記録 [cite: 27]
                If Not dictDuplicate.Exists(key) Then
                    dictDuplicate.Add key, CStr(i)
                Else
                    dictDuplicate(key) = dictDuplicate(key) & "," & CStr(i) [cite: 28]
                End If
            End If
        End If
    Next i
    
    ' 2. GAS APIから最新データを取得 [cite: 28]
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", apiUrl, False
    http.send
    
    If http.Status <> 200 Then
        Err.Raise 1000, , "通信エラー (Status: " & http.Status & ")" [cite: 29]
    End If
    
    responseText = http.responseText
    
    ' HTMLエラー検知（フェイルセーフ） [cite: 29]
    If InStr(1, responseText, "<html", vbTextCompare) > 0 Or InStr(1, responseText, "<!DOCTYPE", vbTextCompare) > 0 Then
        MsgBox "システムエラー：エラーページを受信しました。GASのデプロイを確認してください。", vbCritical
        GoTo CleanUp
    End If
    
    If Trim(responseText) = "" Then [cite: 30]
        MsgBox "スプレッドシートに取り込むデータがありません。", vbInformation
        GoTo CleanUp
    End If
    
    spDataRows = Split(responseText, vbLf)
    
    ' 3. モードの自動判定（受信データの6列目を確認） [cite: 30, 31]
    modeVal = "full" 
    For i = LBound(spDataRows) To UBound(spDataRows)
        If Trim(spDataRows(i)) <> "" Then
            spDataCols = Split(spDataRows(i), ",")
            If UBound(spDataCols) >= 5 Then modeVal = Trim(spDataCols(5)) [cite: 31]
            Exit For
        End If
    Next i
    
    ' 4. 処理の分岐 [cite: 31]
    If modeVal = "area_only" Then
        ' ---------------------------------------------------------
        ' 【列変更（エリア更新）モード】
        ' --------------------------------------------------------- [cite: 31]
        If MsgBox("【列変更モード】のデータを受信しました。" & vbCrLf & "エリア情報のみを更新します。実行しますか？", vbQuestion + vbYesNo) = vbNo Then GoTo CleanUp
        updateCount = 0 [cite: 31]
        
        For i = LBound(spDataRows) To UBound(spDataRows) [cite: 32]
            If Trim(spDataRows(i)) <> "" Then
                spDataCols = Split(spDataRows(i), ",")
                If UBound(spDataCols) >= 4 Then
                    orderNo = Trim(spDataCols(0)): rowNo = Trim(spDataCols(2)): areaVal = Trim(spDataCols(3)) [cite: 32]
                    key = orderNo & "_" & rowNo [cite: 33]
                    
                    ' 一致した場合のみBW列(75)とBY列(77)を変更
                    If dictMaster.Exists(key) Then
                        targetRow = dictMaster(key)
                        ws.Cells(targetRow, 75).Value = areaVal ' BW列 [cite: 33]
                        ws.Cells(targetRow, 77).Value = Format(Now, "yyyy/mm/dd hh:mm:ss") ' BY列（時刻） [cite: 34]
                        updateCount = updateCount + 1
                    End If
                End If
            End If [cite: 35]
        Next i
        MsgBox "列（エリア）の変更が完了しました。" & vbCrLf & updateCount & "件のデータを更新しました。", vbInformation [cite: 35]

    Else
        ' ---------------------------------------------------------
        ' 【全体棚卸モード】
        ' --------------------------------------------------------- [cite: 35]
        If MsgBox("【全体棚卸モード】のデータを受信しました。" & vbCrLf & "照合と未発見(QRに無し)のあぶり出しを行います。実行しますか？", vbQuestion + vbYesNo) = vbNo Then GoTo CleanUp
        
        newTargetRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1 [cite: 35]
        If newTargetRow < 7 Then newTargetRow = 7
        
        For i = LBound(spDataRows) To UBound(spDataRows) [cite: 36]
            If Trim(spDataRows(i)) <> "" Then
                spDataCols = Split(spDataRows(i), ",")
                If UBound(spDataCols) >= 4 Then
                    orderNo = Trim(spDataCols(0)): customerName = Trim(spDataCols(1)): rowNo = Trim(spDataCols(2)) [cite: 36]
                    areaVal = Trim(spDataCols(3)): qtyVal = Trim(spDataCols(4)) [cite: 37]
                    key = orderNo & "_" & rowNo
                    
                    If Not dictProcessed.Exists(key) Then dictProcessed.Add key, "" [cite: 37]
                    
                    If dictMaster.Exists(key) Then [cite: 37]
                        targetRow = dictMaster(key) [cite: 38]
                        ws.Cells(targetRow, 74).Value = "ある" ' BV列
                        ws.Cells(targetRow, 75).Value = areaVal ' BW列 [cite: 38]
                        
                        ' 数量の数値比較 [cite: 38]
                        If IsNumeric(qtyVal) Then numQtyVal = CLng(qtyVal) Else numQtyVal = 0 [cite: 39]
                        excelQty = Trim(CStr(ws.Cells(targetRow, 23).Value))
                        If IsNumeric(excelQty) Then numExcelQty = CLng(excelQty) Else numExcelQty = 0
                        
                        bxMsg = "" [cite: 39]
                        If numQtyVal > numExcelQty Then bxMsg = "数量多い" ElseIf numQtyVal < numExcelQty Then bxMsg = "数量少い" [cite: 40]
                        
                        ' ダブリ情報の結合 [cite: 40]
                        If dictDuplicate.Exists(key) Then
                            If bxMsg <> "" Then bxMsg = bxMsg & "（" & dictDuplicate(key) & "行とダブリあり）" Else bxMsg = dictDuplicate(key) & "行とダブリあり" [cite: 41]
                        End If
                        
                        ws.Cells(targetRow, 76).Value = bxMsg ' BX列 [cite: 41]
                        ws.Cells(targetRow, 77).Value = Format(Now, "yyyy/mm/dd hh:mm:ss") ' BY列 [cite: 41]
                        
                    Else [cite: 42]
                        ' マスターにない新規データの追加
                        targetRow = newTargetRow
                        ws.Cells(targetRow, 12).Value = orderNo: ws.Cells(targetRow, 13).Value = rowNo [cite: 42]
                        ws.Cells(targetRow, 18).Value = customerName: ws.Cells(targetRow, 74).Value = "データに無し" [cite: 43]
                        ws.Cells(targetRow, 77).Value = Format(Now, "yyyy/mm/dd hh:mm:ss")
                        
                        dictMaster.Add key, targetRow [cite: 43]
                        newTargetRow = newTargetRow + 1 [cite: 44]
                    End If
                End If
            End If
        Next i
        
        ' QRになかったマスターデータ（無色行）に「QRに無し」を記録 [cite: 44]
        For Each k In dictMaster.keys
            If Not dictProcessed.Exists(k) Then [cite: 45]
                targetRow = dictMaster(k)
                ws.Cells(targetRow, 77).Value = "QRに無し" ' BY列 [cite: 45]
            End If
        Next k
        
        MsgBox "全体棚卸の照合と反映が完了しました。", vbInformation [cite: 45]
    End If
    GoTo CleanUp

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical [cite: 45]

CleanUp:
    ' --- 高速化の解除 --- [cite: 45]
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True [cite: 46]
    Application.ScreenUpdating = True
    Set dictMaster = Nothing: Set dictDuplicate = Nothing: Set dictProcessed = Nothing: Set http = Nothing [cite: 46]
End Sub

' =====================================================================================
' マクロ2：棚卸結果リセット（次月準備用） [cite: 46]
' 概要: 判定列(BV〜BY)をすべて消去する
' =====================================================================================
Sub 棚卸結果リセット()
    Dim ws As Worksheet
    Dim lastRow As Long
    If MsgBox("棚卸結果（BV〜BY列）をすべてリセットしますか？" & vbCrLf & "※この操作は元に戻せません。", vbExclamation + vbYesNo) = vbNo Then Exit Sub
    
    Set ws = ThisWorkbook.Sheets("日延べ")
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    If lastRow >= 7 Then
        ws.Range(ws.Cells(7, 74), ws.Cells(lastRow, 77)).ClearContents
    End If
    MsgBox "リセットが完了しました。新しい全体棚卸を開始できます。", vbInformation
End Sub