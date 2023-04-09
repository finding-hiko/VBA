Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

  If KeyCode <> vbKeyReturn Then Exit Sub 'Enter以外は、処理を終了

  ListBox1.Clear 'リストボックスをクリア

  'データベースを検索
  For i = 2 To Sheets("住所録").Cells(Rows.Count, "A").End(xlUp).Row
    '部分一致で商品を検索
    If InStr(Sheets("住所録").Cells(i, "A"), TextBox1.Text) > 0 Then
      ListBox1.AddItem Sheets("住所録").Cells(i, "A") 'リストボックスに値を追加
    End If
  Next

  'リストボックスにデータがある場合
  If ListBox1.ListCount > 0 Then
    ListBox1.SetFocus 'リストボックスをフォーカス
    ListBox1.ListIndex = 0 '一番上を選択
  Else
    KeyCode = 0 'テキストボックスをフォーカスしたままにする
    MsgBox "データがありません"
  End If

End Sub
Sub ボタン1_Click()
Userform1.Show

End Sub

Sub 印刷()

    ActiveWindow.SelectedSheets.PrintOut ActivePrinter:="FUJI XEROX ApeosPort-VII C5573"
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

  If KeyCode <> vbKeyReturn Then Exit Sub 'Enter以外は、処理を終了

  ListBox1.Clear 'リストボックスをクリア

  'データベースを検索
  For i = 2 To Sheets("住所録").Cells(Rows.Count, "A").End(xlUp).Row
    '部分一致で商品を検索
    If InStr(Sheets("住所録").Cells(i, "A"), TextBox1.Text) > 0 Then
      ListBox1.AddItem Sheets("住所録").Cells(i, "A") 'リストボックスに値を追加
    End If
  Next

  'リストボックスにデータがある場合
  If ListBox1.ListCount > 0 Then
    ListBox1.SetFocus 'リストボックスをフォーカス
    ListBox1.ListIndex = 0 '一番上を選択
  Else
    KeyCode = 0 'テキストボックスをフォーカスしたままにする
    MsgBox "データがありません"
  End If

End Sub



Sub 郵便番号()
    Dim shtMain As Worksheet
    Dim shtData As Worksheet
    Dim varZipCd As Variant
    Dim varAddress As Variant
    Dim lastRow As Long
    Dim nowRow As Long
    Dim i As Long


    '②「データ」シートを変数に格納する
    Set shtData = WorkboSheets("データ")

    '③「住所検索」シートの初期化
    TextBox2.Text = ""

    '④郵便番号データの最終行を取得する
    lastRow = shtData.Cells(shtData.Rows.Count, 3).End(xlUp).Row

    '⑤郵便番号データを配列に格納する
    varZipCd = shtData.Range("C1:C" & lastRow)

    '⑥住所データを配列に格納する
    varAddress = shtData.Range("G1:I" & lastRow)

    nowRow = 1

    Do While True

        '⑦「住所検索」シートの現在行を次の行に変更する
        nowRow = nowRow + 1

        '⑧郵便番号が空なら処理を抜ける
        If TextBox3.Value = "" Then
            Exit Do
        End If

        '⑨郵便番号で「データ」シートの郵便番号を探し、該当行の住所を取得する
        For i = 1 To UBound(varZipCd)
            If TextBox2.Value = varZipCd(i, 1) Then
                 TextBox3.Value = varAddress(i, 1) & varAddress(i, 2) & varAddress(i, 3)
                Exit For
            End If
        Next
    Loop

    MsgBox "完了"

End Sub


Sub 明細クリア() '下記セルの値をクリアする

   Application.ScreenUpdating = False

   Worksheets("入力画面").Range("A3").MergeArea.ClearContents
   Worksheets("入力画面").Range("A4").ClearContents
   Worksheets("入力画面").Range("A5").MergeArea.ClearContents
   Worksheets("入力画面").Range("A6:A8").ClearContents
   Worksheets("入力画面").Range("B21").ClearContents

   Worksheets("入力画面").Range("G3").MergeArea.ClearContents
   Worksheets("入力画面").Range("G4").ClearContents
   Worksheets("入力画面").Range("G5").MergeArea.ClearContents
   Worksheets("入力画面").Range("G6:G8").ClearContents
   Worksheets("入力画面").Range("H21").ClearContents

   Worksheets("入力画面").Range("A25").MergeArea.ClearContents
   Worksheets("入力画面").Range("A26").ClearContents
   Worksheets("入力画面").Range("A27").MergeArea.ClearContents
   Worksheets("入力画面").Range("A28:A30").ClearContents
    Worksheets("入力画面").Range("B43").ClearContents

   Worksheets("入力画面").Range("G25").MergeArea.ClearContents
   Worksheets("入力画面").Range("G26").ClearContents
   Worksheets("入力画面").Range("G27").MergeArea.ClearContents
   Worksheets("入力画面").Range("G28:G30").ClearContents
    Worksheets("入力画面").Range("H43").ClearContents

   Application.ScreenUpdating = True
   End Sub


Sub 印刷ダイアログ()
    Application.Dialogs(xlDialogPrint).Show
End Sub