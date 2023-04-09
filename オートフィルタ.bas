Attribute VB_Name = "Module1"
Option Explicit
Sub 文字列での絞り込み抽出()
  If Range("C2").Value = "" Then
    MsgBox "商品名を入力してください(あいまい検索)。"
    Exit Sub
  Else
    Worksheets("抽出画面").Range("A1").AutoFilter Field:=17, Criteria1:="*" & Worksheets("入力画面").Range("C2").Value & "*"
    Call 列抽出
    Call セル交互色
  End If
End Sub

Sub 規格での絞り込み抽出()
    Dim x As String
  If Range("C5").Value = "" Then
    MsgBox "規格を入力してください(あいまい検索)。"
    Exit Sub
  Else
    x = StrConv(Range("C5").Value, vbNarrow)
    Worksheets("抽出画面").Range("A1").AutoFilter Field:=18, Criteria1:="*" & x & "*"
    Worksheets("入力画面").Range(Range("B15"), Range("B" & Cells.Rows.Count)).EntireRow.Clear
    Call 列抽出
    Call セル交互色
  End If
End Sub

Sub 仕入先での絞り込み抽出()
  If Range("C7").Value = "" Then
    MsgBox "仕入先名を入力してください(あいまい検索)。"
    Exit Sub
  Else
    Worksheets("抽出画面").Range("A1").AutoFilter Field:=8, Criteria1:="*" & Worksheets("入力画面").Range("C7").Value & "*"
     Worksheets("入力画面").Range(Range("B15"), Range("B" & Cells.Rows.Count)).EntireRow.Clear
    Call 列抽出
    Call セル交互色
  End If
End Sub
Sub 販売先での絞り込み抽出()
  If Range("C9").Value = "" Then
    MsgBox "販売先名を入力してください(あいまい検索)。"
    Exit Sub
  Else
    Worksheets("抽出画面").Range("A1").AutoFilter Field:=10, Criteria1:="*" & Worksheets("入力画面").Range("C9").Value & "*"
    Worksheets("入力画面").Range(Range("B15"), Range("B" & Cells.Rows.Count)).EntireRow.Clear
    Call 列抽出
    Call セル交互色
  End If
End Sub
Sub 複数検索での絞り込み抽出()
  If Range("C2").Value = "" And Range("C3").Value = "" Then
  MsgBox "商品名を入力して下さい"
  Exit Sub

  Else
    Worksheets("抽出画面").Range("A1").AutoFilter Field:=17, Criteria1:="*" & Worksheets("入力画面").Range("C2").Value & "*", _
    Operator:=xlAnd, Criteria2:="*" & Worksheets("入力画面").Range("C3").Value & "*"
    Worksheets("入力画面").Range(Range("B15"), Range("B" & Cells.Rows.Count)).EntireRow.Clear
    Call 列抽出
    Call セル交互色
  End If

End Sub
Sub オートフィルターの解除()
  Worksheets("抽出画面").Range("A1").AutoFilter
  Worksheets("入力画面").Range("C2").Value = ""
  Worksheets("入力画面").Range("C3").Value = ""
  Worksheets("入力画面").Range("C5").Value = ""
  Worksheets("入力画面").Range("C7").Value = ""
  Worksheets("入力画面").Range(Range("B15"), Range("B" & Cells.Rows.Count)).EntireRow.Clear
End Sub

Sub 列抽出()
Dim データ範囲 As Range
Dim 抽出列 As Variant
Dim i As Long

Set データ範囲 = Sheets("抽出画面").Range("A1").CurrentRegion
抽出列 = Array(2, 8, 10, 14, 15, 17, 18, 19, 20, 21, 22, 23)
For i = 0 To UBound(抽出列)
データ範囲.Columns(抽出列(i)).Copy Sheets("入力画面").Range("B15").Offset(0, i)
Next i
End Sub

Sub セル交互色()
  Dim i As Long
  Dim 最終行 As Long
  Dim 最終列 As Long
   最終行 = Cells(Rows.Count, 2).End(xlUp).Row
   最終列 = Cells(15, Columns.Count).End(xlToLeft).Column
Application.ScreenUpdating = False
  For i = 15 To 最終行 Step 2
    Worksheets("入力画面").Range(Cells(i, 2), Cells(i, 最終列)).Interior.Color = RGB(192, 192, 192)
    Worksheets("入力画面").Range(Cells(i, 2), Cells(i, 最終列)).Borders.LineStyle = True

  Next i
  Worksheets("入力画面").Range(Cells(15, 2), Cells(最終行, 最終列)).Borders.LineStyle = True
Application.ScreenUpdating = True
End Sub