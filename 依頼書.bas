Option Explicit

Public r As String
Public Newbtn(1 To 30) As New Class1
Public branch As String  'branchは文字と定義

Sub Main()
Call changesheetname("依頼書コピー")
End Sub

Public Sub changesheetname(bookname As String)
Dim ws As Worksheet, flag As Boolean
Dim RC As Integer
label1100:
flag = False

For Each ws In Worksheets
If ws.Name = bookname Then flag = True
Next ws
If flag = True Then
RC = MsgBox(bookname & "が存在します。書き換えますか？", vbYesNo + vbQuestion, "確認")
If RC = vbYes Then
Application.DisplayAlerts = False
Worksheets(bookname).Delete
Application.DisplayAlerts = True
Else
bookname = InputBox("名前を入力")
GoTo Labe1100
End If
ActiveSheet.Name = bookname
End If
End Sub

Sub Macro1()
myNumber = Worksheets.Count
Sheets("依頼書").Copy after:=Worksheets("依頼書")
ActiveSheet.Name = Worksheets("依頼書").Range("A5").Value & "へ" & Worksheets("依頼書").Range("B12").Value & myNumber
End Sub


Sub 出力()

    Dim i As String


    Worksheets("依頼書").Copy after:=Worksheets("依頼書")
    If Worksheets("依頼書").Range("A5").Value = "" Then
        ActiveSheet.Name = "某物件"

    Else: ActiveSheet.Name = Worksheets("依頼書").Range("A5").Value & "へ" & Worksheets("依頼書").Range("B12").Value


    End If


End Sub

Sub 明細クリア() '下記セルの値をクリアする

   Worksheets("依頼書").Range("A5:A6").ClearContents
   Worksheets("依頼書").Range("A10:A11").ClearContents
   Worksheets("依頼書").Range("B12").MergeArea.ClearContents

   Worksheets("依頼書").Range("A13:F33").ClearContents
   Worksheets("依頼書").Range("H23:H24").ClearContents
   Worksheets("依頼書").Range("G22").ClearContents

End Sub

Sub 印刷()

    ActiveWindow.SelectedSheets.PrintOut ActivePrinter:="FUJI XEROX ApeosPort-VII C5573"
End Sub
Sub ドキュワークスに印刷()

    ActiveWindow.SelectedSheets.PrintOut ActivePrinter:="DocuWorks Printer"
End Sub
Sub 印刷ドキュワーク()

  Dim resultMessage As String
  resultMessage = test

  If resultMessage <> "" Then
    MsgBox resultMessage, vbCritical
  Else
    MsgBox "処理成功", vbInformation
  End If

End Sub
Function test() As String
On Error GoTo Test_Err
  test = ""
    ActiveWindow.SelectedSheets.PrintOut ActivePrinter:="FUJI XEROX ApeosPort-VII C5573 FAX"
Exit Function
Test_Err:
  'エラー時にエラー情報を返す
  test = "【処理エラー】" & vbCrLf & _
         "エラー番号：" & Err.Number & vbCrLf & _
         "エラーメッセージ：" & Err.Description

End Function


Sub 営業所送付先()
Dim i, cnt, le As Integer 'iは整数と定義
    Dim ctrl As Control 'ctrlをControlと定義
    Dim Newbtn(1 To 30) As New Class4
    cnt = 0
    le = 0
    With Worksheets("送付先リスト")
        For i = 2 To 16
                branch = .Cells(i, 3)
                With UserForm31
                    Set ctrl = .Controls.Add("Forms.CommandButton.1", "Bnt" & cnt) '今からユーザーフォームにボタンを追加します

                    If cnt >= 5 And cnt <= 9 Then
                        le = 1
                    ElseIf cnt >= 10 And cnt <= 14 Then
                        le = 2
                    ElseIf cnt >= 15 And cnt <= 19 Then
                        le = 3
                    ElseIf cnt >= 20 And cnt <= 24 Then
                        le = 4
                    End If

                    ctrl.Height = 65
                    ctrl.Top = (cnt - le * 5) * 66
                    ctrl.Left = le * 200
                    ctrl.Width = 200
                    ctrl.Caption = branch
                    ctrl.FontName = "メイリオ"
                    ctrl.FontBold = True
                    ctrl.FontSize = 12
                    ctrl.ForeColor = &HFFFFFF
                    ctrl.BackColor = &H404040
                End With
                cnt = cnt + 1
                Set Newbtn(cnt).button = ctrl
        Next
    End With

    If le = 1 Then
        UserForm31.Width = 405
    ElseIf le = 2 Then
        UserForm31.Width = 610
    ElseIf le = 3 Then
        UserForm31.Width = 815
    ElseIf le = 4 Then
        UserForm31.Width = 1020
    Else
        UserForm31.Width = 200
    End If

    UserForm31.Show

For i = 2 To 30
If Worksheets("依頼書").Range("A27").Value = Worksheets("送付先リスト").Cells(i, 3).Value Then
Worksheets("依頼書").Range("A24").Value = "送付先"
Worksheets("依頼書").Range("A25").Value = Worksheets("送付先リスト").Cells(i, 4).Value
Worksheets("依頼書").Range("A26").Value = Worksheets("送付先リスト").Cells(i, 5).Value
Worksheets("依頼書").Range("A28").Value = "TEL　" & Worksheets("送付先リスト").Cells(i, 6).Value
End If
Next

End Sub

Sub 定価テンプレ()
Worksheets("依頼書").Range("A17") = "下記商品の定価、仕切、運賃を教えて下さい。"
End Sub

Sub 樹木在庫確認テンプレ()
Worksheets("依頼書").Range("A17") = "下記樹種の見積と在庫の有無を教えて下さい。"
End Sub

Sub 商品在庫確認テンプレ()
Worksheets("依頼書").Range("A17") = "下記商品の見積と在庫の有無を教えて下さい。"
End Sub

Sub セル結合()
Worksheets("依頼書").Range("A13:F33").Merge
End Sub


Sub Macro2()
'
' Macro2 Macro
'

'
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "C:\Users\RK0141120\Desktop\依頼書作成くん.pdf", Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        False
    ActiveWindow.SmallScroll Down:=24
    Range("H30").Select
End Sub
Sub Macro4()
'
' Macro4 Macro

    Sheets("原紙").Select
    Range("A2:F36").Select
    Selection.Copy
    Sheets("依頼書").Select
    Range("A2:F36").Select
    ActiveSheet.Paste

    Worksheets("依頼書").Range("H23:H24").ClearContents
   Worksheets("依頼書").Range("G22").ClearContents

End Sub





Public Sub FAXCopy()

    Dim oTextBox As Object
    Dim sText As String

    'TextBox作成　Forms.TextBox.1　は ClassID
    Set oTextBox = CreateObject("Forms.TextBox.1")

    sText = Worksheets("依頼書").Range("H24").Value


    With oTextBox

       .MultiLine = True
       .Text = sText             'TextBoxへコピーする文字列設定
       .SelStart = 0
       .SelLength = .TextLength
       .Copy

    End With

    Set oTextBox = Nothing

End Sub

Public Sub TELCopy()

    Dim oTextBox As Object
    Dim sText As String

    'TextBox作成　Forms.TextBox.1　は ClassID
    Set oTextBox = CreateObject("Forms.TextBox.1")

    sText = Worksheets("依頼書").Range("H23").Value


    With oTextBox

       .MultiLine = True
       .Text = sText             'TextBoxへコピーする文字列設定
       .SelStart = 0
       .SelLength = .TextLength
       .Copy

    End With

    Set oTextBox = Nothing

End Sub

Private Sub button_click()
Worksheets("依頼書").Range(r) = button.Caption 'ボタンの題名を取得
UserForm3.Hide

End Sub


Private Sub button_click()
Worksheets("依頼書").Range(r) = button.Caption 'ボタンの題名を取得
UserForm4.Hide
End Sub

Private Sub button_click()
Worksheets("依頼書").Range(r) = button.Caption 'ボタンの題名を取得

Unload UserForm17

End Sub

Private Sub button_click()
Worksheets("依頼書").Range("A27") = button.Caption 'ボタンの題名を取得

UserForm31.Hide
End Sub


Private Sub button_click()
Worksheets("依頼書").Range("C10") = button.Caption 'ボタンの題名を取得

UserForm33.Hide

End Sub