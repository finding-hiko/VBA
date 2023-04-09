Option Explicit
Public WithEvents button As CommandButton
Private Sub button_click()
Worksheets("請求書").Range("A5") = button.Caption 'ボタンの題名を取得
myForm2.Hide
End Sub

Sub 請求書作成()
    Dim cnt, cnt2, j As Long
    Dim rw, x As Long
    Dim Last_Row As Double



    Worksheets("売上").Range("O1").AutoFilter  'Oのセルの納品日を日付順に並び替える。
    Worksheets("売上").AutoFilter.Sort.SortFields.Clear
    Worksheets("売上").AutoFilter.Sort.SortFields.Add Key:=Range( _
        "O1:O10000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("売上").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Worksheets("売上").Range("O1").AutoFilter


        rw = 37
        For cnt = 1 To 50
            Worksheets("請求書").Range("H3").Value = Date 'H3のセルに日付を入力する
            Worksheets("請求書").Range("H" & rw).Value = Date  'H*rwのセルに日付を入力する
            rw = rw + 33
        Next

        rw = 41
        For cnt = 1 To 50
            Worksheets("請求書").Range("A14:H30").ClearContents  'A14~H28のセルの値を消去する
            Worksheets("請求書").Range("A" & rw, "H" & rw + 22).ClearContents  'A*rw~H*rw+10のセルの値を消去する
            rw = rw + 33
        Next

        rw = 14
        For cnt = 2 To 9000

            If Worksheets("売上").Range("J" & cnt).Value = Worksheets("請求書").Range("A5").Value Then
                Worksheets("請求書").Range("A" & rw).Value = Worksheets("売上").Range("O" & cnt).Value
                Worksheets("請求書").Range("B" & rw).Value = Worksheets("売上").Range("Q" & cnt).Value
                Worksheets("請求書").Range("C" & rw).Value = Worksheets("売上").Range("R" & cnt).Value
                Worksheets("請求書").Range("D" & rw).Value = Worksheets("売上").Range("S" & cnt).Value
                Worksheets("請求書").Range("E" & rw).Value = Worksheets("売上").Range("T" & cnt).Value
                Worksheets("請求書").Range("G" & rw).Value = Worksheets("売上").Range("U" & cnt).Value
                Worksheets("請求書").Range("F" & rw).Value = Int(Worksheets("請求書").Range("G" & rw).Value / Worksheets("請求書").Range("D" & rw).Value)
                Worksheets("請求書").Range("H" & rw).Value = Worksheets("売上").Range("X" & cnt).Value
                rw = rw + 1
                If (rw - 31) Mod 33 = 0 Then 'そのページの最後の行にいったら次のページの行へ移動する
                    rw = rw + 10
                End If

            End If
        Next
        rw = 64
        x = Worksheets("請求書").Range("G31").Value  '各ページの小計を足して、B10セルへ合計として出力する
        For cnt = 1 To 50
            x = Worksheets("請求書").Range("G" & rw).Value + x
            rw = rw + 33
        Next
            Worksheets("請求書").Range("B10").Value = x

End Sub
'

Sub PDF化()

    Dim fileName As String  'そのページの明細に値がある場合、そのページをPDF化する
    Dim pdf_name As String
    Dim x, cnt As Long
        x = 1
    For cnt = 0 To 49
        If Worksheets("請求書").Range("A" & 1625 - cnt * 33) <> "" Then
            x = x + 1
        End If
    Next

    Dim WSH As Variant
        Set WSH = CreateObject("WScript.Shell")
        pdf_name = WSH.SpecialFolders("Desktop") & "\VBA見積書\【" & Format(Date, "yyyymmdd") & "】 " & Worksheets("請求書").Range("A5") & "様" & ".pdf"

    Worksheets("請求書").ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdf_name, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, From:=1, To:=x, _
        OpenAfterPublish:=False

End Sub

Public customers As String
Sub 顧客名検索()

    customers = StrConv(InputBox("キーワードを入力してください"), vbWide)
    Dim Newbtn(1 To 50) As New Class1
    Dim myDic, myDic2 As Object, myKey, myKey2 As Variant 'myDicをobjectとし、myKeyはvariant(何にでもなれる変数,反則技)とする
    Dim c, d As Variant, varDate  As Variant 'c,vardateをvariantとする
        Set myDic = CreateObject("Scripting.Dictionary")  'myDicの意味をCreateObject("Scripting.Dictionary")を定義している。基本必要
        Set myDic2 = CreateObject("Scripting.Dictionary") 'myDic2の意味をCreateObject("Scripting.Dictionary")を定義している。基本必要
        With Worksheets("売上")
            varDate = .Range("J2", .Range("J" & Rows.Count).End(xlUp)).Value  '販売先のセルすべてをvarDateに入れる
        End With
        For Each c In varDate  'いつものforと一緒の考え方
            If Not c = Empty Then  'cが空白じゃなければ、下記作業を進める。空白なら飛ばす
                If Not myDic.Exists(c) Then '重複をなくすための動作
                    myDic.Add c, Null
                End If
            End If
       Next

       myKey = myDic.Keys

     For Each d In myKey
        If d Like "*" & customers & "*" Then
            If Not myDic2.Exists(d) Then
            myDic2.Add d, Null
            End If
        End If
    Next
    myKey2 = myDic2.Keys
    myForm.ComboBox1.List = myKey2

    Dim i, cnt, le As Integer 'iは整数と定義
    Dim ctrl As Control 'ctrlをControlと定義
    cnt = 0
    le = 0
        For Each i In myKey2
            branch = i
                With myForm2
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

                    ctrl.Height = 66
                    ctrl.Top = (cnt - le * 5) * 66
                    ctrl.Left = le * 158
                    ctrl.Width = 158
                    ctrl.Caption = branch
                    ctrl.FontName = "MS UI Gothic"
                    ctrl.FontBold = True
                    ctrl.FontSize = 13
                    ctrl.ForeColor = &H80000012
                    ctrl.BackColor = &H8000000F
                End With
                cnt = cnt + 1
                Set Newbtn(cnt).button = ctrl
'
        Next

    If le = 1 Then
        myForm2.Width = 328.2
    ElseIf le = 2 Then
        myForm2.Width = 495
    ElseIf le = 3 Then
        myForm2.Width = 638.5
    ElseIf le = 4 Then
        myForm2.Width = 793
    Else
        myForm2.Width = 165
    End If

    myForm2.Show


End Sub