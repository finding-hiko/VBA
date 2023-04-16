# オートフィルタ

【概要】
指定セルに販売先・仕入先・商品・規格等の情報を入力することで
売上データからフィルタリングしたデータを簡易的に呼び起こすシステムです。

【使用技術】
・ ExcelVBA

【使用例】
1."入力画面"シートを開きます。（"データ"シートは売上データの格納庫になっております。）
![スクリーンショット 2023-04-16 162237](https://user-images.githubusercontent.com/118088137/232280897-33e829b1-8fcc-446b-89ee-3718dac5a2d7.PNG)

2."入力画面"シートの検索したい項目にキーワードを入力し、隣の実行ボタンを押すと"データ"シートより
抽出されたデータが下部に表示されます。
（今回は商品名の入力セルに"りんご"と入力し検索しました）
![スクリーンショット 2023-04-16 161952](https://user-images.githubusercontent.com/118088137/232281045-194a6ae9-a5a2-4192-895f-a2e14e0e5e66.PNG)

3.さらに条件を絞って検索したい場合には、他の項目にもキーワードを入れると、さらに絞って抽出された
データが下部に表示されます。
![スクリーンショット 2023-04-16 162430](https://user-images.githubusercontent.com/118088137/232281368-e37dff91-7528-4903-82cf-0485fc01c82f.PNG)

![スクリーンショット 2023-04-16 162543](https://user-images.githubusercontent.com/118088137/232281387-db9034d7-858d-4bc3-845e-da3217c2a136.PNG)

4."オートフィルタを解除"のボタンを押すとフィルタリングが解除され、最初の画面に戻ります。

# 住所録

入力フォームに必要事項を明記することで、レターパックに貼るシール印刷用の
フォーマットに転記されるシステムです。

住所録の保存機能や呼び起こし機能も搭載されております。

使用技術
・ ExcelVBA

# 依頼書

見積依頼、サンプル依頼、問い合わせ等の依頼に際しての
依頼書フォーマットです。

送付先名や担当者等の最低限の入力部分をボタン操作にしているので
キーボード操作を削減し時間短縮を図っております。

使用技術
・ ExcelVBA

# 納品書

売上データをもとに顧客・物件ごとの納品書を作成するシステムです。

使用技術
・ ExcelVBA
