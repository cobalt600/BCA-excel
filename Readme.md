# プレートリーダーから吐き出されるExcelを整形して印刷したい

## 背景
割と古いプレートリーダーが、これまた古いPCにつながっている。当然スタンドアロン仕様なので、私物USBデータをエクスポートして持ち出すのだが、あまりイケてないエクセルファイルが出力されてくる。

これまでなぜかフォーマットがイカれたxlsファイル（開くたびに「余白が変です」「セルの大きさが変です」「お前の頭も変です」みたいなエラーメッセージが出る）で出力されていたが、制御ソフト側の設定をいじったらまともな形のxlsxで吐き出されるようになった（sample.xlsxに例を示す）

## ほしいもの
最終産物が「印刷物」という点で、本職エンジニアの方はイライラしそうですが、バイオ実験系では、手書き実験ノートに出力データの貼付（保存先を明示）がグローバルスタンダードなので、仕方ないです。
- 生データ（自分のHDDに保存するため）
- 整形されたデータ（データを読んで解釈するため）
- 印刷用データ（実験ノートに貼付するため）

## いま使っているもの
下記の動をするエクセルのマクロを大急ぎで組んで使っている。

1. Hドライブ（USBメモリ）のBCAフォルダの中に入っているエクセルファイルを開く
1. 1枚目に「Result」という集計用のシートを作成する。
1. 欲しい情報が入っている「特定のシートの特定のセル」を絶対参照で指定、コピーして「Result」シートに貼り付ける
1. 上書き保存して終了

これで生データ（シート4枚）と、その左側に集計されたデータ（Result）ができる。あとはこれを自分のHDDにに移動させて、Resultシートを印刷する。

## 不満なこと
ラボに備え付けられているのはWinだが、研究に使用している科研費PCはMacで、VBAがうまく動かない。
Parallelsは入れてあるけど、いちいち移動するのも面倒くさい。

そもそもxlsxをいちいち開くので、挙動が遅い（学生実習のデータを捌くときは。これを班の数だけ処理しないといけない（20枚以上）

それならPythonの勉強も兼ねてOpenpyxlで作れば？ということで。

## 生データ（sample.xlsx）の解説
1枚目のシート（4つ表が並んでいるやつ）の一番下の表がBCAアッセイの濃度。左右（Horizontal）でデュプリケートしている。前半6つが検量線（2-1-0.5-0.25-0.125-0.0625）、ブランクを挟んで、その後のいくつかがサンプルの濃度。

3枚目のシートにRとR^2が表示されている。これは検量線の精度を指名していて、一応マイルールとして、これが0.98を下回ったら、やり直すことにしている（普段これを下回ることは殆どないので、下回ったらなにかを間違えている可能性があるため）

## やりたいこと
### 最低限
Result_sample.jpgみたいなシートを1シート目に作りたい。
- Resultの頭ににRun information（Rawデータの保存先とか）とカーブデータ（rとかr2とか）を追加
- 検量線の画像を貼り付け

### 発展編
データの平均値を取得したい
- サンプル濃度の生データと平均データを表示
- ResultのシートだけをPDFで出力して、印刷にまわしてもいいかもしれない（ラボのKyoceraのプリンタがExcelを印刷するとずれるため（これはサポート切れドライバの既知の問題点なので解消できない））

## 今できてること
ほぼイメージ通りのものが完成しています。あとはラボのWinで動くことを確認するだけ。
分岐させていたMaster以外のブランチは用済みなので削除しました。
