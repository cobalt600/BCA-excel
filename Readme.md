機械の設定をいじったらまともな形のxlsxで吐き出されるようになった
sample.xlsxの書式からずれることは絶対にない

1枚目のシート（4つ表が並んでいるやつ）の一番下の表がBCAアッセイの濃度
左右（Horizontal）でデュプリケートしている
前半6つが検量線、ブランクを挟んで、その後のいくつかがサンプル

やりたいことは
  Resultの頭ににRun information（Rawデータの保存先とか）とカーブデータ（rとかr2とか）を追加
  検量線の画像orグラフを書いて
  その下にサンプル濃度の生データと平均データを表示

これを吐き出しつつ、ResultのシートだけをPDFで出力して、印刷にまわしたい（実験ノートに貼り付けたい）
