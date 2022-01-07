#プレートリーダーから吐き出されたエクセルファイルから必要な情報を抜き出すプログラム
#吸光度とかはいらないけど、最小二乗法でフィッティングされた濃度（平均が取れればなおよしだが、場所がコロコロ変わる）とr，r^2が最低限必要
#余力があればグラフも吐き出したい
#最終的には【印刷して】実験ノートに貼る必要がある（左右の余白を確保してA4に印刷、上半分に全部の情報が収まっていると使い勝手が良い）
#最終アウトプットはxlsxでもwordでもPDFでもいい。改変可能性を減らすためにはPDFが良いかな？
#同時並行でGitの使い方を理解したい、ので、とりあえずネットにあった適当なコードを貼り付けたものを初版として、これを改変していく。

#エクセルファイルを読み込む（とりあえずカレントディレクトリでいいや）
import pandas as pd
input_file_name = 'BCA.xlsx'
input_book = pd.ExcelFile(input_file_name)

#シートの名称は不変だけど、とりあえず取得
input_sheet_name = input_book.sheet_names
num_sheet = len(input_sheet_name)

#シートの名前と総数を表示してみる
print ("Sheet の数:", num_sheet)
print (input_sheet_name)

#ここから先は後で書く
