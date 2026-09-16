LTspice CSV Converter

概要
LTspiceから出力した波形テキストを、Excelで扱いやすいCSVへ変換します。
例: (2.083e+01dB,-1.66e-01°) を V_out_dB と V_out_phase_deg の2列に分けます。

使い方 1: GUI
1. Anaconda Promptを開きます。
2. このフォルダへ移動します。
3. 次を実行します。
   python ltspice_csv_converter.py
4. 変換したいLTspiceのtxtファイルを選びます。
5. 元ファイルと同じフォルダに「元ファイル名_excel.csv」が作成されます。

使い方 2: コマンドで変換
python ltspice_csv_converter.py "C:\path\to\input.txt"

出力先を指定する場合:
python ltspice_csv_converter.py "C:\path\to\input.txt" -o "C:\path\to\output.csv"

複数ファイルをまとめて変換する場合:
python ltspice_csv_converter.py "C:\path\to\a.txt" "C:\path\to\b.txt" -o "C:\path\to\output_folder"

補足
- Python標準機能だけで動くので、追加インストールは不要です。
- 出力CSVはUTF-8 BOM付きなので、日本語版Excelでも開きやすい形式です。
- 入力文字コードは utf-8 / cp932 / shift_jis / cp1252 / latin-1 を順に試します。
