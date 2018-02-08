# OpenCsvByExcel

## 概要 - Overview

ExcelでCSV/TSVを開く際、テキストインポートウィザードなどの煩雑な手法を使わずに、ファイルのドラッグ＆ドロップでファイルを開くことができます。  
機械的に出力されたCSV/TSVを取り扱う際に便利です。

## 特徴 - Features

* 0から始まる数値項目の先頭0が欠損してしまう状況を防ぎます
* カラム内に改行を含む項目を行落ちさせずに正しく1カラムとして読み込みます
* CSV/TSVファイルの文字エンコーディングを自動で判断します
* 区切り文字を拡張子から推定します

## システム要求 - System Requirements

* Microsoft Office Excel(現在サポートされているバージョンに限る)
* .NET Framework 4.5以上

## インストール方法 - Install

1. [リリース](https://github.com/mitaken/OpenCsvByExcel/releases)より最新版をダウンロード
1. 適当なフォルダにZipファイルを展開します
1. Shortcut.vbsを実行すると右クリックメニューの送るにショートカットが作成されます

## アンインストール方法 - Uninstall

1. フォルダ内にあるShortcut.vbsを実行すると、右クリックメニューの送るのショートカットが削除されます
1. フォルダを削除します（レジストリなどは触っていません）

## 使い方 - Usage

1. CSV/TSVファイルを右クリックし、送るから ```Open CSV by Excel``` を選択すると自動でExcelが立ち上がります

## 設定情報 - Configurations

アプリケーションフォルダに内包されている ```OpenCsvByExcel.exe.config``` を開いて修正してください

### CsvExtensions

CSV（カンマ区切りファイル）として取り扱う拡張子を羅列します

### TsvExtensions

TSV（タブ区切りファイル）として取り扱う拡張子を羅列します

### ParallelOpen

複数ファイルをドラッグ＆ドロップした際に並列で開くファイル数を設定します  
1に設定するとファイルは直列で開かれます

### FallbackCharset

文字エンコーディングの自動検出が失敗した際にこのエンコーディング指定でファイルを読み取ります

### MaxColumnSize

1カラムの最大文字数を設定します  
1カラムあたりに大きな文字が設定される場合はこの値を修正してください

### HasHeaderRecord

ヘッダー行を認識させます  
ヘッダーが存在するファイルを読み込む場合はTrue、存在しない場合はFalseを指定してください

### AdjustColumnWidth

カラム幅をデータの内容に応じて自動で調整します  
自動調整する場合はTrue、初期サイズのままにする場合はFalseを指定してください

### DetectColumnCountChanges

行の途中でカラム数が変更されていた場合に処理を中断します  
中断する場合はTrue、継続する場合はFalseを指定します

## 使用ライブラリ - Libraries

* [Mozilla Universal Charset Detector](https://github.com/errepi/ude)
  * 文字コード自動検出
* [CsvHelper](https://github.com/JoshClose/CsvHelper)
  * CSV/TSVファイルの読み込み
* [ComInvoker](https://github.com/mitaken/ComInvoker)
  * ADO/ExcelのCOMリソース開放

## バージョン情報 - Version Histories
* 1.0
  * 初版リリース
* 1.1
  * [Range.CopyFromRecordset](https://msdn.microsoft.com/ja-jp/vba/excel-vba/articles/range-copyfromrecordset-method-excel) から [QueryTables.Add](https://msdn.microsoft.com/ja-jp/vba/excel-vba/articles/querytables-add-method-excel) に読み込みを変更
* 1.2
  * ヘッダー行とカラム数の不一致を許容するオプションを追加
* 1.3
  * エラー表示改善およびヘッダー行のチェックを実施し、カラム名が空の場合にエラーとならないように対応した
