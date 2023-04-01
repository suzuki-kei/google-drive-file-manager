# google-drive-file-manager

## 概要

Google Drive のファイルを管理するための機能を提供します.
Google スプレッドシートのメニューから以下の操作が可能になります.

 * ドキュメントインデックスの作成
 * ファイル名の一括変更

## 前提条件

事前に以下がインストールされている必要があります.

 * npm (https://www.npmjs.com/)

## 使用方法

(1) npm で依存パッケージをインストールします.

    $ npm install

(2) clasp でログインします.

    $ npm exec -c 'clasp login'

(3) Google Apps Script のプロジェクトを作成します.

    $ npm exec -c 'clasp create --type sheets --title "File Manager" --rootDir src'

NOTE: .clasp.json が生成される場所が間違っているようなので移動します.

    $ mv src/.clasp.json .

(4) clasp で push します.

    $ npm exec -c 'clasp push'

(5) スプレッドシートを開きます.

メニューに "File Manager" が表示されます.

## 各種コマンド

clasp のヘルプを表示する:

    $ npm exec -c 'clasp help'

script.google.com に push する:

    $ npm exec -c 'clasp status'
    $ npm exec -c 'clasp push'

ブラウザで開く:

    $ npm exec -c 'clasp open'

