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

(1) スプレッドシートを新規作成します.

既存のスプレッドシートを使用しても構いません.

(2) スクリプトエディタを開きます.

スプレッドシートの以下のメニューを選択します.

    ツール > スクリプトエディタ

(3) スクリプト ID を確認します.

スクリプトエディタの以下のメニューを選択します.

    ファイル > プロジェクトのプロパティ

(4) npm で依存パッケージをインストールします.

    $ npm install

(5) clasp でログインします.

    $ npm run clasp login

(6) push 先の scriptId を設定します.

    $ npm run clasp setting scriptId xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

(7) clasp で push します.

    $ npm run clasp push

(8) スプレッドシートを開き直します.

スプレッドシートを開き直すと, メニューに "File Manager" が表示されます.

## 各種コマンド

clasp のヘルプを表示する:

    $ npm run clasp help

script.google.com に push する:

    $ npm run clasp status
    $ npm run clasp push

ブラウザで開く:

    $ npm run clasp open

