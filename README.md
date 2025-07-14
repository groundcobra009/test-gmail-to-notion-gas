# Gmail to Notion 連携システム

GmailからNotionデータベースに自動でメールを転送するGoogle Apps Script（GAS）システムです。

## 概要

このシステムは、指定したGmailラベルのメールを自動的にNotionデータベースに同期します。Google Spreadsheetを設定画面として使用し、定期的にメールを処理することができます。

## 必要なもの

- Googleアカウント
- Notionアカウント
- Notion Integration Token
- Notionデータベース

## セットアップ手順

### 1. Google Spreadsheetの準備

1. 新しいGoogle Spreadsheetを作成
2. サンプルファイルをコピー: [https://docs.google.com/spreadsheets/d/1soX-kLFySIM7OjetVCW7hpsk_Q8oaUuTe06B2ngOQ48/edit?usp=sharing](https://docs.google.com/spreadsheets/d/1soX-kLFySIM7OjetVCW7hpsk_Q8oaUuTe06B2ngOQ48/edit?usp=sharing)

### 2. Google Apps Scriptのセットアップ

1. Spreadsheet上で「拡張機能」→「Apps Script」を選択
2. `appsscript.json`と`コード.js`の内容をコピー&ペースト
3. プロジェクトを保存

### 3. Notion側の準備

#### 3.1 Integration Tokenの取得
1. [Notion Developers](https://www.notion.so/my-integrations)にアクセス
2. 「New integration」をクリック
3. 適切な名前を入力し、Workspaceを選択
4. 「Submit」をクリック
5. 表示されるTokenをコピー（後で使用）

#### 3.2 データベースの作成と設定
1. Notionで新しいデータベースを作成
2. タイトル列（デフォルト名：「タイトル」）があることを確認
3. データベースページで「・・・」→「Add connections」→作成したIntegrationを選択
4. データベースURLからIDを取得（例：`https://notion.so/xxx?v=yyy`の`xxx`部分）

### 4. 設定シートの作成と設定

1. Spreadsheetに戻り、「Gmail-Notion連携」メニューから「設定シートを作成」を実行
2. 作成された設定シートに以下を入力：

| 設定項目 | 値 | 説明 |
|----------|-----|------|
| Notion インテグレーションキー | `secret_xxx...` | 取得したIntegration Token |
| Notion データベースID | データベースID | NotionデータベースのID |
| タイトルプロパティ名 | タイトル | Notionデータベースのタイトル列名 |
| Gmailラベル名 | Notion | 処理対象のGmailラベル名 |
| 最終処理日時 | （空白） | 自動更新される |

### 5. Gmailラベルの作成

1. Gmailで処理対象のラベル（デフォルト：「Notion」）を作成
2. Notionに送信したいメールにラベルを付与

### 6. 権限の設定

初回実行時に以下の権限を承認：
- Gmailの読み取り
- Spreadsheetの編集
- 外部URLへのアクセス（Notion API）

## 使用方法

### 手動実行
「Gmail-Notion連携」メニューから「手動で同期を実行」を選択

### 自動実行の設定
「Gmail-Notion連携」メニューから「1時間ごとのトリガーを設定」を選択

### トリガーの削除
「Gmail-Notion連携」メニューから「トリガーを削除」を選択

## 機能説明

### 主な機能
- Gmailの指定ラベルからメール取得
- Notionデータベースへのページ作成
- メール情報（件名、送信者、日時、本文）の転送
- 最終処理日時の記録（重複処理防止）
- 定期実行トリガーの設定

### データ構造
Notionに作成されるページには以下が含まれます：
- タイトル：メールの件名
- メール情報：送信者と日時
- 本文：メール本文（2000文字制限）

## 注意事項

- メール本文は2000文字まで（Notion APIの制限）
- 1回の実行で最大50件のメールを処理
- 最終処理日時以降のメールのみ処理
- HTMLメールはプレーンテキストに変換

## トラブルシューティング

### よくあるエラー
1. **「設定シートが見つかりません」**
   - 「設定シートを作成」を実行してください

2. **「NotionのインテグレーションキーとデータベースIDを設定してください」**
   - 設定シートの値を正しく入力してください

3. **「ラベルが見つかりません」**
   - Gmailで指定したラベルを作成してください

4. **「Notion APIエラー」**
   - Integration TokenとデータベースIDを確認
   - データベースにIntegrationが接続されているか確認

### ログの確認
Apps Scriptエディタの「実行数」タブでログとエラーを確認できます。

## セキュリティ

- Integration Tokenは厳重に管理してください
- Spreadsheetの共有設定に注意してください
- 不要になったIntegrationは削除してください