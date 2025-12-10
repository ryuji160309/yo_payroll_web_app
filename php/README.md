# PHP キャッシュ実装案

GitHub Pages では PHP が実行できないため、ここでは **PHP 対応サーバーに配置して使うための下準備** をまとめています。`php/` 以下をそのままコピーすれば、Google スプレッドシートの XLSX をサーバー側で定期取得・キャッシュできます。

## ディレクトリ構成
- `config.php` … 取得対象のスプレッドシート URL や保存先を定義します。
- `helpers.php` … URL 変換とキャッシュ用の共通関数です。
- `fetch.php` … 取得スクリプト本体。CLI または HTTP から実行できます。
- `data/` … キャッシュファイル置き場（`.gitkeep` は空ディレクトリ保持用）。

## 使い方
1. `php/` ディレクトリを PHP が動くサーバーにコピーします。
2. 必要に応じて `config.php` の URL（設定シートや各店舗のスプレッドシート）を書き換えます。
3. コマンドラインで実行する場合:
   ```bash
   php fetch.php
   # JSON で確認したい場合
   php fetch.php --json
   ```
4. ブラウザ経由で `fetch.php` にアクセスすると、JSON 形式の最新キャッシュ状況を返します。
5. `data/` 配下に `settings.xlsx`、`<店舗キー>.xlsx`、`settings.json`、`cache-index.json` が保存されるので、フロントエンドからは Google ではなくこれらのファイルを参照させる運用が可能です。

## 定期実行（例）
cron で 5 分〜15 分おきに実行すると、クライアントは常にサーバー側のキャッシュから素早くデータを取得できます。

```
*/10 * * * * /usr/bin/php /path/to/php/fetch.php > /var/log/yo-payroll-fetch.log 2>&1
```

## 留意点
- `fetch_and_cache_sheet` は `If-Modified-Since` を送信し、変更がない場合は 304 を受け取ってスキップします。Google 側のトラフィックを抑えつつ最新状態を維持できます。
- CSV での取得も `fetch_and_cache_sheet($key, $url, $dataDir, 'csv')` で可能ですが、現状の設定は XLSX です。
- 生成された `cache-index.json` をフロントエンドで読めば、いつキャッシュされたデータなのか把握できます。
- `settings.json` には設定シートの解析結果（店舗一覧、パスワード、賃金設定、除外ワード、最終取得時刻、キャッシュ済み XLSX のパス）が含まれ、PHP が使えない環境でもそのまま配信できます。
