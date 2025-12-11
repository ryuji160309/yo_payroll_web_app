# PHPキャッシュ連携手順

このディレクトリには、GitHub Pages では利用できない PHP を使ってスプレッドシートをサーバー側で事前取得し、`php-cache` 配下に静的ファイルとして保存するためのスクリプトを配置しています。静的に保存された JSON / XLSX をフロントエンドが参照することで、従来のブラウザ経由ダウンロードよりも開始までの待ち時間を短縮できます。PHP が利用できない環境では、フロントエンドは自動で元のダウンロード方式にフォールバックします。

## フォルダ構成

- `php-cache/` … 生成されるキャッシュ格納先（書き込み可能にしてください）
  - `settings.json` / `settings.xlsx` … 設定スプレッドシートをキャッシュしたファイル
  - `sheets/<FILE_ID>.json` / `sheets/<FILE_ID>.xlsx` … 店舗ごとのスプレッドシートキャッシュ
- `php/fetch_settings.php` … 設定スプレッドシートを取得してキャッシュに保存
- `php/fetch_store_sheets.php` … 設定シート内の店舗一覧を読み取り、各店舗スプレッドシートをキャッシュ
- `php/config.php` … URL や列位置などの設定値
- `php/lib/` … 共通ヘルパーと XLSX パーサー

## 前提条件

- PHP 8.1 以上（`ZipArchive` 拡張有効）
- 外部 HTTP へ接続できるネットワーク
- Web サーバー（Apache/Nginx 等）から `php-cache` ディレクトリを配信できること

## 設置手順

1. リポジトリ全体を PHP が動作する一般サーバーへ配置します（`/var/www/html/yo_payroll_web_app` など）。
2. `php-cache` ディレクトリが Web サーバーから読み取り可能で、PHP 実行ユーザーが書き込み可能になるように権限を設定します。
   ```bash
   cd /var/www/html/yo_payroll_web_app
   mkdir -p php-cache/sheets
   chown -R www-data:www-data php-cache
   chmod -R 775 php-cache
   ```
3. Apache の例: `DocumentRoot` をリポジトリ直下に設定し、`php` ディレクトリを PHP として実行できるようにします。
   ```apacheconf
   <VirtualHost *:80>
     DocumentRoot /var/www/html/yo_payroll_web_app
     <Directory /var/www/html/yo_payroll_web_app>
       AllowOverride All
       Require all granted
     </Directory>
     <Directory /var/www/html/yo_payroll_web_app/php>
       Require all granted
       # PHP-FPM を利用している場合は、適宜 ProxyPassMatch 等を設定してください。
     </Directory>
   </VirtualHost>
   ```
4. Nginx + PHP-FPM の例: 静的ファイルはそのまま返し、`/php/` 配下のみ PHP-FPM に渡します。
   ```nginx
   server {
     listen 80;
     server_name example.com;
     root /var/www/html/yo_payroll_web_app;

     location /php/ {
       include fastcgi_params;
       fastcgi_pass unix:/run/php/php8.2-fpm.sock;
       fastcgi_param SCRIPT_FILENAME $document_root$fastcgi_script_name;
     }
   }
   ```

## 定期実行（cron）の例

設定スプレッドシートと店舗スプレッドシートを定期的に取得するために、以下のような cron を設定します。

```cron
# 毎時更新
0 * * * * /usr/bin/php /var/www/html/yo_payroll_web_app/php/fetch_settings.php >> /var/log/yo-payroll-cache.log 2>&1
5 * * * * /usr/bin/php /var/www/html/yo_payroll_web_app/php/fetch_store_sheets.php >> /var/log/yo-payroll-cache.log 2>&1
```

`SETTINGS_URL` を環境変数で上書きしたい場合は、`/etc/systemd/system/` 側でサービス化して環境変数を指定するか、cron エントリに `SETTINGS_URL=...` を付与してください。

## 動作確認

1. `php/fetch_settings.php` を実行して `php-cache/settings.json` が生成されることを確認します。
2. `php/fetch_store_sheets.php` を実行し、`php-cache/sheets/` に店舗ごとの JSON/XLSX が生成されることを確認します。
3. ブラウザで `settings.html` を開き、「PHP経由のキャッシュ取得時刻」が表示されていれば PHP キャッシュが利用されています。PHP にアクセスできない場合はトーストでフォールバックメッセージが表示され、従来のダウンロードに切り替わります。

## 注意事項

- 取得したファイルはサーバー上に平文で保存されます。公開範囲やアクセス制御は環境に応じて適切に設定してください。
- Google 側の URL 形式が変わった場合は、`php/config.php` の URL を更新してください。
- ダウンロードが失敗した場合は例外で終了し、標準エラー出力に理由を表示します。cron のログを参照して原因を特定してください。
