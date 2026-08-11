# StoreStock-Web Version1（チェック版）

GitHub Pagesへ置くのは `index.html / style.css / config.js / app.js` だけです。
実在庫Excel・実売上データはGitHubへ保存しません。

## 月曜在庫の取り込み
サイト下部の「月曜在庫の取り込み」でExcelを選択し、内容確認後にコピーします。
その内容を共通の月曜在庫GoogleスプレッドシートのA1へ貼り付けます。

## 重要
専用の月曜在庫GoogleスプレッドシートURLがまだ未設定です。
作成後、`config.js` の `STOCK_SHEET` に id / gid / editUrl を入れます。
未設定の状態では検索を開始しない安全設計です。

## 可変対応
- 行数増減OK
- 店舗列の位置変更OK
- 店舗列追加で既存店舗は壊れない
- 店舗名は見出しで判定
