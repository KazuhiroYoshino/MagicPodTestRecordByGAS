# GASを使ったテスト結果の転記
## magicpod-analyzerを使って、API経由でMagicPodのテスト結果を取得し、テスト結果表にsuccess / failedを記録します。
* テスト結果表（スプレッドシート）と同じGoogleフォルダに、magicpod-analyzerで取得したJSONデータを置きます。
* GASを実行します。
## スプレッドシートは、テスト計画シートとテスト結果シートを用意します。
* 30分単位で実行するMagicPod自動テストの設定名称をテスト計画シートに記入しておきます。
* 曜日毎にシート範囲に名前を付けておきます。月曜日のテストは、rangeMonday
* テスト結果シートは、横軸に30分単位の計画した時刻、縦軸は日付とします。
* JSONデータからテスト計画に記載の自動テスト設定名称を検索して、テスト日付の行に結果を色分けして転記します。
