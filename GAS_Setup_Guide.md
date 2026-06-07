# 📊 Googleスプレッドシート・GAS自動連携セットアップガイド

このガイドでは、LP（ランディングページ）の申し込みフォームとGoogleスプレッドシートを連携させ、ジロウさんが**「スプレッドシートで入金チェックを入れるだけで、LPの進捗ゲージが自動で更新される仕組み」**を構築する手順を解説します。

---

## 🛠️ ステップ 1：Googleスプレッドシートの新規作成

1. [Googleスプレッドシート](https://sheets.google.com/) を開き、新しいスプレッドシート（空白）を作成します。
2. スプレッドシートのタイトル（左上）を `コザ高女子テニス部遠征資金造成・申込管理` に変更します。
3. シートの1行目（ヘッダー）に、左から順に以下の項目を入力します。
   * **A1**: `タイムスタンプ`
   * **B1**: `お名前`
   * **C1**: `フリガナ`
   * **D1**: `メールアドレス`
   * **E1**: `お電話番号`
   * **F1**: `ご関係`
   * **G1**: `ネイビー S`
   * **H1**: `ネイビー M`
   * **I1**: `ネイビー L`
   * **J1**: `ネイビー XL`
   * **K1**: `バーガンディー S`
   * **L1**: `バーガンディー M`
   * **M1**: `バーガンディー L`
   * **N1**: `バーガンディー XL`
   * **O1**: `アイビーグリーン S`
   * **P1**: `アイビーグリーン M`
   * **Q1**: `アイビーグリーン L`
   * **R1**: `アイビーグリーン XL`
   * **S1**: `受け取り方法`
   * **T1**: `ご配送先住所`
   * **U1**: `寄付金額`
   * **V1**: `送金方法`
   * **W1**: `メッセージ`
   * **X1**: `合計申込金額`
   * **Y1**: `★入金確認（チェック）`

*(※項目名はお好きな形に変更しても構いませんが、GASの自動書き込みは左からの列順（A列〜Y列）で行われます。)*

---

## 💻 ステップ 2：GAS（Google Apps Script）の書き込み

1. スプレッドシートのメニューバーから **「拡張機能」** ➔ **「Apps Script」** を選択します。
2. 最初から表示されている `function myFunction() { ... }` のコードをすべて消去します。
3. 以下のプログラムコードを丸ごとコピーして貼り付けます。

```javascript
/**
 * 【重要】初回のみ実行する設定用プログラム（メール送信の許可を出すため）
 * GASエディタ上のプルダウンで「authorizeEmail」を選択して「実行」を押してください。
 */
function authorizeEmail() {
  MailApp.getRemainingDailyQuota();
  console.log("メール送信の許可が完了しました！");
}

/**
 * フォームデータの受信 (POSTリクエスト時の処理)
 * LPから送信されたデータをスプレッドシートの最終行に追記します
 */
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // LPから送信されたデータを受け取る
    var data = e.parameter;
    
    var timestamp = new Date();
    
    // 送信されてきたデータをスプレッドシートの配列に並べる
    var row = [
      timestamp,
      data.name || "",
      data.kana || "",
      data.email || "",
      data.phone || "",
      data.relation || "",
      Number(data.navyS) || 0,
      Number(data.navyM) || 0,
      Number(data.navyL) || 0,
      Number(data.navyXL) || 0,
      Number(data.burgundyS) || 0,
      Number(data.burgundyM) || 0,
      Number(data.burgundyL) || 0,
      Number(data.burgundyXL) || 0,
      Number(data.ivyGreenS) || 0,
      Number(data.ivyGreenM) || 0,
      Number(data.ivyGreenL) || 0,
      Number(data.ivyGreenXL) || 0,
      data.delivery || "",
      data.address || "",
      Number(data.donationAmount) || 0,
      data.payment || "",
      data.message || "",
      Number(data.totalAmount) || 0,
      false // ★入金確認チェックボックスの初期値 (未チェック: FALSE)
    ];
    
    // シートの最終行に書き込み
    sheet.appendRow(row);
    
    // 追加した行のY列（25列目）に自動でチェックボックスを挿入
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 25).insertCheckboxes();
    
    // ▼ 自動返信メールの送信処理 ▼
    var emailAddress = data.email || "";
    if (emailAddress) {
      var supporterName = data.name || "支援者";
      var subject = "【コザ高女子テニス部】資金造成ご支援のお申し込み完了と送金のお願い";
      var tshirtDetails = [];
      var nS = Number(data.navyS) || 0;
      var nM = Number(data.navyM) || 0;
      var nL = Number(data.navyL) || 0;
      var nXL = Number(data.navyXL) || 0;
      var bS = Number(data.burgundyS) || 0;
      var bM = Number(data.burgundyM) || 0;
      var bL = Number(data.burgundyL) || 0;
      var bXL = Number(data.burgundyXL) || 0;
      var iS = Number(data.ivyGreenS) || 0;
      var iM = Number(data.ivyGreenM) || 0;
      var iL = Number(data.ivyGreenL) || 0;
      var iXL = Number(data.ivyGreenXL) || 0;
      
      var nArr = [];
      if(nS > 0) nArr.push("S:" + nS + "枚");
      if(nM > 0) nArr.push("M:" + nM + "枚");
      if(nL > 0) nArr.push("L:" + nL + "枚");
      if(nXL > 0) nArr.push("XL:" + nXL + "枚");
      if(nArr.length > 0) tshirtDetails.push("【ネイビー】" + nArr.join(" / "));

      var bArr = [];
      if(bS > 0) bArr.push("S:" + bS + "枚");
      if(bM > 0) bArr.push("M:" + bM + "枚");
      if(bL > 0) bArr.push("L:" + bL + "枚");
      if(bXL > 0) bArr.push("XL:" + bXL + "枚");
      if(bArr.length > 0) tshirtDetails.push("【バーガンディー】" + bArr.join(" / "));
      
      var iArr = [];
      if(iS > 0) iArr.push("S:" + iS + "枚");
      if(iM > 0) iArr.push("M:" + iM + "枚");
      if(iL > 0) iArr.push("L:" + iL + "枚");
      if(iXL > 0) iArr.push("XL:" + iXL + "枚");
      if(iArr.length > 0) tshirtDetails.push("【アイビーグリーン】" + iArr.join(" / "));

      var body = supporterName + " 様\n\n" +
        "この度は、コザ高女子テニス部の資金造成へ温かいご支援を賜り、心より感謝申し上げます。\n" +
        "以下の内容でお申し込みを受け付けました。\n\n" +
        "【お申し込み内容】\n" +
        (tshirtDetails.length > 0 ? "・支援Tシャツ：\n  " + tshirtDetails.join("\n  ") + "\n" : "") +
        "・ご支援総額： " + Number(data.totalAmount).toLocaleString() + " 円\n" +
        "・ご希望の送金方法： " + (data.payment === 'bank' ? '銀行振込' : 'PayPay送金') + "\n\n" +
        "大変お手数ですが、6月30日（火）までに以下の方法で送金をお願いいたします。\n\n";

      if (data.payment === 'bank') {
        body += "■ 銀行振込先\n" +
          "〇〇銀行 〇〇支店\n" +
          "口座番号：〇〇〇〇〇〇\n" +
          "口座名義：コザ高校女子テニス部保護者会\n\n";
      } else if (data.payment === 'paypay') {
        body += "■ PayPay送金先\n" +
          "PayPay ID: yujiro0070111\n" +
          "※PayPayアプリの「送る・受け取る」から上記のIDを検索し、ご支援金額のご送金をお願いいたします。\n" +
          "※送金時のメッセージ欄にご本人様のお名前をご入力ください。\n\n";
      }

      body += "※ご入金が確認でき次第、特設サイトの進捗状況へ反映させていただきます。\n" +
        "引き続き、コザ高女子テニス部への温かいご声援をよろしくお願いいたします！\n\n" +
        "---------------------------------------\n" +
        "コザ高等学校女子テニス部 保護者会・部員一同\n" +
        "---------------------------------------";

      try {
        MailApp.sendEmail({
          to: emailAddress,
          subject: subject,
          body: body,
          name: "コザ高女子テニス部 保護者会"
        });
      } catch (e) {
        console.error("メール送信エラー: " + e);
      }
    }
    // ▲ 自動返信メールの送信処理 ここまで ▲
    
    var response = { status: "success" };
    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch(error) {
    var response = { status: "error", message: error.toString() };
    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * リアルタイム進捗状況の取得 (GETリクエスト時の処理)
 * 「入金確認」チェックボックスがONの支援額と人数を集計して返します
 */
function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    
    var totalAmount = 0;
    var supporterCount = 0;
    
    if (lastRow > 1) {
      // 2行目から最終行までの全データを取得
      var rows = sheet.getRange(2, 1, lastRow - 1, 17).getValues();
      
      for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        // index 16 (17列目 = Q列) が「入金確認」チェックボックス
        // index 15 (16列目 = P列) が「合計申込金額」
        var isConfirmed = row[16];
        if (isConfirmed === true || isConfirmed === "TRUE" || isConfirmed === "true") {
          totalAmount += Number(row[15]) || 0;
          supporterCount++;
        }
      }
    }
    
    var result = {
      status: "success",
      amount: totalAmount,
      supporters: supporterCount
    };
    
    // JSON形式でレスポンス
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch(error) {
    var errorResult = {
      status: "error",
      message: error.toString()
    };
    return ContentService.createTextOutput(JSON.stringify(errorResult))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
```

4. 貼り付けたら、エディタの上部にある **「保存」アイコン（フロッピーディスクのマーク）** をクリックして保存します。

---

## 🚀 ステップ 3：Webアプリとしての公開（デプロイ）

プログラムを外部のLPから呼び出せるように公開します。

> [!WARNING]
> **GASの超重要ルール**：スクリプト（コード）を後から書き換えた場合、「上書き保存」だけでは反映されません！必ず毎回 **「新しいデプロイ」** を行って新しいURLをコピーするか、デプロイの管理から「新バージョン」を作成する必要があります。

1. 画面右上にある **「デプロイ」** ボタンをクリックし、**「新しいデプロイ」** を選択します。
2. 種類の選択（歯車アイコン）をクリックし、**「ウェブアプリ」** を選択します。
3. 設定項目を以下のように指定します。
   * **説明**: `遠征資金造成API` (何でもOKです)
   * **次のユーザーとして実行**: `自分（あなたのGoogleアカウント）`
   * **アクセスできるユーザー**: `全員` **（※必ず「全員」にしてください。そうしないとLPからアクセスできません）**
4. **「デプロイ」** ボタンをクリックします。
5. 初回は「アクセスの承認」を求められるので、以下の手順で承認します：
   * 「アクセスを承認」をクリック。
   * 自分のGoogleアカウントを選択。
   * 警告画面が出たら、左下の **「詳細」**（または Advanced）をクリックし、**「コザ高女子テニス部遠征資金造成・申込管理（安全ではないページ）に移動」** をクリックします。
   * 次の画面で **「許可」** (Allow) をクリックします。
6. デプロイが完了すると、**「ウェブアプリのURL」** が表示されます。このURL（`https://script.google.com/macros/s/XXXXXX/exec`）をコピーしておきます。

---

## 🔗 ステップ 4：LPのHTMLファイルにURLを紐付ける

1. 作成したLPのファイル `index.html` をエディタで開きます。
2. 1465行目付近（または `const GAS_API_URL` で検索）にある以下の記述を見つけます。

```javascript
        // Google Apps Script Web App URL (※ジロウさんがGAS公開後に取得するURLをここに貼り付けます)
        const GAS_API_URL = 'https://script.google.com/macros/s/YOUR_GAS_API_ID/exec';
```

3. `'https://script.google.com/macros/s/YOUR_GAS_API_ID/exec'` の部分を、先ほどコピーした **「ウェブアプリのURL」** に書き換えます。

これで連携設定は完了です！

---

## 📈 運用の流れ

1. **LPから申し込みが入る**
   * 自動的にスプレッドシートに新しい行が追加され、タイムスタンプや金額、サイズ内訳、Q列のチェックボックス（最初は未チェック）がセットされます。
2. **ジロウさんが銀行口座やPayPayを確認する**
   * 「コザ太郎」さんから3,000円の入金があるのを確認。
3. **スプレッドシートでチェックを入れる**
   * スプレッドシートの「コザ太郎」の行の **Q列（★入金確認）のチェックボックスをポチッとクリック（ON）** にします。
4. **LPの進捗状況が自動アップデート！**
   * LPを開き直すと、チェックを入れたデータだけが集計され、進捗メーター（現在の支援額・達成率・人数）がリアルタイムに連動して増加します。
