function fetchSpecificEmails() {
  const now = new Date();
  // 現在時刻から3分前のUNIXタイムスタンプを取得
  const timeOffset = 5 * 60 * 1000;
  // const timeOffset = 3 * 60 * 1000;
  const targetTime = new Date(now.getTime() - timeOffset);
  const unixTimestamp = Math.floor(targetTime.getTime() / 1000);
  // スプレッドシート情報取得
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetTitle = spreadsheet.getName();
  const titleMatch = spreadsheetTitle.match(/^(\d{4})年(\d{1,2})月/);
  if (!titleMatch) {
    throw new Error("スプレッドシートのタイトルから年月を抽出できませんでした。");
  }
  const year = parseInt(titleMatch[1], 10);
  const month = parseInt(titleMatch[2], 10);
  const startDate = new Date(year, month - 1, 1);
  const endDate = new Date(year, month, 0, 23, 59, 59);
  const startTimestamp = Math.floor(startDate.getTime() / 1000);
  const endTimestamp = Math.floor(endDate.getTime() / 1000);
  // スプレッドシートは月ごとに新しく作成され、指定された年月以外のシートに自動でデータが記入されるのを防ぐため
  // 検索時間が範囲外の場合は処理を終了
  if (unixTimestamp < startTimestamp || unixTimestamp > endTimestamp) {
    console.log("UNIXタイムスタンプがスプレッドシートの年月の範囲外です。処理を終了します。");
    return;
  }
  // 共通ラベルの取得または作成
  let label = GmailApp.getUserLabelByName("転記済み");
  if (!label) {
    label = GmailApp.createLabel("転記済み");
    console.log("ラベル '転記済み' を作成しました。");
  }
  // ▼ DockpitFree 処理
  const dockpitSheet = spreadsheet.getSheetByName("DockpitFree申し込み");
  if (dockpitSheet) {
    const dockpitQuery = `subject:"フォームハンドラー: DockpitFree申込_アカウント発行" OR "フォームハンドラー: DockpitFree申込_既存ユーザー" after:${unixTimestamp}`;
    const dockpitThreads = GmailApp.search(dockpitQuery);
    dockpitThreads.forEach(thread => {
      if (thread.getLabels().some(l => l.getName() === "転記済み")) return;
      thread.getMessages().forEach(msg => {
        const body = msg.getPlainBody();
        const company = body.match(/Company:\s+(.+)/)?.[1]?.trim() || "会社不明";
        const lastName = body.match(/Last name:\s+(.+)/)?.[1]?.trim() || "";
        const firstName = body.match(/First name:\s+(.+)/)?.[1]?.trim() || "";
        const fullName = lastName && firstName ? `${lastName} ${firstName}` : "不明";
        dockpitSheet.appendRow(["", company, "", fullName]);
        thread.markRead();
        thread.addLabel(label);
      });
    });
    console.log(`DockpitFree：${dockpitThreads.length} 件のスレッドを処理しました。`);
  } else {
    console.error("シート 'DockpitFree申し込み' が存在しません。");
  }
  // ▼ ホワイトペーパーDL 処理
  const whitepaperSheet = spreadsheet.getSheetByName("ホワイトペーパーDL");
  if (whitepaperSheet) {
    const whitepaperQuery = `subject:"フォーム: 【マナミナ】ホワイトペーパー" after:${unixTimestamp}`;
    const whitepaperThreads = GmailApp.search(whitepaperQuery);
    const startRow = 6;
    const lastRow = whitepaperSheet.getLastRow();
    const rowCount = Math.max(1, lastRow - startRow + 1);
    const colStart = 2;
    const colCount = 3;
    const existingData = whitepaperSheet.getRange(startRow, colStart, rowCount, colCount).getValues();
    // 取得先に\tが入ってくるものが出てきたため、\tを削除して文字列として認識するための処理を追加
    const existingSet = new Set(
      existingData.map(row =>
        `${String(row[0] ?? '').replace(/\t/g, '').trim()}|${String(row[1] ?? '').replace(/\t/g, '').trim()}`
      )
    );
    whitepaperThreads.forEach(thread => {
      if (thread.getLabels().some(l => l.getName() === "転記済み")) return;

      const subject = thread.getFirstMessageSubject();
      const match = subject.match(/ホワイトペーパー_?(.+)\s[^ ]*$/);

      if (!match || !match[1]) {
        console.error('件名にマッチしませんでした。')
      }

      const messages = thread.getMessages();
      const body = messages[0].getPlainBody(); // 最初のメッセージだけ処理

      // \tや:が含まれていたら削除してから転記するように修正
      const company = body.match(/Company:\s+(.+)/)?.[1]?.replace(/[:\t]/g, '').trim() || "会社不明";
      const lastName = body.match(/Last name:\s+(.+)/)?.[1]?.replace(/[:\t]/g, '').trim() || "";
      const firstName = body.match(/First name:\s+(.+)/)?.[1]?.replace(/[:\t]/g, '').trim() || "";
      const fullName = lastName && firstName ? `${lastName} ${firstName}` : "不明";
      const key = `${company}|${fullName}`;

      if (existingSet.has(key)) {
        console.log(`既に存在するデータです（${key}）。スキップ。`);
        return;
      }

      whitepaperSheet.appendRow(["", company, fullName, "", "", "", match[1]]);
      thread.markRead();
      thread.addLabel(label);
    });
    console.log(`ホワイトペーパーDL：${whitepaperThreads.length} 件のスレッドを処理しました。`);
  } else {
    console.error("シート 'ホワイトペーパーDL' が存在しません。");
  }

  // ▼ tool ferret ホワイトペーパーDL 処理
  const ferretSheet = spreadsheet.getSheetByName("ferret資料請求");
  if (ferretSheet) {
    const ferretQuery = `subject:"【tool ferret】" after:${unixTimestamp}`;
    const ferretThreads = GmailApp.search(ferretQuery);

    const startRow = 5;
    const colStart = 2;
    const colCount = 3;

    const lastRow = ferretSheet.getLastRow();
    const rowCount = lastRow - startRow + 1;

    const existingData = rowCount > 0
      ? ferretSheet.getRange(startRow, colStart, rowCount, colCount).getValues()
      : [];

    const existingSet = new Set(
      existingData.map(row =>
        `${String(row[0] ?? '').replace(/\t/g, '').trim()}|${String(row[2] ?? '').replace(/\t/g, '').trim()}`
      )
    );

    ferretThreads.forEach(thread => {
      if (thread.getLabels().some(l => l.getName() === "転記済み")) return;

      const body = thread.getMessages()[0].getPlainBody();

      const nameMatch = body.match(/([^\s]+)\s*様よりホワイトペーパーのダウンロード/);
      const materialMatch = body.match(/資料名:\s*(.+)/);
      const companyMatch = body.match(/勤務先企業名:\s*(.+)/);

      const name = nameMatch ? nameMatch[1].trim() : "不明";
      const material = materialMatch ? materialMatch[1].trim() : "不明";
      const company = companyMatch ? companyMatch[1].trim() : "不明";

      const key = `${company}|${name}`;
      if (existingSet.has(key)) {
        console.log(`既に存在するデータです（${key}）。スキップ。`);
        return;
      }

      // データを書き込む最下段行（8行目以降）
      const currentLastRow = Math.max(ferretSheet.getLastRow(), startRow);
      const writeRow = currentLastRow + 1;

      ferretSheet.getRange(writeRow, 1, 1, 6).setValues([[
        "",
        company,
        "",
        name,
        "",
        material
      ]]);

      thread.markRead();
      thread.addLabel(label);
    });

    console.log(`tool ferret ホワイトペーパーDL：${ferretThreads.length} 件のスレッドを処理しました。`);
  } else {
    console.error("シート 'ferret資料請求' が存在しません。");
  }
}

function attachingPostedLabels() {
  // 今月頭のUNIXタイムスタンプを取得
  const now = new Date();
  const jstOffset = 9 * 60; // JSTはUTC+9時間（分単位）
  const startOfMonth = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), 1, 0, -jstOffset));  // 今月1日 00:00 JST を作成
// UNIXタイムスタンプ（秒単位）に変換
  const unixTimestamp = Math.floor(startOfMonth.getTime() / 1000);

  // スプレッドシート情報取得
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetTitle = spreadsheet.getName();
  const titleMatch = spreadsheetTitle.match(/^(\d{4})年(\d{1,2})月/);
  if (!titleMatch) {
    throw new Error("スプレッドシートのタイトルから年月を抽出できませんでした。");
  }
  const year = parseInt(titleMatch[1], 10);
  const month = parseInt(titleMatch[2], 10);
  const startDate = new Date(year, month - 1, 1);
  const endDate = new Date(year, month, 0, 23, 59, 59);
  const startTimestamp = Math.floor(startDate.getTime() / 1000);
  const endTimestamp = Math.floor(endDate.getTime() / 1000);
  // スプレッドシートは月ごとに新しく作成され、指定された年月以外のシートに自動でデータが記入されるのを防ぐため
  // 検索時間が範囲外の場合は処理を終了
  if (unixTimestamp < startTimestamp || unixTimestamp > endTimestamp) {
    console.log("UNIXタイムスタンプがスプレッドシートの年月の範囲外です。処理を終了します。");
    return;
  }
  // 共通ラベルの取得または作成
  let label = GmailApp.getUserLabelByName("転記済み");
  if (!label) {
    label = GmailApp.createLabel("転記済み");
    console.log("ラベル '転記済み' を作成しました。");
  }
  // ▼ DockpitFree 処理
  const dockpitSheet = spreadsheet.getSheetByName("DockpitFree申し込み");
  if (dockpitSheet) {
    const dockpitQuery = `subject:"フォームハンドラー: DockpitFree申込_アカウント発行" OR "フォームハンドラー: DockpitFree申込_既存ユーザー" after:${unixTimestamp}`;
    const dockpitThreads = GmailApp.search(dockpitQuery);
    dockpitThreads.forEach(thread => {
      if (thread.getLabels().some(l => l.getName() === "転記済み")) return;
        thread.markRead();
        thread.addLabel(label);
    });
    console.log(`DockpitFree：${dockpitThreads.length} 件のメールに転記済みラベルを貼り付けました。(転記済みラベルが貼り付けてあるものを含む)`);
  } else {
    console.error("シート 'DockpitFree申し込み' が存在しません。");
  }
  // ▼ ホワイトペーパーDL 処理
  const whitepaperSheet = spreadsheet.getSheetByName("ホワイトペーパーDL");
  if (whitepaperSheet) {
    const whitepaperQuery = `subject:"フォーム: 【マナミナ】ホワイトペーパー" after:${unixTimestamp}`;
    const whitepaperThreads = GmailApp.search(whitepaperQuery);

    whitepaperThreads.forEach(thread => {
      if (thread.getLabels().some(l => l.getName() === "転記済み")) return;
      thread.markRead();
      thread.addLabel(label);
    });
    console.log(`ホワイトペーパーDL：${whitepaperThreads.length} 件のメールに転記済みラベルを貼り付けました。(転記済みラベルが貼り付けてあるものを含む)`);
  } else {
    console.error("シート 'ホワイトペーパーDL' が存在しません。");
  }
}