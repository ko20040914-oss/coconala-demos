/** スプレッドシート側の「フォーム送信時」インストール型トリガー用。 */
const CONFIG = Object.freeze({
  // 以下はプレースホルダーです。ご自身の設定に置き換えてください。
  SPREADSHEET_ID: 'REPLACE_WITH_YOUR_SPREADSHEET_ID',
  SHEET_NAME: '受付台帳',
  OWNER_EMAIL: 'owner@example.com',
  SEND_EMAILS: false // デモモード（送信なし）。メールテスト・実運用では自分の非公開コピーでのみ true にする。
});
const HEADERS = ['受付日時', '名前', '希望日時', 'コース', '連絡先メール', 'ステータス'];

function onFormSubmit(e) {
  if (!e || !e.namedValues) {
    throw new Error('手動実行はできません。スプレッドシートのフォーム送信時トリガーを使用してください。');
  }
  const fields = HEADERS.slice(1, 5);
  const values = fields.map(function (key) {
    const answer = e.namedValues[key];
    return answer && answer.length ? String(answer[0]).trim() : '';
  });
  const notes = [];
  fields.forEach(function (key, i) {
    if (!values[i]) notes.push('注意：' + key + '未入力');
  });
  const email = values[3];
  if (email && !isSingleEmail_(email)) notes.push('注意：連絡先メール形式不正');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const book = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = book.getSheetByName(CONFIG.SHEET_NAME) || book.insertSheet(CONFIG.SHEET_NAME);
    if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
    const actual = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
    if (actual.join('\t') !== HEADERS.join('\t')) {
      throw new Error('受付台帳のヘッダーが一致しません。READMEの列順を確認してください。');
    }
    const row = sheet.getLastRow() + 1;
    // 数式として解釈されないよう、利用者入力を文字列として保存。
    sheet.getRange(row, 1, 1, 6).setValues([[
      new Date(), ...values.map(safeCell_), notes.concat(['受付済み・通知処理中']).join(' / ')
    ]]);
    SpreadsheetApp.flush(); // 通知前に受付データを保存
    const detail = fields.map(function (key, i) {
      return key + '：' + (values[i] || '（未入力）');
    }).join('\n');
    sendMail_(CONFIG.OWNER_EMAIL, '【体験レッスン】新しい申込を受け付けました',
      detail + '\n\n' + notes.join('\n'), 'オーナー通知', notes);
    if (isSingleEmail_(email)) {
      sendMail_(email, '【体験レッスン】お申込を受け付けました',
        (values[0] || 'お申込者') + ' 様\n\nお申込ありがとうございます。\n' + detail +
        '\n\nこれは受付の確認です。予約確定ではありません。店舗からの連絡をお待ちください。',
        '確認メール', notes);
    } else {
      notes.push('確認メール未送信：連絡先メール未入力または形式不正');
    }
    sheet.getRange(row, 6).setValue(notes.join(' / '));
  } finally {
    lock.releaseLock();
  }
}

function isSingleEmail_(value) {
  return /^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?)+$/.test(value);
}

function safeCell_(value) {
  return /^[=+@-]/.test(value) ? "'" + value : value;
}

function sendMail_(to, subject, body, label, notes) {
  if (!CONFIG.SEND_EMAILS) {
    notes.push(label + '未送信：デモモード');
    return;
  }
  if (!isSingleEmail_(to) || /@(?:.*\.)?example\.(com|net|org)$/i.test(to)) {
    notes.push(label + '未送信：宛先未設定・形式不正・ダミー宛先');
    return;
  }
  try {
    MailApp.sendEmail({ to: to, subject: subject, body: body });
    notes.push(label + '送信済み');
  } catch (error) {
    // 例外の詳細には個人情報が含まれ得るため、台帳・ログへ出力しない。
    notes.push('注意：' + label + '送信失敗（権限・送信上限等を確認）');
  }
}
