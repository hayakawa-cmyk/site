// ================================================================
//  Pokemon Card Order — Google Apps Script Web App
//
//  【デプロイ手順】
//  1. https://script.google.com で新規プロジェクト作成
//  2. このコードを貼り付け
//  3. 右上「デプロイ」→「新しいデプロイ」
//  4. 種類: ウェブアプリ
//     ・次のユーザーとして実行: 自分
//     ・アクセスできるユーザー: 全員
//  5. 「デプロイ」→ 表示されたURLを index.html の GAS_ENDPOINT に貼り付け
// ================================================================

const RECIPIENT = 'hayakawa@wotaquest.com';

// ── エントリーポイント ──────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = createOrderSheet(data);
    sendOrderEmail(data, ss);
    return respond({ ok: true, sheetUrl: ss.getUrl() });
  } catch (err) {
    return respond({ ok: false, error: err.message });
  }
}

function doGet(e) {
  return respond({ ok: true, message: 'Pokemon Card Order API is running.' });
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── スプレッドシート作成 ────────────────────────────────────────
function createOrderSheet(data) {
  const now      = new Date();
  const dateStr  = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  const idStr    = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd-HHmmss');
  const safeName = (data.full_name || 'Unknown')
    .replace(/[\/\\:*?"<>|]/g, '_')
    .substring(0, 30);

  const ss    = SpreadsheetApp.create('Order_' + idStr + '_' + safeName);
  const sheet = ss.getActiveSheet();
  sheet.setName('Order');

  // ── データ行を組み立てる ──
  const rows = [];

  // 注文情報
  rows.push(['ORDER INFORMATION', '', '', '']);
  rows.push(['Order ID',             'ORD-' + idStr, '', '']);
  rows.push(['Order Date (JST)',     dateStr,          '', '']);
  rows.push(['Price Reference Date', data.price_date || '', '', '']);
  rows.push(['', '', '', '']);

  const custHeaderRow = rows.length + 1;
  rows.push(['CUSTOMER INFORMATION', '', '', '']);
  rows.push(['Full Name',          data.full_name     || '', '', '']);
  rows.push(['Email',              data.email         || '', '', '']);
  rows.push(['Phone',              data.phone         || '', '', '']);
  rows.push(['Address Line 1',     data.address_line1 || '', '', '']);
  rows.push(['Address Line 2',     data.address_line2 || '', '', '']);
  rows.push(['City',               data.city          || '', '', '']);
  rows.push(['State / Province',   data.state         || '', '', '']);
  rows.push(['ZIP / Postal Code',  data.zip           || '', '', '']);
  rows.push(['Country',            data.country       || '', '', '']);
  rows.push(['', '', '', '']);

  const itemHeaderRow = rows.length + 1;
  rows.push(['ORDER ITEMS', '', '', '']);

  const colHeaderRow = rows.length + 1;
  rows.push(['Product', 'Unit Price (JPY)', 'Qty (BOX)', 'Subtotal (JPY)']);

  // 商品行をパース: "商品名 x2 (¥26,950)"
  const items = (data.order_items || '').split(' | ');
  items.forEach(function(item) {
    var m = item.match(/^(.+?) x(\d+) \((.+)\)$/);
    if (m) {
      rows.push([m[1].trim(), '', Number(m[2]), m[3]]);
    }
  });

  rows.push(['', '', '', '']);
  const totalRow = rows.length + 1;
  rows.push(['', '', 'TOTAL', data.total_jpy || '']);

  // ── 一括書き込み ──
  sheet.getRange(1, 1, rows.length, 4).setValues(rows);

  // ── スタイリング ──
  var sectionStyle = {
    background: '#1b263b',
    fontColor:  '#f4a261',
    bold: true,
    fontSize: 11
  };

  // セクションヘッダー（行番号は1始まり）
  [1, custHeaderRow, itemHeaderRow].forEach(function(r) {
    sheet.getRange(r, 1, 1, 4)
      .merge()
      .setBackground(sectionStyle.background)
      .setFontColor(sectionStyle.fontColor)
      .setFontWeight('bold')
      .setFontSize(sectionStyle.fontSize)
      .setHorizontalAlignment('left');
  });

  // 商品一覧の列ヘッダー
  sheet.getRange(colHeaderRow, 1, 1, 4)
    .setBackground('#243447')
    .setFontColor('#e8e8e8')
    .setFontWeight('bold');

  // 合計行
  sheet.getRange(totalRow, 3, 1, 2).setFontWeight('bold');

  // 列幅
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 160);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 130);

  return ss;
}

// ── メール送信 ──────────────────────────────────────────────────
function sendOrderEmail(data, ss) {
  // SpreadsheetをPDFに変換して添付
  var pdfBlob = DriveApp.getFileById(ss.getId())
    .getAs(MimeType.PDF);
  pdfBlob.setName(
    'Order_' + (data.full_name || 'Unknown').replace(/[\/\\:*?"<>|]/g, '_') + '.pdf'
  );

  var address = [
    data.address_line1,
    data.address_line2,
    data.city,
    data.state,
    data.zip,
    data.country
  ].filter(Boolean).join(', ');

  var itemLines = (data.order_items || '')
    .split(' | ')
    .map(function(s) { return '  • ' + s; })
    .join('\n');

  var body = [
    '新しい注文が届きました。',
    '',
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
    '注文者情報',
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
    '氏名:   ' + (data.full_name || ''),
    'メール: ' + (data.email     || ''),
    '電話:   ' + (data.phone     || ''),
    '住所:   ' + address,
    '',
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
    '注文商品',
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
    itemLines,
    '',
    '合計金額:   ' + (data.total_jpy  || ''),
    '価格参照日: ' + (data.price_date || ''),
    '',
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
    'スプレッドシート: ' + ss.getUrl()
  ].join('\n');

  GmailApp.sendEmail(
    RECIPIENT,
    '[ポケカ注文] ' + (data.full_name || 'Unknown') + ' — ' + (data.total_jpy || ''),
    body,
    {
      attachments: [pdfBlob],
      replyTo:     data.email || '',
      name:        'Pokemon Card Order System'
    }
  );
}
